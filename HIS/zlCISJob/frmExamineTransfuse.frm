VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExamineTransfuse 
   Caption         =   "输血审核管理"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14805
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   14805
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   12240
      TabIndex        =   38
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "住院"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   40
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "门诊"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         Top             =   -10
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "使用场合"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   60
         Width           =   4000
      End
   End
   Begin VB.Frame fraPati 
      Caption         =   "病人信息"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   16
      Top             =   600
      Width           =   11295
      Begin VB.ComboBox cbo过敏 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   697
         Width           =   4815
      End
      Begin VB.PictureBox picInShow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   8445
         TabIndex        =   17
         Top             =   360
         Width           =   8450
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   18
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   19
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblCaption 
            Caption         =   "床号："
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   26
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "护理等级："
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   25
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblCaption 
            Caption         =   "病况："
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   24
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "体重："
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "入院时间："
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   27
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   29
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblCaption 
         Caption         =   "诊断："
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "过敏药物："
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         Caption         =   "年龄："
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "性别："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox picUnAudited 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4440
      ScaleHeight     =   5895
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   2640
      Width           =   9735
      Begin VB.PictureBox picDateY 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9615
         TabIndex        =   42
         Top             =   120
         Width           =   9615
         Begin VB.CommandButton cmdFindY 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   44
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboDateY 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   30
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   45
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   46
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin VB.Label lblPri 
            Height          =   300
            Index           =   0
            Left            =   7080
            TabIndex        =   49
            Top             =   120
            Width           =   8000
         End
         Begin VB.Label lblDateY 
            Caption         =   "开嘱时间"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   75
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2460
            TabIndex        =   47
            Top             =   90
            Width           =   1890
         End
      End
      Begin VB.PictureBox picDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   9375
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   30
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   11
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   12
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin VB.Label lblPri 
            Height          =   300
            Index           =   1
            Left            =   7200
            TabIndex        =   52
            Top             =   120
            Width           =   7995
         End
         Begin VB.Label lblDate 
            Caption         =   "签发时间"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2460
            TabIndex        =   13
            Top             =   90
            Width           =   1890
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAudit 
         Height          =   4860
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   8505
         _cx             =   15002
         _cy             =   8572
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   41
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin VB.Frame fraDoctor 
      Caption         =   "医生"
      ForeColor       =   &H000040C0&
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   4020
         Left            =   105
         TabIndex        =   1
         Top             =   1500
         Width           =   3330
         _Version        =   589884
         _ExtentX        =   5874
         _ExtentY        =   7091
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picRule 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   3375
         TabIndex        =   50
         Top             =   5520
         Width           =   3375
         Begin VB.Label lbl 
            Caption         =   "血液审核规定"
            Height          =   3135
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.CheckBox chkIsShowAll 
         Caption         =   "只显示有申请的医生"
         Height          =   180
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   3
         Top             =   788
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   315
         TabIndex        =   6
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&D)"
         Height          =   180
         Left            =   315
         TabIndex        =   5
         Top             =   420
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7335
      Left            =   3720
      TabIndex        =   36
      Top             =   1845
      Width           =   11355
      _Version        =   589884
      _ExtentX        =   20029
      _ExtentY        =   12938
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   10575
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21034
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
   Begin MSComctlLib.ImageList img16 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":0000
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":005E
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":00BC
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":011A
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgAdvice 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":0234
            Key             =   "签名"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmExamineTransfuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mobjBar As CommandBar
Private mobjPopup As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean
Private mintAuditPrivs As Integer '当前用户审核权限:0:不具有任何权限，不能审核签发任何量，1：800ml以下可审核签发，2：1600以下均可以审核签发，3：可审核、签发所有。
Private mstrButPri As String '拒绝审核、审核、取消审核、签发、取消签发 按钮是否可用（1表示可用，0表示不可用；例 11010）
Private mbln启用输血三级审核 As Boolean
Private mlngFindNum As Long
Private mstrChangeRows As String   '记录修改的行
Private mstr签名IDs As String      '取消审核的时候记录一个病人处理过的签名ID
'手术审核暂时不启用签名功能，所以判断加了 And 1 = 0
Private mblnTmp As Boolean
Private mrsDefine As ADODB.Recordset
Private mint场合 As Integer
Private mclsMipModule As zl9ComLib.clsMipModule
Private Enum Enum_Dor
    COL_人员ID = 0
    col_姓名 = 1
    COL_专业技术职务 = 2
    COL_管理职务 = 3
    COL_拼音简码 = 4
    COL_五笔简码 = 5
    COL_所属部门 = 6
    COL_所属部门ID = 7
End Enum

Private Enum Enum_Type  '指定要更新申请到的状态
    t_待审核 = 1
    t_待签发 = 2
    t_已签发 = 3
    t_已拒绝 = 4
End Enum

Private Enum Enum_Advice_New
    col_选择 = 0
    COL_审核说明 = 1
    COL_审核时间 = 2
    COL_病人姓名 = 3
    COL_医嘱内容 = 4
    col_期效 = 5
'用简洁模式，所以总量和单量隐藏起来，和医嘱内容合并
    COL_总量 = 6
    col_输血时间 = 7
    COL_开始时间 = 8
    col_输血总量 = 9
    col_24h输血量 = 10
    col_审核状态说明 = 11
'隐藏列
    col_医嘱ID = 12
    col_相关ID = 13
    col_性别 = 14
    col_年龄 = 15
    COL_体重 = 16
    COL_入院时间 = 17
    col_床号 = 18
    COL_病况 = 19
    COL_护理等级 = 20
    col_病人Id = 21
    col_主页ID = 22
    COL_组ID = 23
    COL_诊疗类别 = 24
    COL_病人来源 = 25
    COL_签名id = 26
    COL_医嘱状态 = 27
    col_挂号单 = 28
    col_审核状态 = 29
    
    COL_门诊号 = 30
    col_住院号 = 31
    COL_当前床号 = 32
    COL_开嘱医生 = 33
    COL_开嘱时间 = 34
    COL_开嘱科室ID = 35
    COL_出院科室ID = 36
    COL_当前病区ID = 37
    COL_主项目ID = 38
    
    col_配血医嘱 = 39 '1－用血医嘱，0－备血医嘱
    COL_申请序号 = 40
    COL_操作类型 = 41
End Enum

Private Enum enum_Info
    info_入院时间 = 0
    info_性别 = 1
    info_年龄 = 2
    info_床号 = 3
    info_护理等级 = 4
    info_病况 = 5
    info_诊断 = 6
    info_体重 = 7
End Enum

Public Function ShowMe(frmParent As Object, ByVal int场合 As Integer, Optional ByRef ojbMip As Object) As Boolean
'参数：mint场合=1门诊，2住院
    On Error Resume Next
    
    mint场合 = int场合
    If Not ojbMip Is Nothing Then Set mclsMipModule = ojbMip
    Call frmExamineTransfuse.Show(0, frmParent)
End Function

Private Sub cboDept_Click()
    Call LoadDoc
End Sub
Private Sub cmdFindY_Click()
    Call LoadAdvice(False)
End Sub

Private Sub LoadDoc()
'加载权限比操作员低的医生
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    Dim datBegint As Date
    Dim datEnd As Date
    
    If cboDept.ListIndex = -1 Then Exit Sub
    On Error GoTo errH
    Screen.MousePointer = 11
    stbThis.Panels(2).Text = "请选择一位开嘱医生。"
    rptDoc.Records.DeleteAll
    vsAudit.Rows = 1: vsAudit.AddItem ""
    
    '此处排除用血医嘱
    If tbcSub.Selected.Tag = "待审核" Or tbcSub.Selected.Tag = "待签发" Then
        datBegint = CDate(dtpTimeY(0).Value)
        datEnd = CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct F.开嘱医生 From 诊疗项目目录 K,病人医嘱记录 H ,病人医嘱记录 F " & _
            " Where (K.操作类型='8' and nvl(K.执行分类,0)=0 or K.操作类型='9') And K.ID=H.诊疗项目ID And H.相关ID=F.id And " & _
             IIf(mbln启用输血三级审核 = False, " f.审核状态 in (1,7) ", IIf(tbcSub.Selected.Tag = "待审核", " f.审核状态=1 ", " f.审核状态=7 ")) & " and f.医嘱状态=1 And F.开嘱时间 Between [4] And [5] And f.病人来源=[3] And f.诊疗类别 ='K') F"
    Else
        datBegint = CDate(dtpTime(0).Value)
        datEnd = CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct f.开嘱医生 From 诊疗项目目录 K,病人医嘱记录 H ,病人医嘱记录 F,病人医嘱状态 G " & _
            " Where (K.操作类型='8' and nvl(K.执行分类,0)=0  or K.操作类型='9') And K.ID=H.诊疗项目ID And H.相关ID=F.id And F.id=g.医嘱id and G.操作类型 in (11,12,14)" & _
            " And G.操作时间 Between [4] And [5] And f.病人来源=[3] And f.诊疗类别 ='K') F"
    End If
    
    strSQL = "Select DISTINCT a.Id, A.性别" & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", ",b.部门ID,e.名称 as 所属部门") & ",a.姓名,a.专业技术职务,a.管理职务, Upper(zlSpellCode(a.姓名)) As 拼音简码, Upper(Zlwbcode(a.姓名)) As 五笔简码" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 D,部门表 E" & IIf(chkIsShowAll.Value, strTmp, "") & vbNewLine & _
            "Where a.Id = b.人员id And e.ID=b.部门ID And d.人员id = a.Id  And d.人员性质 = '医生' And " & vbNewLine & _
            "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & vbNewLine & _
            "   " & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.部门id=[2]") & _
            IIf(chkIsShowAll.Value, " And  f.开嘱医生 = a.姓名 ", "")
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)), IIf(optOccasion(0).Value, 2, 1), datBegint, datEnd)
    
    With rptDoc
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!性别 & "" = "女", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!专业技术职务 & "")
            Set objItem = objRecord.AddItem(rsTmp!管理职务 & "")
            Set objItem = objRecord.AddItem(rsTmp!拼音简码 & "")
            Set objItem = objRecord.AddItem(rsTmp!五笔简码 & "")
            If Val(cboDept.ItemData(cboDept.ListIndex)) <> -1 Then
                Set objItem = objRecord.AddItem(rsTmp!所属部门 & "")
                Set objItem = objRecord.AddItem(rsTmp!部门ID & "")
            End If
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    mlngFindNum = 0
    Screen.MousePointer = 0
    Call vsAudit_KeyPress(vbKeyBack)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptDoc.Visible = False Then Exit Sub
    If rptDoc.Records.Count > 0 Then
        If rptDoc.SelectedRows.Count = 0 Then Exit Sub
        strSubhead = rptDoc.SelectedRows(0).Record(col_姓名).Value & "输血审核清单"
    Else
        Exit Sub
    End If
    
    '调用打印部件处理
    Set objPrint.Body = Me.vsAudit
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Cancle()
'功能：取消保存
    Dim i As Long
    With vsAudit
        If MsgBox("本次修改的内容未保存，是否继续？", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            If tbcSub.Selected.Tag = "已审核" Then
                Call LoadAdvice(True)
            Else
                Call LoadAdvice
            End If
            mblnIsUpdate = False
            mstrChangeRows = ""
        End If
    End With
End Sub

Private Sub SaveAudit()
'功能：保存审核信息
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strDate As String
    
    With vsAudit
        If .EditText <> "" Then .TextMatrix(.Row, .Col) = .EditText
        If Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value) = UserInfo.ID Then
            MsgBox "不能审核自己申请的输血医嘱", vbInformation, Me.Caption
            Exit Sub
        End If
        '判断是否为新版血液审核
'        If mbln启用输血三级审核 Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_选择) = flexChecked Then
                    Exit For
                End If
                If i = .Rows - 1 Then
                MsgBox "您还未选择需要处理的项目信息，请选择！", vbInformation, Me.Caption
                Exit Sub
                End If
            Next
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_选择) = flexChecked And .TextMatrix(i, COL_审核说明) <> "" Then
                    If MsgBox("确定后会删除说明，继续？", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
                    
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            For i = 1 To .Rows - 1
                '一个病人调用一次
                If RowIn同一病人(i, lngBegin, lngEnd, vsAudit) Then
                    Call SaveAuditOnePati(lngBegin, lngEnd, strDate)
                    i = lngEnd
                Else
                    Call SaveAuditOnePati(i, i, strDate)
                End If
            Next
            Call LoadAdvice
        mstrChangeRows = ""
        mblnIsUpdate = False
        End With
End Sub

Private Sub SaveAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, ByVal strDate As String)
'功能：保存审核信息
'参数：从第几行开始，到第几行结束（同一个病人）
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String
    Dim strSource As String, strSign As String
    Dim lng签名ID As Long, lng证书ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim strSignSQL As String
    Dim int状态 As Integer
    Dim lngMsgRow As Long, lngBlood As Long
    Dim rsTmp As ADODB.Recordset
    Dim intQuestion As Integer, intAudit As Integer
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .TextMatrix(i, col_医嘱ID) = "" Then Exit Sub
            If .Cell(flexcpChecked, i, col_选择) = flexChecked Then
                If tbcSub.Selected.Tag = "待签发" Then
                    int状态 = 3
                ElseIf tbcSub.Selected.Tag = "待审核" Then
                    If mbln启用输血三级审核 Then
                        Select Case Val(.TextMatrix(i, col_24h输血量))
                            Case Is >= 1600
                                intAudit = 3
                            Case Is >= 800
                                intAudit = 2
                            Case Else
                                intAudit = 1
                        End Select
                        If mintAuditPrivs >= intAudit And intAudit > 1 Then '只有800以上才会进入此阶段
                            If intQuestion = 0 Then
                                If MsgBox("您可以直接完成" & .TextMatrix(i, COL_病人姓名) & "审核与签发，点击“是”直接完成，点击“否”仅进行审核。", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                                    intQuestion = 1
                                Else
                                    intQuestion = 2
                                End If
                            End If
                            If intQuestion = 1 Then
                                int状态 = 3
                            ElseIf intQuestion = 2 Then
                                int状态 = 6
                            End If
                        Else
                            If intAudit > 1 Then
                                int状态 = 6
                            Else
                                int状态 = 3
                            End If
                        End If
                    Else    '不启用三级审核直接通过。
                        int状态 = 3
                    End If
                End If
                If int状态 = 3 Then lngMsgRow = i  '新流程签发完成医嘱就可以发送
                '如果未启用血库系统，状态3为血库待接收，更改为审核通过（状态1）
                If int状态 = 3 And Not gbln血库系统 Then int状态 = 1
                strSQL = Val(.TextMatrix(i, col_医嘱ID)) & "|" & "Zl_医嘱审核管理_Audit(" & Val(.TextMatrix(i, col_医嘱ID)) & "," & int状态 & "," & _
                        "'" & UserInfo.姓名 & "'," & strDate & ",''"
                colsql.Add strSQL, "C" & colsql.Count + 1
            End If
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colsql.Count
        strSQL = Mid(colsql("C" & i), InStr(colsql("C" & i), "|") + 1) & ",2)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    If strSignSQL <> "" Then
        Call zlDatabase.ExecuteProcedure(strSignSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If mint场合 = 2 Then
        '触发医嘱新下达消息/输血配血申请消息
        With vsAudit
            If lngMsgRow <> 0 Then
                strSQL = "select a.操作类型 from 诊疗项目目录 a where a.id=[1]"
                If Not mbln启用输血三级审核 Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngMsgRow, COL_主项目ID)))
                    Call ZLHIS_CIS_001(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_病人Id)), .TextMatrix(lngMsgRow, COL_病人姓名), .TextMatrix(lngMsgRow, col_住院号), "", Val(.TextMatrix(lngMsgRow, COL_病人来源)), Val(.TextMatrix(lngMsgRow, col_主页ID)), _
                        Val(.TextMatrix(lngMsgRow, COL_当前病区ID)), "", Val(.TextMatrix(lngMsgRow, COL_出院科室ID)), "", "", .TextMatrix(lngMsgRow, COL_当前床号), Val(.TextMatrix(lngMsgRow, col_医嘱ID)), 0, 1, "K", rsTmp!操作类型 & "", _
                        .TextMatrix(lngMsgRow, COL_开嘱医生), .TextMatrix(lngMsgRow, COL_开嘱时间), .TextMatrix(lngMsgRow, COL_开嘱科室ID), "")
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngMsgRow, COL_主项目ID)))
                    Call ZLHIS_CIS_001(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_病人Id)), .TextMatrix(lngMsgRow, COL_病人姓名), .TextMatrix(lngMsgRow, col_住院号), "", Val(.TextMatrix(lngMsgRow, COL_病人来源)), Val(.TextMatrix(lngMsgRow, col_主页ID)), _
                        Val(.TextMatrix(lngMsgRow, COL_当前病区ID)), "", Val(.TextMatrix(lngMsgRow, COL_出院科室ID)), "", "", .TextMatrix(lngMsgRow, COL_当前床号), Val(.TextMatrix(lngMsgRow, col_医嘱ID)), 0, 1, "K", rsTmp!操作类型 & "", _
                        .TextMatrix(lngMsgRow, COL_开嘱医生), .TextMatrix(lngMsgRow, COL_开嘱时间), .TextMatrix(lngMsgRow, COL_开嘱科室ID), "")
                End If
            End If
            If lngBlood <> 0 Then
                lngMsgRow = lngBlood
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        Call ZLHIS_CIS_031(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_病人Id)), .TextMatrix(lngMsgRow, COL_病人姓名), .TextMatrix(lngMsgRow, col_住院号), "", Val(.TextMatrix(lngMsgRow, COL_病人来源)), Val(.TextMatrix(lngMsgRow, col_主页ID)), _
                            Val(.TextMatrix(lngMsgRow, COL_当前病区ID)), "", Val(.TextMatrix(lngMsgRow, COL_出院科室ID)), "", "", .TextMatrix(lngMsgRow, COL_当前床号), Val(.TextMatrix(lngMsgRow, col_医嘱ID)), _
                            .TextMatrix(lngMsgRow, COL_开嘱医生), .TextMatrix(lngMsgRow, COL_开嘱时间), .TextMatrix(lngMsgRow, COL_开嘱科室ID), "")
                    End If
                End If
            End If
        End With
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTime(0).MaxDate = curDate
    dtpTime(1).MaxDate = curDate
    cmdFind.Visible = False
    
    Select Case cboTime.ListIndex
    Case 0 '今日
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '最近二天
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '最近三天
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '最近一周
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '最近一月
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 '指  定
        If Me.Visible Then dtpTime(0).SetFocus
        cmdFind.Visible = True
    End Select
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 And Me.Visible Then
        If chkIsShowAll.Value = 1 Then
            Call LoadDoc
        Else
            Call LoadAdvice(True)
        End If
    End If
End Sub

Private Sub CancleAudit()
'取消审核
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnIsCheck As Boolean
    
    With vsAudit
        If Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value) = UserInfo.ID Then
            MsgBox "不能取消自己申请的输血医嘱", vbInformation, Me.Caption
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = vbChecked Then Exit For
            If i = .Rows - 1 Then
                Call MsgBox("未勾选要取消的项目，请查证！", vbInformation, Me.Caption)
                Exit Sub
            End If
        Next
        '判断是否有勾选的，有勾选就以勾选为准
        For i = i To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = vbChecked Then
                If RowIn同一病人(i, lngBegin, lngEnd, vsAudit) Then
                    Call CancleAuditOnePati(lngBegin, lngEnd)
                    i = lngEnd
                Else
                    Call CancleAuditOnePati(i, i)
                End If
            End If
        Next
        Call LoadAdvice(True)
    End With
End Sub

Private Sub CancleAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, Optional ByVal blnIsNoCheck As Boolean, _
        Optional ByVal lng签名ID_IN As Long, Optional ByVal blnIsRecursive As Boolean = True)
'功能：取消审核
'参数：lngBegin从第几行开始，lngEnd到第几行结束（同一个病人）
'     blnIsNoCheck=没有勾选则已选中行为准取消审核
'     lng签名ID_IN：用于递归调用，如果第一次循环中发现有签名ID<>0，则递归调用本函数，并把这里的签名ID传入，
'    加入到字符串mstr签名IDs里，第二次进来则处理签名ID的医嘱,如果再发现与传入的签名ID不一样，并且又不在字符串mstr签名IDs中，则为新的，则再递归调用。
'    blnIsRecursive:是否递归，默认为要递归
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String, blnTrans As Boolean
    Dim strSource As String, strSign As String, strDate As String
    Dim lng证书ID As Long, lng签名ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim rsSQL As New ADODB.Recordset, strExp As String, strTmp As String
    Dim arrIDs(4) As String, arrExp(4) As String
    Dim rsChk As Recordset
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .Cell(flexcpChecked, i, 0) = vbChecked Then
                If tbcSub.Selected.Tag = "待签发" And .TextMatrix(i, COL_操作类型) = 18 Then '存在审核过程，回退至待审核
                    arrIDs(t_待审核) = arrIDs(t_待审核) & IIf(arrIDs(t_待审核) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                    arrExp(t_待审核) = arrExp(t_待审核) & IIf(arrExp(t_待审核) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                ElseIf tbcSub.Selected.Tag = "待签发" And Nvl(.TextMatrix(i, COL_操作类型), 0) = 0 Then '不存在审核过程，拒绝该申请
                    If MsgBox("您若取消 " & .TextMatrix(i, COL_病人姓名) & "：" & .TextMatrix(i, COL_医嘱内容) & "的申请，会直接拒绝该审核申请，是否确定？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        arrIDs(t_已拒绝) = arrIDs(t_已拒绝) & IIf(arrIDs(t_已拒绝) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                        arrExp(t_已拒绝) = arrExp(t_已拒绝) & IIf(arrExp(t_已拒绝) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                    End If
                ElseIf tbcSub.Selected.Tag = "待审核" Then  '拒绝审核申请
                    arrIDs(t_已拒绝) = arrIDs(t_已拒绝) & IIf(arrIDs(t_已拒绝) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                    arrExp(t_已拒绝) = arrExp(t_已拒绝) & IIf(arrExp(t_已拒绝) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                ElseIf tbcSub.Selected.Tag = "已审核" Then  '兼容处理，非三级审核回退至待审核或者取消拒绝
                    If .TextMatrix(i, col_审核状态) = 3 Then
                        If MsgBox("您可审核 " & .TextMatrix(i, COL_病人姓名) & "：" & .TextMatrix(i, COL_医嘱内容) & ",是否直接审核？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                            arrIDs(t_待审核) = arrIDs(t_待审核) & IIf(arrIDs(t_待审核) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                            arrExp(t_待审核) = arrExp(t_待审核) & IIf(arrExp(t_待审核) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                        Else
                            arrIDs(t_已签发) = arrIDs(t_已签发) & IIf(arrIDs(t_已签发) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                            arrExp(t_已签发) = arrExp(t_已签发) & IIf(arrExp(t_已签发) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                        End If
                    ElseIf .TextMatrix(i, col_审核状态) = 4 Or .TextMatrix(i, col_审核状态) = 2 Then
                        arrIDs(t_待审核) = arrIDs(t_待审核) & IIf(arrIDs(t_待审核) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                        arrExp(t_待审核) = arrExp(t_待审核) & IIf(arrExp(t_待审核) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                    End If
                ElseIf tbcSub.Selected.Tag = "已签发" And (.TextMatrix(i, col_审核状态) = 4 Or .TextMatrix(i, col_审核状态) = 2) Then
                    strSQL = "select 1 from 病人医嘱状态 where 操作类型 = 17 and  医嘱id = [1] and rownum < 2"
                    Set rsChk = zlDatabase.OpenSQLRecord(strSQL, "查询该医嘱是否签发过", .TextMatrix(i, col_医嘱ID))
                    If rsChk.BOF Then '无数据，无法判断，退回至待审核
                        arrIDs(t_待审核) = arrIDs(t_待审核) & IIf(arrIDs(t_待审核) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                        arrExp(t_待审核) = arrExp(t_待审核) & IIf(arrExp(t_待审核) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                    Else
                        arrIDs(t_待签发) = arrIDs(t_待签发) & IIf(arrIDs(t_待签发) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                        arrExp(t_待签发) = arrExp(t_待签发) & IIf(arrExp(t_待签发) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                    End If
                ElseIf tbcSub.Selected.Tag = "已签发" And .TextMatrix(i, col_审核状态) = 3 Then    '已签发页面取消拒绝
                    If MsgBox("您可签发 " & .TextMatrix(i, COL_病人姓名) & "：" & .TextMatrix(i, COL_医嘱内容) & ",是否直接签发？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        arrIDs(t_已签发) = arrIDs(t_已签发) & IIf(arrIDs(t_已签发) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                        arrExp(t_已签发) = arrExp(t_已签发) & IIf(arrExp(t_已签发) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                    Else
                        strSQL = "select 1 from 病人医嘱状态 where 操作类型 = 17 and  医嘱id = [1] and rownum < 2"
                        Set rsChk = zlDatabase.OpenSQLRecord(strSQL, "查询该医嘱是否签发过", .TextMatrix(i, col_医嘱ID))
                        If rsChk.BOF Then '无数据，无法判断，退回至待审核
                            arrIDs(t_待审核) = arrIDs(t_待审核) & IIf(arrIDs(t_待审核) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                            arrExp(t_待审核) = arrExp(t_待审核) & IIf(arrExp(t_待审核) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                        Else
                            arrIDs(t_待签发) = arrIDs(t_待签发) & IIf(arrIDs(t_待签发) = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                            arrExp(t_待签发) = arrExp(t_待签发) & IIf(arrExp(t_待签发) = "", "", ",") & .TextMatrix(i, COL_审核说明)
                        End If
                    End If
                End If
            End If


        Next
    End With
    Call CancleAuditOnePatiChild(arrIDs(t_待审核), strDate, arrExp(t_待审核), 1)
    Call CancleAuditOnePatiChild(arrIDs(t_待签发), strDate, arrExp(t_待签发), 7)
    Call CancleAuditOnePatiChild(arrIDs(t_已签发), strDate, arrExp(t_已签发), IIf(gbln血库系统, 4, 2))
    Call CancleAuditOnePatiChild(arrIDs(t_已拒绝), strDate, arrExp(t_已拒绝), 3)
    mstr签名IDs = "0"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CancleAuditOnePatiChild(ByVal strIDs As String, ByVal strDate As String, ByVal strExp As String, ByVal intType As Integer)
    Dim strTmp As String, strSQL As String
    Dim rsSQL As New ADODB.Recordset
    Dim i As Long, blnTrans As Boolean
    
    On Error GoTo errH
    If strIDs <> "" Then
        Call SQLRecord(rsSQL)
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = "Zl_医嘱审核管理_Cancel('" & strIDs & "',2," & intType & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
        For i = 0 To UBound(Split(strIDs, ","))
            If strExp <> "" Then
                If UBound(Split(strExp, ",")) <= i Then
                    strTmp = Split(strExp, ",")(i)
                Else
                    strTmp = ""
                End If
            End If
            strSQL = "Zl_医嘱审核管理_Update('" & Split(strIDs, ",")(i) & "'," & strDate & ",'" & strTmp & "',2,'" & UserInfo.姓名 & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
        Next
        If Not SQLRecordExecute(rsSQL, Me.Caption) Then blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext <> conMenu_Edit_Audit Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
     Case conMenu_Edit_Audit '右键快速审核
        For i = vsAudit.FixedRows To vsAudit.Rows - 1
            If vsAudit.Cell(flexcpChecked, i, col_选择) = vbChecked Then Exit For
        Next
        If i < vsAudit.Rows And vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_选择) = vbUnchecked Then
            If MsgBox("本次审核操作只保存已勾选的医嘱，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
        ElseIf i = vsAudit.Rows And vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_选择) = vbUnchecked Then
            vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_选择) = vbChecked
        End If
        Call SaveAudit
    Case conMenu_Edit_AdviceUnAudit '取消审核
        Call CancleAudit
        Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_UnUse, conMenu_Edit_StopAudit ' 拒绝审核 取消签发
        Call CancleAudit
        Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_MediAudit, conMenu_Edit_Send '审核,签发
            Call SaveAudit
            Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_ApplyView '查看输血申请单
        If vsAudit.Row <= 0 Then Exit Sub
        If Val(vsAudit.TextMatrix(vsAudit.Row, col_医嘱ID)) = 0 Then Exit Sub
        Call gobjKernel.ShowBloodApply(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_医嘱ID)))
    Case conMenu_Tool_Archive '电子病案查阅
        If vsAudit.Row <= 0 Then Exit Sub
        If Val(vsAudit.TextMatrix(vsAudit.Row, col_医嘱ID)) = 0 Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_病人Id)), Val(vsAudit.TextMatrix(vsAudit.Row, col_主页ID)))
    Case conMenu_View_Find '查找
        txtFind.SetFocus '有时需要定位一下
        If txtFind.Text <> "" Then
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_FindNext '查找下一个
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
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
    Case conMenu_View_Refresh '刷新
        If tbcSub.Selected.Tag = "待审核" Or tbcSub.Selected.Tag = "待签发" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
        If mbln启用输血三级审核 Then Call vsAudit_KeyPress(vbKeyBack)
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptDoc.SelectedRows.Count = 0 Or vsAudit.Row <= 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "执行科室=" & cboDept.ItemData(cboDept.ListIndex))
                Else
                    With vsAudit
                        If .TextMatrix(.Row, COL_病人来源) = "2" Then
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "执行科室=" & cboDept.ItemData(cboDept.ListIndex), "审核人=" & rptDoc.SelectedRows(0).Record(col_姓名).Value, _
                                "病人ID=" & .TextMatrix(.Row, col_病人Id), "主页ID=" & .TextMatrix(.Row, col_主页ID), "医嘱ID=" & .TextMatrix(.Row, col_医嘱ID))
                        Else
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "执行科室=" & cboDept.ItemData(cboDept.ListIndex), "审核人=" & rptDoc.SelectedRows(0).Record(col_姓名).Value, _
                                "病人ID=" & .TextMatrix(.Row, col_病人Id), "挂号单=" & .TextMatrix(.Row, col_挂号单), "医嘱ID=" & .TextMatrix(.Row, col_医嘱ID))
                        End If
                    End With
                End If
            End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With fraDoctor
        .Top = lngTop
        .Left = lngLeft + 100
        .Height = lngBottom - lngTop - stbThis.Height
        If Not mbln启用输血三级审核 Then rptDoc.Height = .Height - 1600
    End With
    If mbln启用输血三级审核 Then
        rptDoc.Height = fraDoctor.Height - 1600 - picRule.Height
        picRule.Top = rptDoc.Top + rptDoc.Height
    End If
        picRule.Width = rptDoc.Width
        With fraPati
            .Top = fraDoctor.Top
            .Left = fraDoctor.Left + fraDoctor.Width + 45
            .Width = lngRight - fraDoctor.Width - 200
        End With
    
    With tbcSub
        .Top = fraPati.Top + fraPati.Height + 45
        .Left = fraPati.Left
        .Height = fraDoctor.Height - fraPati.Height - 45
        .Width = fraPati.Width + 50
    End With
    
    Me.Refresh
End Sub

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '根据权限设置按钮可见状态
    Select Case Control.ID
        Case conMenu_Edit_AdviceUnAudit '取消审核
            If Not mbln启用输血三级审核 Then
                If tbcSub.Selected.Tag <> "已审核" Then Control.Visible = False: Exit Sub
            Else
                If tbcSub.Selected.Tag <> "待签发" And mbln启用输血三级审核 Then Control.Visible = False: Exit Sub
            End If
        Case conMenu_Edit_Send
            If tbcSub.Selected.Tag <> "待签发" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_MediAudit, conMenu_Edit_UnUse
            If tbcSub.Selected.Tag <> "待审核" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_StopAudit
            If tbcSub.Selected.Tag <> "已签发" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_AdviceUnAudit
            If tbcSub.Selected.Tag <> "已审核" Then Control.Visible = False: Exit Sub
        Case conMenu_Tool_Archive '电子病案查阅
            If GetInsidePrivs(p电子病案查阅) = "" Then
                Control.Visible = False
                Exit Sub
            End If
    End Select
    Control.Visible = True
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
        Case conMenu_Edit_AdviceUnAudit '取消审核
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 3, 1) = "1"
            'If Not mbln启用输血三级审核 Then Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, col_医嘱状态) = "1"
        Case conMenu_Edit_UnUse '拒绝审核
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 1, 1) = "1"
        Case conMenu_Edit_MediAudit '审核
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 1, 1) = "1"
        Case conMenu_Edit_StopAudit '取消签发
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 5, 1) = "1" And mbln启用输血三级审核
        Case conMenu_Edit_Send '签发
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, col_审核状态) = "7" And Mid(mstrButPri, 4, 1) = "1" And mbln启用输血三级审核
        Case conMenu_View_Refresh, conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '刷新,打印
            Control.Enabled = Not mblnIsUpdate
            If mblnIsUpdate Then
                cboDept.Enabled = False
                txtFind.Enabled = False
                fraDoctor.Enabled = False
                cboDept.BackColor = &H8000000F
                txtFind.BackColor = &H8000000F
                cmdFind.Enabled = True
                cboTime.Enabled = False
            Else
                cboDept.Enabled = True
                txtFind.Enabled = True
                fraDoctor.Enabled = True
                cboTime.Enabled = True
                cmdFind.Enabled = True
                cboDept.BackColor = &H80000005
                txtFind.BackColor = &H80000005
            End If
        
        Case conMenu_Edit_ApplyView
            Control.Enabled = Val(vsAudit.TextMatrix(vsAudit.Row, COL_申请序号)) > 0
        Case conMenu_Tool_Archive '电子病案查阅
            Control.Enabled = vsAudit.Row > 0
            If Control.Enabled Then
                Control.Enabled = Val(vsAudit.TextMatrix(vsAudit.Row, col_医嘱ID)) <> 0
            End If
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
        Case conMenu_View_FindNext '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkIsShowAll_Click()
    If mblnTmp Then Exit Sub
    Call LoadDoc
End Sub

Private Sub cmdFind_Click()
    If chkIsShowAll.Value = 1 Then
        Call LoadDoc
    Else
        Call LoadAdvice(True)
    End If
End Sub

Private Sub GetLocalSetting()
'获取本地参数
    cboTime.ListIndex = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "时间范围", 0)
    mblnTmp = True
    chkIsShowAll.Value = Val(zlDatabase.GetPara("只显示有申请的医生", glngSys, mlngModul, 0) & "")
    mblnTmp = False
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mbln启用输血三级审核 = gbln输血申请三级审核
    lbl.Visible = mbln启用输血三级审核
    lbl.Enabled = mbln启用输血三级审核
    lbl.Caption = "根据中华人民共和国卫生部第85号令" & vbCrLf & vbCrLf & _
                    "一、同一患者24小时内申请备血量少于800ml的，由具有中级以上专业技术职务任职资格的医师提出申请，上级医师核准签发后，方可备血；" & vbCrLf & _
                    "二、同一患者24小时内申请备血量在800ml（含）-1600ml之间的，由具有中级以上专业技术职务任职资格的医师提出申请，上级医师审核，科室主任核准签发后，方可备血；" & vbCrLf & _
                    "三、同一患者24小时内申请备血量达到或超过1600ml的，由具有中级以上专业技术职务任职资格的医师提出申请，科室主任审核后，报医务部门批准签发，方可备血。" & vbCrLf & _
                    "以上条款规定不适用于急救用血。"
    lbl.ForeColor = vbBlue
    mstrPrivs = GetInsidePrivs(p输血审核管理)
    If mbln启用输血三级审核 Then Call GetPower
    mlngModul = p输血审核管理
    mblnIsUpdate = False
    mstrChangeRows = ""
    mstr签名IDs = "0"
    optOccasion(IIf(mint场合 = 2, 0, 1)).Value = True
    
    '---cboTime
    cboTime.AddItem "今    日"
    cboTime.AddItem "最近二天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指  定]"
    cboTime.ListIndex = 0
    '---cboDateY
    cboDateY.AddItem "今    日"
    cboDateY.AddItem "最近二天"
    cboDateY.AddItem "最近三天"
    cboDateY.AddItem "最近一周"
    cboDateY.AddItem "最近一月"
    cboDateY.AddItem "[指  定]"
    cboDateY.ListIndex = 3
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If mbln启用输血三级审核 Then
            .InsertItem(0, "  待审核  ", picUnAudited.hwnd, 0).Tag = "待审核"
            .InsertItem(1, "  待签发  ", picUnAudited.hwnd, 0).Tag = "待签发"
            .InsertItem(2, "  已签发  ", picUnAudited.hwnd, 0).Tag = "已签发"
            
            .Item(2).Selected = True
            .Item(1).Selected = True
            .Item(0).Selected = True
        Else
            .InsertItem(0, "  待审核  ", picUnAudited.hwnd, 0).Tag = "待审核"
            .InsertItem(1, "  已审核  ", picUnAudited.hwnd, 0).Tag = "已审核"
            lblDate.Caption = "审核时间"
            .Item(1).Selected = True
            .Item(0).Selected = True
        End If
    End With
    
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
    Call MainDefCommandBar
    
    'vsFlexGrid
    '-----------------------------------------------------
        strHead = ",450,1;审核说明,2000,1;审核时间;病人姓名,1000,1;医嘱内容,3460,1;期效,500,1;单量;输血时间,1550,1;开始时间,1550,1;本次治疗申请总量,1550,1;24小时输血量,1000,1;审核状态说明,800,1;医嘱ID;相关ID ; 性别;年龄;体重;入院时间;床号; 组号; 病况;护理等级;病人ID; 主页ID; 组ID;诊疗类别 ;病人来源;签名id;医嘱状态;挂号单;审核状态"
        strHead = strHead & ";门诊号;住院号;当前床号;开嘱医生;开嘱时间;开嘱科室id;出院科室id;当前病区id;主项目ID;配血医嘱;申请序号;操作状态"
        Call Grid.Init(vsAudit, strHead)
        vsAudit.ExtendLastCol = True
        vsAudit.Editable = flexEDKbdMouse
        vsAudit.ColDataType(col_选择) = flexDTBoolean
        vsAudit.Cell(flexcpChecked, 0, col_选择) = flexcpChecked
        vsAudit.Cell(flexcpPictureAlignment, 0, col_选择) = flexPicAlignCenterCenter
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mrsDefine = InitAdviceDefine
    Call GetLocalSetting '本地参数
    Call LoadDept
End Sub

Private Sub GetPower()
    Dim strSQL As String
    Dim rs As Recordset
    
    mintAuditPrivs = 0
    '具有医务科权限，则可以审核所有
    If InStr(";" & mstrPrivs & ";", ";医务科;") > 0 Then
        mintAuditPrivs = 3
    ElseIf InStr(";" & mstrPrivs & ";", ";科主任;") > 0 Then
        mintAuditPrivs = 2
    Else
        strSQL = "select 专业技术职务,管理职务 from 人员表 where id = " & UserInfo.ID
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "提取职务信息")
        If Not rs.BOF Then
            If rs("专业技术职务") = "主治医师" Or rs("专业技术职务") = "主任医师" Or rs("专业技术职务") = "副主任医师" Then
                mintAuditPrivs = 1
            End If
        End If
    End If
End Sub

Private Sub cboDateY_Click()
    Dim curDate As Date
    
    dtpTimeY(0).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    dtpTimeY(1).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTimeY(0).MaxDate = curDate
    dtpTimeY(1).MaxDate = curDate
    cmdFindY.Visible = False
    
    Select Case cboDateY.ListIndex
    Case 0 '今日
        dtpTimeY(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '最近二天
        dtpTimeY(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '最近三天
        dtpTimeY(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '最近一周
        dtpTimeY(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '最近一月
        dtpTimeY(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 '指  定
        If Me.Visible Then dtpTimeY(0).SetFocus
        cmdFindY.Visible = True
    End Select
    
    If cboDateY.ListIndex <> cboDateY.ListCount - 1 Then
        If Me.Visible Then
            If chkIsShowAll.Value = 1 Then
                Call LoadDoc
            Else
                Call LoadAdvice '(True)
            End If
        End If
    End If
End Sub


Private Sub LoadDept()
'加载操作员所属科室
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    strSQL = "Select B.ID,B.编码,B.名称 " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", ",A.缺省") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", "部门人员 A, ") & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", " And a.部门id = B.Id And A.人员ID = [1] ") & vbNewLine & _
            "  And C.工作性质 = '临床' And Instr([2],C.服务对象 || '')>0   And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"

    On Error GoTo errH
    cboDept.Clear
    '所有部门
    If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then
        cboDept.AddItem "所有部门"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, IIf(optOccasion(0).Value, "2,3", "1,3"))
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '所属缺省
        If InStr(";" & mstrPrivs & ";", ";所有部门;") = 0 Then
            If rsTmp!缺省 = 1 Then
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptDoc
        
        Set objCol = .Columns.Add(COL_人员ID, "人员ID", 0, False)
        Set objCol = .Columns.Add(col_姓名, "姓名", 70, True)
        Set objCol = .Columns.Add(COL_专业技术职务, "专业技术职务", 80, True)
        Set objCol = .Columns.Add(COL_管理职务, "管理职务", 80, True)
        Set objCol = .Columns.Add(COL_拼音简码, "拼音简码", 0, False)
        Set objCol = .Columns.Add(COL_五笔简码, "五笔简码", 0, False)
        Set objCol = .Columns.Add(COL_所属部门, "所属部门", 0, False)
        Set objCol = .Columns.Add(COL_所属部门ID, "所属部门ID", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的医生..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then .GroupsOrder.Add .Columns(COL_所属部门)
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
     
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MediAudit, "审核(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "拒绝审核(&R)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "取消审核(&U)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "签发(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "取消(&Q)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
        objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Preview
    End With
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MediAudit, "审核(&U)")
            objControl.BeginGroup = True
            objControl.IconId = 21904
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "拒绝审核(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "签发(&A)")
            objControl.BeginGroup = True
            objControl.IconId = 21904
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "取消审核(&U)")
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "取消(&Q)")
            objControl.BeginGroup = True
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        Set objCustom = .Add(xtpControlCustom, conMenu_View_FindType, "场合")
            objCustom.Handle = fraType.hwnd
            objCustom.Flags = xtpFlagRightAlign

    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    '右键菜单(快速审核功能)
    '-----------------------------------------------------
    Set mobjPopup = cbsMain.Add("右键菜单", xtpBarPopup)
    With mobjPopup.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "快速审核")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "取消审核")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅")
    End With
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsUpdate = True Then
        If MsgBox("当前输入的内容未保存，是否要退出？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    mlngFindNum = 0
    Set mclsMipModule = Nothing
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "时间范围", cboTime.ListIndex
    zlDatabase.SetPara "只显示有申请的医生", chkIsShowAll.Value & "", glngSys, mlngModul
End Sub


Private Sub optOccasion_Click(Index As Integer)
    If Me.Visible Then
                Call LoadDept
        vsAudit.Rows = 1
        vsAudit.AddItem ""
    End If
End Sub

Private Sub picUnAudited_Resize()
    On Error Resume Next
    picDate.Move 0, 0, picUnAudited.Width
    picDateY.Move 0, 0, picUnAudited.Width
    vsAudit.Move 0, picDate.Top + picDate.Height, picUnAudited.Width, picUnAudited.Height - picDate.Top + picDate.Height
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '加载医嘱列表
    If tbcSub.Selected.Tag = "待审核" Then
        If Me.Visible Then
            Call LoadAdvice
            If mbln启用输血三级审核 Then Call vsAudit_KeyPress(vbKeyBack)
        End If
    ElseIf tbcSub.Selected.Tag = "待签发" Then
        If Me.Visible Then
            Call LoadAdvice
            Call vsAudit_KeyPress(vbKeyBack)
        End If
    Else
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible And chkIsShowAll.Value = 1 Then
        Call LoadDoc
    End If
    With vsAudit
        If mbln启用输血三级审核 Then
            If Item.Tag = "已签发" Then
                picDate.Visible = True
                picDateY.Visible = False
                Call picUnAudited_Resize
                .ColWidth(COL_审核时间) = 1800
                .ColHidden(COL_审核时间) = False
                .TextMatrix(0, COL_审核时间) = "签发时间"
                .Cell(flexcpChecked, 1, 0) = 2
                If Me.Visible Then Call LoadAdvice(True)
                .ColHidden(col_审核状态说明) = False
                lblPri(1).ForeColor = vbBlue
                lblPri(1).Caption = "您可取消签发的血液容量范围：" & IIf(mintAuditPrivs >= 1, "800ml以下   ", "") & IIf(mintAuditPrivs >= 2, "800ml-1600ml   ", "") & IIf(mintAuditPrivs = 3, "1600ml及以上", "")
                If lblPri(1).Caption = "您可签发的血液容量范围：" Then
                    lblPri(1).Caption = "您不具有取消签发血液的权限！"
                    lblPri(1).ForeColor = vbRed
                End If
                lblPri(1).Width = 8000
            ElseIf Item.Tag = "待签发" Then
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Caption = "您可签发的血液容量范围：" & IIf(mintAuditPrivs >= 1, "800ml以下   ", "") & IIf(mintAuditPrivs >= 2, "800ml-1600ml   ", "") & IIf(mintAuditPrivs = 3, "1600ml及以上", "")
                lblPri(0).ForeColor = vbBlue
                If lblPri(0).Caption = "您可签发的血液容量范围：" Then
                    lblPri(0).Caption = "您不具有签发血液的权限！"
                    lblPri(0).ForeColor = vbRed
                End If
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_审核时间) = 1800
                .ColHidden(COL_审核时间) = False
                .TextMatrix(0, COL_审核时间) = "审核时间"
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_审核状态说明) = True
            ElseIf Item.Tag = "待审核" Then
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Caption = "您可审核的血液容量范围：" & IIf(mintAuditPrivs >= 1, "800ml以下   ", "") & IIf(mintAuditPrivs >= 1, "800ml-1600ml   ", "") & IIf(mintAuditPrivs >= 2, "1600ml及以上", "")
                lblPri(0).ForeColor = vbBlue
                If lblPri(0).Caption = "您可审核的血液容量范围：" Then
                    lblPri(0).Caption = "您不具有审核血液的权限！"
                    lblPri(0).ForeColor = vbRed
                End If
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_审核时间) = 0
                .ColHidden(COL_审核时间) = True
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_审核状态说明) = True
            End If
            If Me.Visible Then Call vsAudit_KeyPress(vbKeyBack)
        Else
            .ColHidden(col_24h输血量) = True
            .ColHidden(col_输血总量) = True
            If Item.Tag = "已审核" Then
                picDate.Visible = True
                picDateY.Visible = False
                Call picUnAudited_Resize
                .ColWidth(COL_审核时间) = 1800
                .ColHidden(COL_审核时间) = False
                .TextMatrix(0, COL_审核时间) = "审核时间"
                .Cell(flexcpChecked, 1, 0) = 2
                If Me.Visible Then Call LoadAdvice(True)
                .ColHidden(col_审核状态说明) = False
            Else
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_审核时间) = 0
                .ColHidden(COL_审核时间) = True
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_审核状态说明) = True
            End If
        End If
    End With
End Sub

Private Sub txtFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub LoadAdvice(Optional ByVal blnIsAudited As Boolean)
'加载待审核和已审核的医嘱
'参数：是否加载已审核医嘱,为空为加载待审核医嘱
    Dim rsTmp As Recordset, rsTemp As Recordset
    Dim strSQL As String, strTemp As String
    Dim strDoorIds As String, strKey As String
    Dim i As Long, j As Long
    Dim lngID As Long       '用于定位
    Dim strFormat As String
    Dim strTmp As String, strType As String, int审核状态 As Integer
    Dim blnDo As Boolean
    Dim strPatis As String   '病人信息字符串：病人ID1:主页ID1,病人ID2:主页ID2···
    Dim strDate As String '日期字符串
    Dim int选择 As Integer
    Dim dbl输血总量 As Double, dbl24h量 As Double
    
    
    If tbcSub.Selected.Tag = "待签发" Then
        strType = "C.操作类型 = 18"
        int审核状态 = 7
    ElseIf tbcSub.Selected.Tag = "待审核" Then
        strType = "C.操作类型 = 19"
        int审核状态 = 1
    Else
        strType = "C.操作类型 in(11,12,14,15)"
    End If
    strSQL = "Select Decode(a.病人来源, 2, a.病人id || '_' || a.主页id || '_' || Nvl(a.婴儿, 0), 1, a.病人id || '_' || a.挂号单) Key,a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别,  Null As 选择, Null As 输入, " & vbNewLine & _
            " Decode(Nvl(a.婴儿, 0), 0, a.姓名, Nvl(q.婴儿姓名, a.姓名 || '之婴' || q.序号)) As 姓名,Decode(Nvl(a.婴儿, 0), 0, a.性别, q.婴儿性别) As 性别," & vbNewLine & _
            " Decode(Nvl(a.婴儿, 0), 0, a.年龄, (Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间) || '天')) As 年龄, p.当前床号 As 床号," & vbNewLine & _
            "       Decode(Nvl(a.医嘱期效, 0), 0, '长嘱', '临嘱') As 期效, To_Char(a.开始执行时间, 'YYYY-MM-DD HH24:MI') As 开始时间, a.医嘱内容,a.审核状态," & vbNewLine & _
            "       Decode(a.总给予量, Null, Null, a.总给予量 || b.计算单位) As 总量, NVL(to_char(A.手术时间,'YYYY-MM-DD HH24:MI'),a.标本部位) As 输血时间, a.执行时间方案 As 执行时间方案, a.病人id," & vbNewLine & _
            "       a.主页id, a.诊疗项目id, a.频率次数, a.频率间隔, a.间隔单位, b.计算单位 As 单量单位, e.体重,e.入院日期,e.入院病况,f.名称 as 护理等级,a.病人来源,A.挂号单,a.申请序号" & vbNewLine & _
            ", c.操作类型, c.操作说明, c.签名id ,a.医嘱状态,c.操作时间 as 审核时间" & _
            ",p.门诊号,p.住院号,p.当前床号,a.开嘱医生,To_Char(a.开嘱时间,'YYYY-MM-DD HH24:MI') As 开嘱时间,a.开嘱科室id,e.出院科室id,e.当前病区id,a.诊疗项目id,h.执行分类" & _
            " From 病人医嘱记录 A, 病人信息 P, 诊疗项目目录 B, 病案主页 E,收费项目目录 F" & vbNewLine & _
            ", (Select 医嘱id,操作时间,操作说明,操作类型,签名ID" & vbNewLine & _
                            "From (Select C.医嘱id,C.操作时间,C.操作说明,C.操作类型,C.签名ID, Row_Number() Over(Partition By C.医嘱id Order By C.操作时间 Desc) Top" & vbNewLine & _
                            "       From 病人医嘱状态 C" & vbNewLine & _
                            "       Where c.操作时间 Between " & IIf(InStr(1, tbcSub.Selected.Tag, "已") > 0, "[3] And [4]", "[6] And [7]") & vbNewLine & _
                            "       and " & strType & " And C.操作人员 =[2])" & vbNewLine & _
                            "Where Top = 1)  C" & ",病人医嘱记录 G,诊疗项目目录 H,病人新生儿记录 Q" & _
            " Where a.病人id = p.病人id And a.诊疗项目id = b.Id  And f.id(+)=e.护理等级id  And" & vbNewLine & _
            "      e.病人id(+) = a.病人id And e.主页id(+) = a.主页id and g.诊疗类别 = 'E' And a.id=g.相关id and g.诊疗项目id=h.id And (H.操作类型='8' and nvl(H.执行分类,0)=0  or H.操作类型='9')  and A.病人ID = Q.病人ID(+) and A.主页ID = Q.主页ID(+) and A.婴儿 = Q.序号(+) " & _
            IIf(InStr(1, tbcSub.Selected.Tag, "已") > 0, " And c.医嘱id = a.Id ", _
            " AND a.id=c.医嘱id(+) And A.开嘱时间 between [6] and [7] And a.医嘱状态 = 1 And a.审核状态 = [8] ") & vbNewLine & _
            "    And a.开嘱医生=[1] And A.病人来源=[5] And a.诊疗类别 ='K'  And a.相关ID is null " & _
            " Order By p.姓名,To_Char(a.开始执行时间, 'YYYY-MM-DD HH24:MI'),Nvl(a.相关id, a.Id),a.id"
            '" & IIf(tbcSub.Selected.Tag = "待审核", "And a.医嘱状态 = 1 ", "") & "
    stbThis.Panels(2).Text = ""
    If rptDoc.SelectedRows.Count = 0 Then
        stbThis.Panels(2).Text = "请选择一位开嘱医生。"
        Exit Sub
    End If
    If rptDoc.SelectedRows(0).GroupRow Then
        vsAudit.Rows = 1
        vsAudit.AddItem ""
        stbThis.Panels(2).Text = "请选择一位开嘱医生。"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rptDoc.SelectedRows(0).Record(col_姓名).Value, UserInfo.姓名, CDate(dtpTime(0).Value), CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60), IIf(optOccasion(0).Value, 2, 1), CDate(dtpTimeY(0).Value), CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60), int审核状态)
    If mbln启用输血三级审核 Then    '当启用了三级审核时，计算24小时输血量以及输血总量
        If Not rsTmp.BOF Then
            rsTmp.MoveFirst
            strTemp = ""
            Do While Not rsTmp.EOF
                If optOccasion(0).Value = True Then
                    strKey = rsTmp("病人id") & ":" & rsTmp("主页id")
                Else
                    strKey = rsTmp("挂号单")
                End If
                If InStr("," & strTemp & ",", "," & strKey & ",") = 0 Then
                        strTemp = strTemp & "," & strKey
                End If
                rsTmp.MoveNext
            Loop
            If Left(strTemp, 1) = "," Then strTemp = Mid(strTemp, 2)
            If optOccasion(0).Value Then
                strSQL = _
                    " Select Key, Id, 开嘱时间, 申请量, 输血时间" & vbNewLine & _
                    " From (With 医嘱记录 As (Select /*+ CARDINALITY(d,10) */" & vbNewLine & _
                    "                     a.病人id || '_' || a.主页id || '_' || Nvl(a.婴儿, 0) Key, a.Id," & vbNewLine & _
                    "                     Decode(Nvl(e.医嘱id, 0), 0, a.诊疗项目id, e.诊疗项目id) 诊疗项目id," & vbNewLine & _
                    "                     Decode(Nvl(e.医嘱id, 0), 0, a.总给予量, e.申请量) 申请量, a.开嘱时间," & vbNewLine & _
                    "                     Nvl(To_Char(a.手术时间, 'YYYY-MM-DD HH24:MI'), a.标本部位) As 输血时间" & vbNewLine & _
                    "                    From 输血申请项目 e, 诊疗项目目录 b, 病人医嘱记录 c, 病人医嘱记录 a, Table(f_Str2list2([1])) d" & vbNewLine & _
                    "                    Where e.医嘱id(+) = a.Id And b.Id = c.诊疗项目id And (b.操作类型 = '8' And Nvl(b.执行分类, 0) = 0 Or b.操作类型 = '9') And" & vbNewLine & _
                    "                          c.诊疗类别 = 'E' And c.相关id = a.Id And a.病人id = d.C1 And a.主页id = d.C2 And a.诊疗类别 = 'K' And" & vbNewLine & _
                    "                          a.医嘱状态 Not In (-1, 2, 4))" & vbNewLine & _
                    "       Select b.Key, b.Id, b.开嘱时间, b.申请量 * Decode(Upper(a.计算单位), 'ML', 1, Nvl(a.计算系数, 1)) 申请量, b.输血时间" & vbNewLine & _
                    "       From 诊疗项目目录 a, 医嘱记录 b" & vbNewLine & _
                    "       Where a.Id = b.诊疗项目id)"
            Else
                strSQL = _
                    " Select Key, Id, 开嘱时间, 申请量, 输血时间" & vbNewLine & _
                    " From (With 医嘱记录 As (Select /*+ CARDINALITY(d,10) */" & vbNewLine & _
                    "                     a.病人id || '_' || a.挂号单 Key, a.Id, Decode(Nvl(e.医嘱id, 0), 0, a.诊疗项目id, e.诊疗项目id) 诊疗项目id," & vbNewLine & _
                    "                     Decode(Nvl(e.医嘱id, 0), 0, a.总给予量, e.申请量) 申请量, a.开嘱时间," & vbNewLine & _
                    "                     Nvl(To_Char(a.手术时间, 'YYYY-MM-DD HH24:MI'), a.标本部位) As 输血时间" & vbNewLine & _
                    "                    From 输血申请项目 e, 诊疗项目目录 b, 病人医嘱记录 c, 病人医嘱记录 a, Table(f_Str2list([1])) d" & vbNewLine & _
                    "                    Where e.医嘱id(+) = a.Id And b.Id = c.诊疗项目id And (b.操作类型 = '8' And Nvl(b.执行分类, 0) = 0 Or b.操作类型 = '9') And" & vbNewLine & _
                    "                          c.诊疗类别 = 'E' And c.相关id = a.Id And a.挂号单 = d.Column_Value And a.诊疗类别 = 'K' And" & vbNewLine & _
                    "                          a.医嘱状态 Not In (-1, 2, 4))" & vbNewLine & _
                    "       Select b.Key, b.Id, b.开嘱时间, b.申请量 * Decode(Upper(a.计算单位), 'ML', 1, Nvl(a.计算系数, 1)) 申请量, b.输血时间" & vbNewLine & _
                    "       From 诊疗项目目录 a, 医嘱记录 b" & vbNewLine & _
                    "       Where a.Id = b.诊疗项目id)"
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
        End If
    End If
    
    With vsAudit
        If Val(.TextMatrix(.Row, col_医嘱ID)) <> 0 Then lngID = Val(.TextMatrix(.Row, col_医嘱ID))
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If mbln启用输血三级审核 Then
                    If Not rsTemp.BOF Then
                        rsTemp.MoveFirst
                        dbl输血总量 = 0
                        dbl24h量 = 0
                        rsTemp.Filter = "Key ='" & rsTmp!Key & "'"
                        Do While Not rsTemp.EOF
                            dbl输血总量 = dbl输血总量 + rsTemp("申请量")
                            If rsTemp!输血时间 <> "" And rsTmp("输血时间") <> "" Then
                                If CDate(rsTemp!输血时间) > CDate(rsTmp("输血时间")) - 1 And CDate(rsTemp!输血时间) <= CDate(rsTmp("输血时间")) Then dbl24h量 = dbl24h量 + rsTemp("申请量")
                            ElseIf rsTemp!输血时间 & "" = "" And rsTmp("输血时间") <> "" Then
                                If CDate(rsTemp!开嘱时间) > CDate(rsTmp("输血时间")) - 1 And CDate(rsTemp!开嘱时间) <= CDate(rsTmp("输血时间")) Then dbl24h量 = dbl24h量 + rsTemp("申请量")
                            End If
                            rsTemp.MoveNext
                        Loop
                    End If
                End If
                If tbcSub.Selected.Tag = "待审核" Then
                    '根据审核状态过滤
                    If rsTmp!审核状态 <> 1 Then GoTo loopNext
                    '根据用户权限过滤
                    If mbln启用输血三级审核 Then
                        If dbl24h量 < 800 And mintAuditPrivs < 1 Then GoTo loopNext
                        If dbl24h量 >= 800 And dbl24h量 < 1600 And mintAuditPrivs < 1 Then GoTo loopNext
                        If dbl24h量 >= 1600 And mintAuditPrivs < 2 Then GoTo loopNext
                    End If
                ElseIf tbcSub.Selected.Tag = "待签发" Then
                    '根据审核状态过滤
                    If rsTmp!审核状态 <> 7 Then GoTo loopNext
                    '根据用户权限过滤
                    If dbl24h量 < 800 And mintAuditPrivs < 1 Then GoTo loopNext
                    If dbl24h量 >= 800 And dbl24h量 < 1600 And mintAuditPrivs < 2 Then GoTo loopNext
                    If dbl24h量 >= 1600 And mintAuditPrivs < 3 Then GoTo loopNext
                End If
                .AddItem ""
                .TextMatrix(i, COL_病人姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, col_期效) = rsTmp!期效 & ""
                .TextMatrix(i, COL_总量) = rsTmp!总量 & ""
                .TextMatrix(i, col_输血时间) = rsTmp!输血时间 & ""
                .TextMatrix(i, COL_开始时间) = rsTmp!开始时间 & ""
                .TextMatrix(i, col_医嘱ID) = rsTmp!ID & ""
                If Val(rsTmp!ID & "") = lngID And lngID <> 0 Then
                    .Row = i
                End If
                If mbln启用输血三级审核 Then
                    .TextMatrix(i, col_24h输血量) = dbl24h量 & "ml"
                    .TextMatrix(i, col_输血总量) = dbl输血总量 & "ml"
                End If
                If tbcSub.Selected.Tag = "已签发" Then
                    .TextMatrix(i, col_审核状态说明) = Decode(rsTmp!审核状态 & "", "", "无需审核", "1", "待审核", "2", "审核通过", "3", "审核未通过", "4", "血库待接收", "5", "血库配血中", "6", "血库停止配血", "7", "血液待签发")
                End If
                .TextMatrix(i, col_相关ID) = rsTmp!相关ID & ""
                .TextMatrix(i, col_性别) = rsTmp!性别 & ""
                .TextMatrix(i, col_年龄) = rsTmp!年龄 & ""
                .TextMatrix(i, COL_体重) = rsTmp!体重 & ""
                .TextMatrix(i, COL_入院时间) = rsTmp!入院日期 & ""
                .TextMatrix(i, col_床号) = rsTmp!床号 & ""
                .TextMatrix(i, COL_护理等级) = rsTmp!护理等级 & ""
                .TextMatrix(i, col_病人Id) = rsTmp!病人ID & ""
                .TextMatrix(i, col_主页ID) = rsTmp!主页ID & ""
                .TextMatrix(i, col_挂号单) = rsTmp!挂号单 & ""
                .TextMatrix(i, col_审核状态) = Val(rsTmp!审核状态 & "")
                .TextMatrix(i, COL_申请序号) = Val(rsTmp!申请序号 & "")
                .TextMatrix(i, COL_操作类型) = Val(rsTmp!操作类型 & "")
                If optOccasion(1).Value Then
                    If InStr(strPatis, rsTmp!病人ID & ":" & rsTmp!挂号单) = 0 Then
                        strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!挂号单
                    End If
                Else
                    If InStr(strPatis, rsTmp!病人ID & ":" & rsTmp!主页ID) = 0 Then
                        strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
                    End If
                End If
                If InStr(strDate, Format(rsTmp!输血时间 & "", "YYYY-MM-DD")) = 0 Then
                    strDate = strDate & "," & Format(rsTmp!输血时间 & "", "YYYY-MM-DD")
                End If
                .TextMatrix(i, COL_组ID) = rsTmp!组ID & ""
                .TextMatrix(i, COL_诊疗类别) = rsTmp!诊疗类别 & ""
                .TextMatrix(i, COL_病人来源) = rsTmp!病人来源 & ""
                .TextMatrix(i, COL_病况) = rsTmp!入院病况 & ""
                '显示简洁模式下的医嘱内容
                strFormat = rsTmp!医嘱内容
                blnDo = True
                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                If blnDo Then
                    strTmp = .TextMatrix(i, COL_总量)
                    If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
                End If
                
                .TextMatrix(i, COL_医嘱内容) = strFormat
                If blnIsAudited Then
                    .TextMatrix(i, COL_签名id) = rsTmp!签名id & ""
                    .TextMatrix(i, COL_医嘱状态) = rsTmp!医嘱状态 & ""
                    
                    int选择 = Val(rsTmp!操作类型 & "") - 10
                    If gbln血库系统 Then
                        If InStr(",2,4,5,", "," & Val(.TextMatrix(i, col_审核状态)) & ",") > 0 Then
                            int选择 = 1
                        End If
                    End If
                    '如果医嘱不再是新开状态，则改变字体颜色
                    If Val(rsTmp!医嘱状态 & "") <> 1 Then
                        .Cell(flexcpForeColor, i, col_选择, i, COL_签名id) = &HC00000
                    End If
                    
                    '进入血库系统后，则改变字体颜色，淡红色
                    If gbln血库系统 Then
                        If InStr(",2,5,", "," & Val(.TextMatrix(i, col_审核状态)) & ",") > 0 Then
                            .Cell(flexcpForeColor, i, col_选择, i, COL_签名id) = &H8080FF
                        End If
                    End If
                    
                End If
                .TextMatrix(i, COL_审核说明) = rsTmp!操作说明 & ""
                .TextMatrix(i, COL_审核时间) = Format(rsTmp!审核时间 & "", "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, COL_门诊号) = rsTmp!门诊号 & ""
                .TextMatrix(i, col_住院号) = rsTmp!住院号 & ""
                .TextMatrix(i, COL_当前床号) = rsTmp!当前床号 & ""
                .TextMatrix(i, COL_开嘱医生) = rsTmp!开嘱医生 & ""
                .TextMatrix(i, COL_开嘱时间) = rsTmp!开嘱时间 & ""
                .TextMatrix(i, COL_开嘱科室ID) = rsTmp!开嘱科室ID & ""
                .TextMatrix(i, COL_出院科室ID) = rsTmp!出院科室ID & ""
                .TextMatrix(i, COL_当前病区ID) = rsTmp!当前病区ID & ""
                .TextMatrix(i, COL_主项目ID) = rsTmp!诊疗项目ID & ""
                .TextMatrix(i, col_配血医嘱) = Val(rsTmp!执行分类 & "")
                i = i + 1
loopNext:           rsTmp.MoveNext
                
            Loop
            
            strPatis = Mid(strPatis, 2)
            strDate = Mid(strDate, 2)
            If .Rows = 1 Then .AddItem ""
        Else
            .AddItem ""
        End If
        strFormat = "【开嘱医生：" & rptDoc.SelectedRows(0).Record(col_姓名).Value & "】"
        If blnIsAudited Then
            strTmp = "在【审核时间：" & Format(dtpTime(0).Value, "YYYY-MM-DD") & " 00:00:00 - " & Format(dtpTime(1).Value, "YYYY-MM-DD") & " 23:59:59】内，"
            If Val(.TextMatrix(1, col_医嘱ID)) = 0 Then
                strTmp = strTmp & strFormat & "不存在被审核过的医嘱。"
            Else
                strTmp = strTmp & strFormat & "共有" & (.Rows - 1) & "条医嘱被审核。"
            End If
        Else
            strTmp = "在【开嘱时间：" & Format(dtpTimeY(0).Value, "YYYY-MM-DD") & " 00:00:00 - " & Format(dtpTimeY(1).Value, "YYYY-MM-DD") & " 23:59:59】内，"
            If Val(.TextMatrix(1, col_医嘱ID)) = 0 Then
                strTmp = strTmp & strFormat & "不存在需要审核的医嘱。"
            Else
                strTmp = strTmp & strFormat & "共有" & (.Rows - 1) & "条医嘱需要审核。"
            End If
        End If
        stbThis.Panels(2).Text = strTmp
        '自动调整行高
        .AutoSize COL_医嘱内容
        .Redraw = flexRDDirect
        If .Row > 0 Then Call vsAudit_AfterRowColChange(1, 1, .Row, COL_审核说明)
    End With
    Call vsAudit_KeyPress(vbKeyBack)
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If zlCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(col_姓名).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(gint简码 = 0, COL_拼音简码, COL_五笔简码)).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        '该行选中且显示在可见区域,并引发SelectionChanged事件
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(col_姓名).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "当前部门没有找到您查找的医生。", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "已经是最后一个医生了。", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub


Private Sub vsAudit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    With vsAudit
        If Me.Visible = False Then Exit Sub
        If NewCol = COL_审核说明 And tbcSub.Selected.Tag = "待审核" Or NewCol = col_选择 Then
            If (Val(vsAudit.Cell(flexcpData, NewRow, col_选择) & "") = "1" And NewCol = COL_审核说明) Or _
                    (vsAudit.TextMatrix(NewRow, COL_医嘱状态) & "" <> "1" And vsAudit.TextMatrix(NewRow, COL_医嘱状态) & "" <> "" And NewCol = COL_审核说明) _
                    Or (tbcSub.Selected.Tag = "已审核" And NewCol = col_选择) Then
                vsAudit.FocusRect = flexFocusNone
            Else
                If .TextMatrix(NewRow, COL_医嘱状态) & "" <> "1" And .TextMatrix(NewRow, COL_医嘱状态) & "" <> "" Then
                    vsAudit.FocusRect = flexFocusNone
                Else
                    vsAudit.FocusRect = flexFocusHeavy
                End If
            End If
        Else
            vsAudit.FocusRect = flexFocusNone
        End If

        '颜色
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)

        If vsAudit.TextMatrix(NewRow, col_医嘱ID) <> "" And NewRow <> 0 Then
            lblInformation(info_入院时间).Caption = Format(.TextMatrix(NewRow, COL_入院时间), "yyyy-MM-dd")
            lblInformation(info_性别).Caption = .TextMatrix(NewRow, col_性别)
            lblInformation(info_年龄).Caption = .TextMatrix(NewRow, col_年龄)
            lblInformation(info_病况).Caption = .TextMatrix(NewRow, COL_病况)
            lblInformation(info_床号).Caption = .TextMatrix(NewRow, col_床号)
            lblInformation(info_护理等级).Caption = .TextMatrix(NewRow, COL_护理等级)
            lblInformation(info_体重).Caption = IIf(Val(.TextMatrix(NewRow, COL_体重) & "") = 0, "", .TextMatrix(NewRow, COL_体重) & "Kg")
            
            '过敏记录
            Call LoadPatiAllergy(Val(.TextMatrix(NewRow, col_病人Id) & ""), cbo过敏)
            
            '诊断
            lblInformation(info_诊断).Caption = GetPatiDiagnose(Val(.TextMatrix(NewRow, col_病人Id) & ""), _
            Val(.TextMatrix(NewRow, col_主页ID) & ""), _
            Val(.TextMatrix(NewRow, COL_病人来源)))
            '住院信息显示
            picInShow.Visible = Not .TextMatrix(NewRow, COL_病人来源) = "1"
        Else
            lblInformation(info_入院时间).Caption = ""
            lblInformation(info_性别).Caption = ""
            lblInformation(info_年龄).Caption = ""
            lblInformation(info_病况).Caption = ""
            lblInformation(info_床号).Caption = ""
            lblInformation(info_护理等级).Caption = ""
            lblInformation(info_体重).Caption = ""
            
            '过敏记录
            cbo过敏.Clear
            
            '诊断
            lblInformation(info_诊断).Caption = ""
            
            picInShow.Visible = True
        End If
    End With
End Sub

Private Sub vsAudit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = COL_审核说明) Then
        Cancel = True
    Else
        With vsAudit
            If .TextMatrix(1, col_医嘱ID) & "" = "" Or Val(.TextMatrix(.Row, col_审核状态)) = 3 Or _
                    (.TextMatrix(Row, COL_医嘱状态) & "" <> "1" And .TextMatrix(Row, COL_医嘱状态) & "" <> "") Then
                Cancel = True
            End If
            If gbln血库系统 And InStr(",2,5,", "," & Val(.TextMatrix(Row, col_审核状态)) & ",") > 0 Then
                Cancel = True
            End If
        End With
    End If
End Sub

Private Sub vsAudit_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsAudit
        mstrButPri = "00000"
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_选择) = vbChecked Then
                Select Case tbcSub.Selected.Tag
                    Case "待审核"
                        mstrButPri = "11000"
                    Case "待签发"
                        mstrButPri = "00110"
                    Case "已签发"
                        mstrButPri = "00001"
                    Case "已审核"
                        mstrButPri = "00100"
                End Select
                Exit For
            End If
        Next
        If tbcSub.Selected.Tag = "待签发" And .Col = col_选择 Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpBackColor, i, col_选择) = &HFFC0FF And .Cell(flexcpChecked, i, col_选择) = vbChecked Then
                    mstrButPri = "00100"
                    Exit For
                End If
            Next
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_选择) = vbChecked Then
                tbcSub.Item(0).Enabled = False
                tbcSub.Item(1).Enabled = False
                If mbln启用输血三级审核 Then tbcSub.Item(2).Enabled = False
                tbcSub(tbcSub.Selected.Index).Enabled = True
                Exit Sub
            End If
        Next
        tbcSub.Item(0).Enabled = True
        tbcSub.Item(1).Enabled = True
        If mbln启用输血三级审核 Then tbcSub.Item(2).Enabled = True
    End With
End Sub

Private Sub vsAudit_Click()
    Call vsAudit_KeyPress(vbKeySpace)
End Sub

Private Function RowIn同一病人(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'功能：判断指定病人是否有其他医嘱
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If lngRow = 0 Then Exit Function
        If .TextMatrix(lngRow - 1, COL_病人姓名) = .TextMatrix(lngRow, COL_病人姓名) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, COL_病人姓名) = .TextMatrix(lngRow, COL_病人姓名) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_病人姓名) = .TextMatrix(lngRow, COL_病人姓名) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, COL_病人姓名) = .TextMatrix(lngRow, COL_病人姓名) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn同一病人 = blnTmp
    End With
End Function

Private Sub vsAudit_KeyPress(KeyAscii As Integer)
    Dim lngloop As Long

    With vsAudit
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call UnAuditEnterNextCell
        ElseIf KeyAscii = vbKeyBack Then
            .Cell(flexcpChecked, 0, col_选择) = flexUnchecked
        ElseIf .Col = COL_审核说明 And .Cell(flexcpForeColor, .Row, COL_病人姓名) <> &HFFC0FF And Val(.TextMatrix(.Row, col_审核状态)) <> 3 Then
            .ComboList = "" '使按钮状态进入输入状态
        ElseIf .Col = col_选择 And KeyAscii = vbKeySpace Then
            If .TextMatrix(1, col_医嘱ID) = "" Then Exit Sub
            If .MouseRow = .FixedRows - 1 Then
                If .Cell(flexcpChecked, 0, col_选择) = flexChecked Then
                    .Cell(flexcpChecked, 0, col_选择, .Rows - 1, col_选择) = flexUnchecked
                Else
                    .Cell(flexcpChecked, 0, col_选择, .Rows - 1, col_选择) = flexChecked
                End If
            ElseIf .MouseRow < .Rows Then
                If .Cell(flexcpChecked, .Row, col_选择) = flexChecked Then
                    .Cell(flexcpChecked, .Row, col_选择) = flexUnchecked
                Else
                    .Cell(flexcpChecked, .Row, col_选择) = flexChecked
                End If
            End If
        End If
        If mbln启用输血三级审核 Then
            Select Case tbcSub.Selected.Tag
                Case "待审核"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, col_24h输血量)) < 800 Then
                            If mintAuditPrivs < 1 Then .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                        ElseIf Val(.TextMatrix(lngloop, col_24h输血量)) >= 1600 Then
                            If mintAuditPrivs < 2 Then .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                        Else
                            If mintAuditPrivs < 1 Then .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                        End If
                    Next
                Case "待签发"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, col_24h输血量)) < 800 Then
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            End If
                        ElseIf Val(.TextMatrix(lngloop, col_24h输血量)) >= 1600 Then
                            If mintAuditPrivs < 2 Then
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            ElseIf mintAuditPrivs = 2 Then
                                .Cell(flexcpBackColor, lngloop, col_选择) = &HFFC0FF
                            End If
                        Else
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            ElseIf mintAuditPrivs = 1 Then '表明只能回退，不能签发
                                .Cell(flexcpBackColor, lngloop, col_选择) = &HFFC0FF
                            End If
                        End If
                        .Cell(flexcpBackColor, lngloop, COL_审核说明) = .Cell(flexcpBackColor, lngloop, col_选择)
                    Next
                Case "已签发"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, COL_医嘱状态)) <> 1 Or InStr(1, "'2'3'4'", "'" & .TextMatrix(lngloop, col_审核状态) & "'") > 1 Then '(Val(.TextMatrix(lngloop, col_审核状态)) <> 4 And Val(.TextMatrix(lngloop, col_审核状态)) <> 3) Then
                            .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            .Cell(flexcpBackColor, lngloop, col_选择, lngloop, col_24h输血量) = &H80000016
                        ElseIf Val(.TextMatrix(lngloop, col_24h输血量)) < 800 Then
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpBackColor, lngloop, col_选择) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            End If
                        ElseIf Val(.TextMatrix(lngloop, col_24h输血量)) >= 1600 Then
                            If mintAuditPrivs < 3 Then
                                .Cell(flexcpBackColor, lngloop, col_选择) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            End If
                        Else
                            If mintAuditPrivs < 2 Then
                                .Cell(flexcpBackColor, lngloop, col_选择) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_选择) = flexUnchecked
                            End If
                        End If
                        .Cell(flexcpBackColor, lngloop, COL_审核说明) = .Cell(flexcpBackColor, lngloop, col_选择)
                    Next
            End Select
        End If
    End With
End Sub

Private Sub vsAudit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = COL_审核说明 Then
        If zlCommFun.ActualLen(vsAudit.Editable) - zlCommFun.ActualLen(vsAudit.EditSelText) >= 100 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
            If KeyAscii = vbKeyReturn Then
                Call UnAuditEnterNextCell
                Exit Sub
            End If
            KeyAscii = 0
        ElseIf Chr(KeyAscii) = "'" Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub UnAuditEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAudit
        If .Col = COL_审核说明 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAudit_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
        
    With vsAudit
        If tbcSub.Selected.Tag = "待签发" And .Col = col_选择 Then
            mstrButPri = "1"
            For i = 1 To .Rows - 1
                If .Cell(flexcpBackColor, i, col_选择) = &HC0FFFF And .Cell(flexcpChecked, i, col_选择) = vbChecked Then
                    mstrButPri = "0"
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsAudit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then mobjPopup.ShowPopup
End Sub

Private Sub vsAudit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = COL_审核说明 Then
        vsAudit.EditSelStart = 0
        vsAudit.EditSelLength = Len(vsAudit.EditText)
    End If
End Sub


