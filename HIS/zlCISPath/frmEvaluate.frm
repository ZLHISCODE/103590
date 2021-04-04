VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEvaluate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "路径评估"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9945
   Icon            =   "frmEvaluate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9945
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgNature 
      Left            =   8280
      Top             =   2160
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
            Picture         =   "frmEvaluate.frx":617A
            Key             =   "Selected"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":6514
            Key             =   "UnSelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":68AE
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":6C48
            Key             =   "UnCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStart 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   2640
      Width           =   9855
      Begin VB.OptionButton optStart 
         Caption         =   "入科时间作为入径第1天,目前是入径第n天"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   29
         Top             =   30
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton optStart 
         Caption         =   "当前时间作为入径第1天"
         Height          =   250
         Index           =   1
         Left            =   5040
         TabIndex        =   28
         Top             =   30
         Width           =   2895
      End
      Begin VB.Label lblStart 
         Caption         =   "路径起点"
         Height          =   230
         Left            =   120
         TabIndex        =   30
         Top             =   45
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   120
         X2              =   10000
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   120
         X2              =   10000
         Y1              =   315
         Y2              =   315
      End
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      TabIndex        =   21
      Top             =   3000
      Width           =   9855
      Begin VB.OptionButton optDate 
         Caption         =   "下一阶段提前至明天"
         Height          =   250
         Index           =   3
         Left            =   4680
         TabIndex        =   36
         Top             =   35
         Width           =   1935
      End
      Begin VB.OptionButton optDate 
         Caption         =   "下一天的阶段延后（继续当前阶段）"
         Height          =   250
         Index           =   2
         Left            =   6720
         TabIndex        =   25
         Top             =   35
         Width           =   3255
      End
      Begin VB.OptionButton optDate 
         Caption         =   "下一阶段提前至今天"
         Height          =   250
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   35
         Width           =   1935
      End
      Begin VB.OptionButton optDate 
         Caption         =   "正常进入下一天"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   35
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   120
         X2              =   10000
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   120
         X2              =   10000
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblDate 
         Caption         =   "时间进度"
         Height          =   230
         Left            =   120
         TabIndex        =   22
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   9855
      Begin VB.OptionButton optResult 
         Caption         =   "变异后结束(&3)"
         Height          =   250
         Index           =   3
         Left            =   7680
         TabIndex        =   4
         Top             =   20
         Width           =   1575
      End
      Begin VB.OptionButton optResult 
         Caption         =   "变异后退出(&2)"
         Height          =   250
         Index           =   2
         Left            =   5480
         TabIndex        =   3
         Top             =   20
         Width           =   1575
      End
      Begin VB.OptionButton optResult 
         Caption         =   "不符合(变异后继续)"
         Height          =   250
         Index           =   1
         Left            =   2920
         TabIndex        =   2
         Top             =   20
         Width           =   1935
      End
      Begin VB.OptionButton optResult 
         Caption         =   "符合(正常)"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   20
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   120
         X2              =   10000
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   120
         X2              =   10000
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Label lblResult 
         Caption         =   "总体结果"
         Height          =   230
         Left            =   120
         TabIndex        =   14
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.Frame fraRemark 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   7560
      TabIndex        =   20
      Top             =   3800
      Width           =   2295
      Begin VB.TextBox txtRemark 
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   650
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPersonnel 
         Height          =   1305
         Left            =   0
         TabIndex        =   10
         Top             =   2895
         Width           =   2415
         _cx             =   4260
         _cy             =   2302
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEvaluate.frx":6FE2
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         OwnerDraw       =   0
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblRemark 
         Caption         =   "2012-12-12评估备注(&R)"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraVariation 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   120
      TabIndex        =   18
      Top             =   3800
      Width           =   7335
      Begin VB.TextBox txtVariation 
         Height          =   300
         Left            =   4245
         MaxLength       =   1000
         TabIndex        =   6
         Top             =   15
         Width           =   2970
      End
      Begin VSFlex8Ctl.VSFlexGrid vsVariation 
         Height          =   3855
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   7215
         _cx             =   12726
         _cy             =   6800
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluate.frx":701C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         OwnerDraw       =   0
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblVariation 
         Caption         =   "变异原因"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   3375
      End
      Begin VB.Label lblSearch 
         Caption         =   "查找(&F)"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9945
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8055
      Width           =   9945
      Begin VB.CommandButton cmdFee 
         Caption         =   "评估费用(&F)"
         Height          =   350
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdTurn 
         Caption         =   "路径跳转(&T)"
         Height          =   350
         Left            =   1440
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdMergeOver 
         Caption         =   "结束合并路径(&M)"
         Height          =   350
         Left            =   2760
         TabIndex        =   33
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8760
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7560
         TabIndex        =   11
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00EFF0E0&
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   215
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsCriterion 
      Height          =   2130
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9735
      _cx             =   17171
      _cy             =   3757
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEvaluate.frx":7081
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9945
      TabIndex        =   15
      Top             =   0
      Width           =   9945
      Begin VB.Label lblNoteOne 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label lblPathTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "路径表名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "导入说明或阶段评估说明"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   400
         Width           =   8895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   120
         Picture         =   "frmEvaluate.frx":7108
         Top             =   45
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmEvaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mlngFun As Long '0-导入评估,1-阶段评估
Private mlngState As Byte '0-查看(导入评估),1-评估,2-修改(阶段评估)
                            '导入评估后不提供修改，要改只能取消导入，重新导入。阶段评估的查看通过调整功能实现,暂不提供取消功能

Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati
Private mstrPath As String '当前导入的路径表名称
Private mlngDiagnosisType As Long '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
Private mlngDiagnosisSorce As Long '诊断来源1-病历；2-入院登记；3-首页整理;4-病案
Private mlng疾病ID As Long
Private mlng诊断ID As Long
Private mlngType As Long    '=1 合并路径 =0 首要路径
Private mlng首要路径记录ID As Long    '首要路径记录的ID
Private mlngMergeID As Long          '查看合并路径的导入评估
Private mbln补录评估 As Boolean    '=True 是补录评估,False=非补录评估
Public mblnImp As Boolean                  '诊断符合时是否允许不导入

'模块变量
Private mrsCondition As ADODB.Recordset
Private mbln项目评估结果 As Boolean
Private mcol As Collection

Private Enum CNAME
    c序号 = 0
    c内容 = 1
End Enum

Private Enum CONST_COL_变异原因
    col变异分类 = 0
    col变异原因 = 1
    col变异选择 = 2
End Enum

Private mblnOK As Boolean
Private mobjParent As Object
Private mcolSQL As New Collection
Private mrsMerge As Recordset
Private mblnPathSend As Boolean


Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, ByVal lngState As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    Optional strPath As String, Optional lngDiagnosisType As Long, Optional lngDiagnosisSorce As Long, Optional ByVal lng疾病ID As Long, _
    Optional ByVal lng诊断ID As Long, Optional ByVal lngType As Long, Optional ByVal lng首要路径记录ID As Long, Optional ByVal lngMergeID As Long, _
    Optional ByVal bln补录 As Boolean = False) As Boolean
    '参数:bln补录  -True 补录评估
    
    mlngFun = lngFun
    mlngState = lngState
    mPati = t_pati
    mlng首要路径记录ID = lng首要路径记录ID
    mPP = t_pp
    mstrPath = strPath  '导入时传入
    mlngDiagnosisType = lngDiagnosisType '导入时传入
    mlngDiagnosisSorce = lngDiagnosisSorce '导入时传入
    mlng疾病ID = lng疾病ID
    mlng诊断ID = lng诊断ID
    mlngType = lngType
    mlngMergeID = lngMergeID
    mbln补录评估 = bln补录
    
    Set mobjParent = frmParent
        
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetCondition(lng评估ID As Long) As ADODB.Recordset
'功能：获取路径评估条件
    Dim strSql As String
    
    On Error GoTo errH
    If mlngFun = 0 Then
        strSql = "Select a.指标ID,a.关系式, a.条件值, a.条件组合" & vbNewLine & _
                "From 路径评估条件 A" & vbNewLine & _
                "Where a.评估ID = [1]"
        Set GetCondition = zlDatabase.OpenSQLRecord(strSql, "读取指标条件", lng评估ID)
    Else
        strSql = "Select a.指标ID, a.关系式, a.条件值, a.条件组合, Nvl(a.项目ID,0) as 项目ID,Nvl(b.执行结果,'无结果') as 执行结果,B.项目内容 " & vbNewLine & _
            "From 路径评估条件 A, (Select A.项目ID, A.执行结果, B.项目内容 From 病人路径执行 A,临床路径项目 B" & vbNewLine & _
            "   Where A.路径记录ID = [2] And A.阶段ID = [3] And A.天数 = [4] And A.项目Id = B.Id) B" & vbNewLine & _
            "Where a.项目ID = b.项目ID(+) And a.评估ID = [1]"
            
        Set GetCondition = zlDatabase.OpenSQLRecord(strSql, "读取指标条件", lng评估ID, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetCriterion() As ADODB.Recordset
'功能：获取导入评估和阶段评估的指标定义
    Dim strSql As String
 
    strSql = "Select a.ID 评估ID, b.ID 指标ID,b.序号, b.评估指标, b.指标结果,b.指标类型" & vbNewLine & _
            "From 临床路径评估 A, 路径评估指标 B" & vbNewLine & _
            "Where a.路径id = [1] And a.版本号 = [2] And a.Id = b.评估id And a.评估类型 = [3]" & IIf(mlngFun = 1, " And a.阶段id = [4]", "") & vbNewLine & _
            "Order by 序号"
    On Error GoTo errH
    Set GetCriterion = zlDatabase.OpenSQLRecord(strSql, "读取路径指标", mPP.路径ID, mPP.版本号, mlngFun + 1, mPP.当前阶段ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiCriterion() As ADODB.Recordset
'功能：mlngfun=0读取病人路径导入评估结果,
'      mlngfun=1修改阶段评估
    Dim strSql As String
    
    On Error GoTo errH
    If mlngFun = 0 Then
        If mlngMergeID = 0 Then
            strSql = "Select a.导入说明,a.未导入原因, a.状态,b.评估指标, b.指标结果" & vbNewLine & _
                    "From 病人临床路径 A, 病人路径指标 B" & vbNewLine & _
                    "Where a.id = [1] And a.id = b.路径记录id(+) And b.评估类型(+)=1"
        Else
            strSql = "Select a.导入说明,Null as 未导入原因,1 as 状态,b.评估指标, b.指标结果" & vbNewLine & _
                    "From 病人合并路径 A, 病人路径指标 B" & vbNewLine & _
                    "Where a.id = [2] And a.id = b.合并路径记录ID(+) And b.评估类型(+)=1"
        End If
        Set GetPatiCriterion = zlDatabase.OpenSQLRecord(strSql, "读取病人路径指标", mPP.病人路径ID, mlngMergeID)
    Else
        strSql = "Select a.评估结果,a.变异原因, Nvl(a.时间进度,0) as 时间进度, a.评估说明,a.评估人,b.评估指标,b.指标结果" & vbNewLine & _
                "From 病人路径评估 A, 病人路径指标 B" & vbNewLine & _
                "Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3]" & vbNewLine & _
                "And a.路径记录id = b.路径记录id(+) And a.阶段id=b.阶段id(+) And a.日期=b.日期(+) And b.评估类型(+)=2"
    
        Set GetPatiCriterion = zlDatabase.OpenSQLRecord(strSql, "读取病人路径指标", mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetMoneyInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的剩余额
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    strSql = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额" & _
            " From 病人余额 Where 性质=1 And 类型 = 2 And 病人ID=[1] " & _
            " Union All Select -1*Nvl(Sum(金额),0) as 费用余额,0 as 预交余额" & _
            " From 保险模拟结算" & _
            " Where 病人ID=[1] And 主页ID=[2]"
    strSql = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSql & ")"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdFee_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strInfo As String, strThisInfo As String, lngDay As Long, lngLen As Long, DatIn As Date
    Dim lng阶段ID As Long, str阶段名称 As String
    Dim cur金额 As Currency, cur金额合计 As Currency
    
    Set rsTmp = GetMoneyInfo(mPati.病人ID, mPati.主页ID)
    If Not rsTmp Is Nothing Then
        strInfo = "未结费用：" & Format(rsTmp!费用余额, "0.00") & ",预交余额：" & Format(rsTmp!预交余额, "0.00")
    End If
    
    If mPP.当前阶段分支ID = 0 Then
        strSql = "Select 标准住院日 From 临床路径版本 Where 路径id = [1] And 版本号 = [2]"
    Else
        strSql = "Select 标准住院日 From 临床路径分支 Where 路径id = [1] And 版本号 = [2] And ID=[3]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID, mPP.版本号, mPP.当前阶段分支ID)
    If InStr(rsTmp!标准住院日, "-") > 0 Then
        lngLen = Split(rsTmp!标准住院日, "-")(0)
    Else
        lngLen = Val(rsTmp!标准住院日)
    End If
    strInfo = strInfo & vbCrLf & "按标准住院日" & lngLen & "天估算，即将发生的费用(不含可选项目)："
    DatIn = GetPatiInPath(mPati, mPP.病人路径ID)
    
    For lngDay = mPP.当前天数 + 1 To lngLen
        lng阶段ID = GetPhaseByDay(mPP.路径ID, mPP.版本号, lngDay, str阶段名称)
        
        cur金额 = GetChargeOfDay(lng阶段ID, lngDay, DatIn)
        cur金额合计 = cur金额合计 + cur金额
       
        strThisInfo = "第" & lngDay & "天：" & IIf(lngDay < 10, Space(2), "") & Format(cur金额, "0.00")
        
        If lngLen > 10 And (lngDay Mod 2) = 0 And lngDay <> mPP.当前天数 + 1 Then
            strInfo = strInfo & vbTab & vbTab & strThisInfo
        Else
            strInfo = strInfo & vbCrLf & strThisInfo
        End If
    Next
    strInfo = strInfo & vbCrLf & "共计：" & Space(4) & Format(cur金额合计, "0.00")
    MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPhaseByDay(ByVal lng路径ID As Long, ByVal lng版本号 As Long, ByVal lng天数 As Long, str阶段名称 As String) As Long
'功能：获取指定天数对应的缺省阶段ID
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ID,名称" & vbNewLine & _
            "From 临床路径阶段" & vbNewLine & _
            "Where 路径id = [1] And 版本号 = [2] And" & vbNewLine & _
            "      (([3] Between 开始天数 And 结束天数) Or (开始天数 = [3] And 结束天数 Is Null) Or (开始天数 Is Null And 结束天数 Is Null))" & vbNewLine & _
            "Order By Decode(开始天数, Null, 1, 0),序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, lng版本号, lng天数)
    GetPhaseByDay = rsTmp!ID
    str阶段名称 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetChargeOfDay(ByVal lng阶段ID As Long, ByVal lng天数 As Long, ByVal DatIn As Date) As Long
'功能：获取指定天数对应的缺省阶段ID
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Zl_Getpathcharge([1],[2],[3],[4],[5],[6],[7]) as 金额 From dual"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.主页ID, mPP.路径ID, mPP.版本号, lng阶段ID, lng天数, DatIn)
    GetChargeOfDay = Val("" & rsTmp!金额)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakeMergePath() As ADODB.Recordset
    Set MakeMergePath = New ADODB.Recordset
    
    MakeMergePath.Fields.Append "Id", adBigInt
    MakeMergePath.Fields.Append "路径id", adBigInt, , adFldIsNullable
    MakeMergePath.Fields.Append "版本号", adBigInt, , adFldIsNullable
    MakeMergePath.Fields.Append "名称", adVarChar, 100, adFldIsNullable
    
    MakeMergePath.Fields.Append "编码", adVarChar, 50, adFldIsNullable
    MakeMergePath.Fields.Append "说明", adVarChar, 500, adFldIsNullable
    MakeMergePath.Fields.Append "状态", adBigInt, , adFldIsNullable   '=0未达到标准住院日,=1达到标准住院日，但未到最后一天，=2当前为标准住院日最后一天
    MakeMergePath.Fields.Append "选择", adBigInt, , adFldIsNullable   '=1选中
    MakeMergePath.Fields.Append "显示", adBigInt, , adFldIsNullable   '=1显示
    MakeMergePath.Fields.Append "允许修改", adBigInt, , adFldIsNullable   '0=允许，=1不允许
    
    MakeMergePath.CursorLocation = adUseClient
    MakeMergePath.LockType = adLockOptimistic
    MakeMergePath.CursorType = adOpenStatic
    MakeMergePath.Open
End Function

Private Sub cmdMergeOver_Click()
'功能：选择需要完成的合并路径
    Dim lngCount As Long
    Dim rsTmp As Recordset
    Dim objPathImport As New frmPathImport
    Dim t_pp As TYPE_PATH_Pati
    '为解决导入时出现多个路径表的问题，
    '原因：直接用frmPathImport.showme相当于用的全局对象，而这里new一个对象来使用，则是避免和导入时使用的导入窗体出现冲突（这里是评估调导入窗体，而导入时是导入窗体调用评估窗体）。
    
    If mrsMerge.RecordCount > 0 Then
        mrsMerge.MoveFirst
        Set rsTmp = zlDatabase.CopyNewRec(mrsMerge)
        Do While Not mrsMerge.EOF
            '如果达到标准住院日，则显示出来
            If Val(mrsMerge!状态 & "") = 1 Or Val(mrsMerge!状态 & "") = 2 Or optDate(1).Value Or optDate(1).Enabled = False Then
                mrsMerge!显示 = 1
            Else
                mrsMerge!显示 = 0
            End If
            
            '如果达到最后一天，且为没有选择下一阶段延后，则必须勾选
            If Val(mrsMerge!状态 & "") = 2 And Not optDate(2).Value Then
                mrsMerge!选择 = 1
                mrsMerge!允许修改 = 1
            Else
                mrsMerge!允许修改 = 0
            End If
            mrsMerge.MoveNext
        Loop
        mrsMerge.MoveFirst
        If Not objPathImport.ShowMe(Me, mPati, 4, t_pp, , , , , , mrsMerge, True) Then
            '还原信息
            Set mrsMerge = rsTmp
            Exit Sub
        End If
        Do While Not mrsMerge.EOF
            If Val(mrsMerge!选择 & "") = 1 Then lngCount = lngCount + 1
            mrsMerge.MoveNext
        Loop
        mrsMerge.MoveFirst
        lblMSG.Caption = "已选择" & lngCount & "个合并路径。"
    Else
        MsgBox "没有可结束的合并路径。", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdTurn_Click()
    Dim lngPathID As Long, lngPathVersion As Long
    Dim str审核人 As String, blnTrnAduit As Boolean
    Dim objPathImport As New frmPathImport
    Dim t_pp As TYPE_PATH_Pati
    
    If InStr(cmdTurn.Tag, ",") > 0 Then
        lngPathID = Split(cmdTurn.Tag, ",")(0)
        lngPathVersion = Split(cmdTurn.Tag, ",")(1)
    End If
    
    
    If objPathImport.ShowMe(Me, mPati, 1, t_pp, mPP.路径ID, lngPathID, lngPathVersion) Then
        '如果没有权限的话先检查前面阶段是否有跳转但未审核的天数
        If InStr(GetInsidePrivs(p临床路径应用), ";跳转审核;") = 0 Then
            If CheckPathIsTurnAduit Then
                str审核人 = zlDatabase.UserIdentify(Me, "前面阶段存在未审核的路径跳转，必须审核后才允许继续。", glngSys, p临床路径应用, "跳转审核")
                If str审核人 = "" Then Exit Sub
                blnTrnAduit = True
            Else
                str审核人 = zlDatabase.UserIdentify(Me, "你没有跳转审核的权限，是否现在就进行审核？", glngSys, p临床路径应用, "跳转审核")
            End If
        Else
            str审核人 = UserInfo.姓名
        End If
        cmdTurn.Tag = lngPathID & "," & lngPathVersion & "," & str审核人 & "," & IIf(blnTrnAduit, "1", "0")
        Call cmdOK_Click
    Else
        cmdTurn.Tag = ""
    End If
End Sub

Private Function CheckPathIsTurnAduit() As Boolean
'功能：检查是否存在未审核的跳转阶段。true为存在
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From 病人路径评估 Where 原路径id is not null And 跳转审核人 is null And 路径记录ID=[1]"
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "跳转审核", mPP.病人路径ID)
     
     CheckPathIsTurnAduit = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If vsCriterion.Visible And vsCriterion.Enabled And vsCriterion.Rows > vsCriterion.FixedRows Then
        vsCriterion.SetFocus
    Else
        If txtRemark.Visible And txtRemark.Enabled Then txtRemark.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub InitFace()
'功能：初始化界面设置
    Dim i As Integer, lngMin As Long, lngMax As Long, lng住院天数 As Long
    Dim strPathin As String, strPatientin As String
    Dim strSql As String, rsTmp As Recordset
    Dim lngState As Long
        
    '1.界面基本初始
    
    fraVariation.BackColor = Me.BackColor
    fraRemark.BackColor = Me.BackColor
    fraDate.BackColor = Me.BackColor
    fraResult.BackColor = Me.BackColor
    fraStart.BackColor = Me.BackColor
    
    lblResult.Tag = "不处理"
    fraStart.Visible = mlngFun = 0
    fraDate.Visible = mlngFun = 1
    lblNoteOne.Visible = False
    
    If mlngFun = 0 Then
        fraStart.Top = fraDate.Top
        If mlngState = 0 Then
        '查看
            strPathin = Format(GetPatiInPath(mPati, mPP.病人路径ID), "YYYY-MM-DD")
            strPatientin = Format(GetPatiInDate(mPati, lng住院天数), "YYYY-MM-DD")
            If strPathin = strPatientin Then
                optStart(0).Value = True
            Else
                optStart(1).Value = True
            End If
            
            optStart(0).Caption = "入科时间作为入径第1天"
            optStart(1).Caption = "以导入当天作为入径第1天"
            optStart(0).Enabled = False
            optStart(1).Enabled = False
        Else
        '不是当天入院的，允许选择起始时间
            Call GetPatiInDate(mPati, lng住院天数)
            If lng住院天数 <= 1 Or mblnPathSend Then
                fraStart.Visible = False
                fraStart.Tag = "不可见"
                vsCriterion.Height = vsCriterion.Height + fraStart.Height
            Else
                optStart(0).Caption = "入科时间作为入径第1天,目前是入径第" & lng住院天数 & "天"
            End If
        End If
        
        If mlngType = 1 And fraStart.Tag <> "不可见" Then
            '合并路径只能以当前天数作为第一天
            fraStart.Visible = False
            fraStart.Tag = "不可见"
            vsCriterion.Height = vsCriterion.Height + fraStart.Height
        End If
        If mlngType = 1 Then txtRemark.Enabled = False: lblVariation.Caption = "合并路径不填写不符合原因。"
        
        Me.Caption = "导入评估"
        lblResult.Caption = "导入结果"
        lblPathTitle.Caption = "导入路径表：" & mstrPath
        lblNote.Caption = "请选择导入结果，如果不符合导入条件，请选择原因并填写说明，以便后续统计分析。"
        lblRemark.Caption = "备注(&R)"
        optResult(0).Caption = "符合(&0)": optResult(1).Caption = "不符合(&1)"
        If (Not mblnImp) And mlngFun = 0 Then
            optResult(1).Enabled = False
        End If
        optResult(2).Visible = False
        optResult(3).Visible = False
        
        cmdFee.Visible = False
        vsPersonnel.Visible = False
        txtRemark.Height = fraRemark.Height - lblRemark.Height - 60
        lblRemark.Top = 0
        txtRemark.Top = lblRemark.Top + lblRemark.Height + 30
        
        cmdMergeOver.Visible = False
        
        If mlngState = 0 Then   '查看
            optResult(0).Enabled = False
            optResult(1).Enabled = False
            txtRemark.Enabled = False
            vsVariation.Enabled = False
            txtVariation.Enabled = False
            
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
        Else
            If mlngType = 0 Then
                cmdOK.Left = cmdCancel.Left
                cmdCancel.Visible = False
            End If
        End If
    Else
        cmdTurn.Visible = mlngState = 1 '评估时选择了跳转路径后，已限制不允许修改，只能取消评估
        lblPathTitle.Visible = False
        lblNote.Top = lblPathTitle.Top
        lblNote.Height = 400
        lblNote.Caption = "请根据病人的当前情况进行评估，以决定是否继续按照路径表制定的计划进行后续工作。如果发生了变异，请选择变异原因，并填写变异说明，以便后续统计分析和持续改进路径表。"
        If mbln补录评估 Then
            '由于提前生成导致当前日期之前补录评估时,只能选择正常进入
            optDate(0).Value = True: optDate(1).Enabled = False: optDate(2).Enabled = False: optDate(3).Enabled = False
            
            lblNoteOne.Visible = True
            lblNoteOne.Top = lblNote.Top + lblNote.Height
            lblNoteOne.Caption = "当前阶段之后已生成其他项目,要想结束路径需取消提前生成的路径项目后再重新评估。"
            lblNoteOne.ForeColor = vbRed
        Else
            If GetNextPhase(mPP.当前阶段ID, mPP.当前阶段分支ID) = 0 Then '没有后续阶段时，不允许选择下一阶段提前
                If optDate(1).Value Then optDate(1).Value = False    '提前至今天
                optDate(1).Enabled = False
                If optDate(3).Value Then optDate(3).Value = False    '提前至明天
                optDate(3).Enabled = False
            End If
        End If
        
        If mPP.合并路径个数 = 0 Or mlngState <> 1 Then
            cmdMergeOver.Visible = False
        Else
            '排除结束了的合并路径
            strSql = "Select a.Id, a.路径id, b.名称, b.编码, b.说明,c.版本号,a.当前阶段ID,d.分支ID as 当前阶段分支ID,a.当前天数" & vbNewLine & _
                    "From 病人合并路径 A, 临床路径目录 B, 临床路径版本 C,临床路径阶段 D" & vbNewLine & _
                    "Where a.路径id = b.Id And a.路径id = c.路径id And a.版本号 = c.版本号 And d.id(+)=a.当前阶段id " & _
                    "  And a.首要路径记录id = [1] and a.结束时间 is Null"

            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
            If rsTmp.RecordCount = 0 Then cmdMergeOver.Visible = False
            If rsTmp.RecordCount > 0 Then
                Set mrsMerge = MakeMergePath
                Do While Not rsTmp.EOF
                    mrsMerge.AddNew
                    mrsMerge!ID = rsTmp!ID
                    mrsMerge!路径ID = Val(rsTmp!路径ID & "")
                    mrsMerge!版本号 = Val(rsTmp!版本号 & "")
                    mrsMerge!名称 = rsTmp!名称 & ""
                    mrsMerge!编码 = rsTmp!编码 & ""
                    Call IsLastDate(, , , Val(rsTmp!路径ID & ""), Val(rsTmp!版本号 & ""), Val(rsTmp!当前阶段ID & ""), Val(rsTmp!当前阶段分支ID & ""), Val(rsTmp!当前天数 & ""), True, lngState, Val(rsTmp!ID & ""))
                    mrsMerge!状态 = lngState
                    
                    rsTmp.MoveNext
                Loop
                If mrsMerge.RecordCount > 0 Then mrsMerge.Update: mrsMerge.MoveFirst
            End If
        End If
        
        lblRemark.Caption = Format(mPP.当前日期, "YYYY-MM-DD") & "评估备注"
        optResult(0).Caption = "正常(&0)"
        optResult(1).Caption = "变异后继续(&1)"
        optResult(2).Caption = "变异后退出(&2)"
        If mbln补录评估 Then
            optResult(2).Enabled = False
            optResult(3).Enabled = False
            cmdTurn.Enabled = False
        Else
            '达到标准住院日天数即可变异后完成。
            '如果没有达到标准住院日，则提供一个选项：提前结束
            If IsLastDate(True, lngMin, lngMax, mPP.路径ID, mPP.版本号, mPP.当前阶段ID, mPP.当前阶段分支ID, mPP.当前天数) Then
                optResult(3).Visible = True
                optResult(3).Caption = "变异后完成(&3)"
            Else
                optResult(3).Caption = "提前完成(&3)"
                optResult(3).Visible = True
            End If
            If InStr(GetInsidePrivs(p临床路径应用), ";结束路径;") = 0 Then
                optResult(2).Visible = False
                optResult(3).Visible = False
            Else
                '89620:具备结束路径和提前完成才允许提前完成,否则禁止。
                If optResult(3).Caption = "提前完成(&3)" Then
                    optResult(3).Visible = (InStr(GetInsidePrivs(p临床路径应用), ";提前完成;") > 0)
                End If
            End If
            
            '超过标准住院日后，不能选择正常
            If mPP.当前天数 > lngMax Then
                optResult(1).Value = True
                optResult(0).Enabled = False: optResult(0).Tag = "禁止选择正常"
            End If
        End If
    End If
    
    '2.评估指标表初始(其中有对评估结果的设置)
    Call InitVsCriterion
                
    lblResult.Tag = ""
    '3.加载变异原因列表
    For i = 0 To optResult.count - 1
        If optResult(i).Value Then
            Exit For
        End If
    Next
    Call optResult_Click(i)
        
    '4.初始化评估人
    If mlngFun = 1 Then
    
        With vsPersonnel
            .Redraw = flexRDNone
            .Editable = flexEDKbdMouse
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = "评估人"
            
            .Rows = 2
            .TextMatrix(1, 0) = UserInfo.姓名  '缺省为当前操作员
            .Redraw = True
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitVsCriterion()
'功能：初始化评估指标表
    Dim strcol As String, arrHead As Variant
    Dim i As Long, lng评估ID As Long
    Dim rsCriterion As ADODB.Recordset
    Dim blnValue As Boolean, blnThis As Boolean
    
    lng评估ID = 0
    Set rsCriterion = GetCriterion
    If rsCriterion.RecordCount > 0 Then
        lng评估ID = rsCriterion!评估ID
        Set mrsCondition = GetCondition(lng评估ID)
        
        strcol = "序号,450,4;评估指标,6800,1;结果,900,1;指标类型;指标结果"
        '1.初始化评估指标表头
        With vsCriterion
            .Redraw = flexRDNone
            .Clear
            .FixedCols = 1: .FixedRows = 1
            arrHead = Split(strcol, ";")
            .Cols = UBound(arrHead) + 1
            .Rows = .FixedRows
            .Rows = .FixedRows + rsCriterion.RecordCount
            .Editable = flexEDKbdMouse
            Set mcol = New Collection
            
            For i = 0 To UBound(arrHead)
                mcol.Add i, Split(arrHead(i), ",")(0)
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
                
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                End If
            Next
            
            '2.加载指标列表
            For i = 1 To rsCriterion.RecordCount
                .RowData(i) = Val(rsCriterion!指标ID)
                .TextMatrix(i, mcol("序号")) = rsCriterion!序号
                .TextMatrix(i, mcol("评估指标")) = rsCriterion!评估指标
                .TextMatrix(i, mcol("结果")) = Split(rsCriterion!指标结果, vbTab)(1)
                
                .TextMatrix(i, mcol("指标类型")) = rsCriterion!指标类型
                .TextMatrix(i, mcol("指标结果")) = rsCriterion!指标结果
                
                rsCriterion.MoveNext
            Next
            
            .Redraw = flexRDDirect
        End With
    
        '3.如果设置了指标条件，根据指标条件设置缺省的评估结果
        If mlngState = 1 And mrsCondition.RecordCount > 0 Then
            If mlngFun = 0 Then
                Call SetResult
            
            '阶段评估的异常结果加载，当与指标不符时，缺省以执行结果为异常说明
            ElseIf mlngFun = 1 Then
                With mrsCondition
                    blnValue = False
                    
                    .Filter = "项目ID<>0"
                    For i = 1 To .RecordCount
                        Select Case !关系式
                            Case "="
                                blnThis = (!执行结果 = !条件值)
                            Case "<>"
                                blnThis = (!执行结果 <> !条件值)
                            Case ">"
                                blnThis = (!执行结果 > !条件值)
                            Case ">="
                                blnThis = (!执行结果 >= !条件值)
                            Case "<"
                                blnThis = (!执行结果 < !条件值)
                            Case "<="
                                blnThis = (!执行结果 <= !条件值)
                            Case "Like"
                                blnThis = (!执行结果 Like "*" & !条件值 & "*")
                            Case Else
                                blnThis = True
                        End Select
                                        
                        If i = 1 Then
                            blnValue = blnThis
                        Else
                            If !条件组合 = 1 Then
                                blnValue = (blnValue And blnThis)
                            Else
                                blnValue = (blnValue Or blnThis)
                            End If
                        End If
                        
                        .MoveNext
                    Next
                    mbln项目评估结果 = blnValue
                    
                    If blnValue Or optResult(0).Enabled = False Then '阶段评估，满足条件时表示变异
                        optResult(1).Value = True   '缺省为变异后继续
                
                        '如果项目执行结果都符合条件，再检查指标内容是否符合
                        Call SetResult
                    Else
                        optResult(0).Value = True
                    End If
                End With
            End If
        End If
    Else
        '没有评估指标时，不显示指标表格
        vsCriterion.Tag = "没有评估指标记录"
        vsCriterion.Visible = False
        If mlngFun = 0 Then
            fraStart.Top = vsCriterion.Top
            fraResult.Top = fraStart.Top + IIf(fraStart.Tag = "不可见", 0, fraStart.Height + 30)
        Else
            fraDate.Top = vsCriterion.Top
            fraResult.Top = fraDate.Top + fraDate.Height + 30
        End If
        fraVariation.Top = fraResult.Top + fraResult.Height
        fraRemark.Top = fraVariation.Top
                
        Me.Height = Me.Height - vsCriterion.Height - 120
    End If
End Sub

Private Sub InitVariation(ByVal lngKind As Long)
'功能：初始化变异原因列表
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    strSql = "Select b.名称 As 分类, a.编码, a.名称, a.简码" & vbNewLine & _
            "From 变异常见原因 A, 变异常见原因 B" & vbNewLine & _
            "Where a.末级 = 1 And a.上级 = b.编码 and a.性质=[1]" & vbNewLine & _
            "order by 分类,编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngKind)
        
    With vsVariation
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If rsTmp.RecordCount > 0 Then
            .MergeCol(col变异分类) = True
            .Rows = .FixedRows + rsTmp.RecordCount
            '缺省不选择
            Set .Cell(flexcpPicture, .FixedRows, col变异选择, .Rows - 1, col变异选择) = imgNature.ListImages(IIf(mlngFun = 0, "UnSelected", "UnCheck")).Picture
            .Cell(flexcpPictureAlignment, .FixedRows, col变异选择, .Rows - 1, col变异选择) = flexPicAlignCenterCenter

            For i = .FixedRows To rsTmp.RecordCount
                .Cell(flexcpData, i, col变异选择) = 0
                
                .RowData(i) = CStr(rsTmp!编码)    '主键
                .TextMatrix(i, col变异分类) = rsTmp!分类
                .TextMatrix(i, col变异原因) = rsTmp!编码 & "-" & rsTmp!名称
                .Cell(flexcpData, i, col变异原因) = "" & rsTmp!简码
                rsTmp.MoveNext
            Next
        End If
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    vsVariation.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnPathSend = CheckPathSend(mPati.病人ID, mPati.主页ID)
    mblnImp = InStr(GetInsidePrivs(p临床路径应用), ";诊断符合允许不导入;") > 0
    Call InitFace
    Call LoadData
    
End Sub

Private Sub SetFillTableByStr(vstmp As VSFlexGrid, strTmp As String, lngCol As Long)
'功能：将字符串的值按分隔符填充到表格中，并且，在未尾新增一空行
'参数：
    Dim i As Long, arrtmp As Variant
    
    arrtmp = Split(strTmp, ",")
    With vstmp
        .Rows = .FixedRows + UBound(arrtmp) + 2 '最后加一行空行
        For i = 0 To UBound(arrtmp)
            .TextMatrix(i + .FixedRows, lngCol) = arrtmp(i)
        Next
        .TextMatrix(.Rows - 1, lngCol) = ""
    End With
End Sub

Private Function Get项目变异原因() As ADODB.Recordset
'功能：获取路径外项目的变异原因
    Dim strSql As String
    If mlngState = 1 Then
        strSql = "Select distinct 变异原因 From (Select 变异原因 From 病人路径执行 " & _
                "Where 路径记录Id = [1] And 阶段ID = [2] And 日期 = [3] And 变异原因 Is Not Null And Nvl(生成时间性质,0)<2 Order by 登记时间)"
    ElseIf mlngState = 2 Then
        strSql = "Select 变异原因 From 病人路径变异 Where 路径记录Id = [1] And 阶段ID = [2] And 日期 = [3] "
    End If
    On Error GoTo errH
    Set Get项目变异原因 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadData()
'功能：加载数据
    Dim i As Long, str变异原因 As String
    Dim j As Long
    Dim rsTmp As ADODB.Recordset
                
    If mlngFun = 1 Then
        Set rsTmp = Get项目变异原因
        If rsTmp.RecordCount > 0 Then
            optResult(0).Enabled = False: optResult(0).Tag = "禁止选择正常"
            optResult(1).Value = True   '变异后继续
        End If
                
        If rsTmp.RecordCount > 0 Then
            For j = 1 To rsTmp.RecordCount
                i = vsVariation.FindRow(CStr(rsTmp!变异原因)) '按编码查找rowdata
                If i > 0 Then
                    vsVariation.Row = i
                    vsVariation.TopRow = i
                    Call vsVariation_Click
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
                           
    '1.评估指标结果
    '查看导入评估的数据
    '评估修改时读出原有评估结果，可能有指标，也可能没有指标
    If mlngFun = 0 And mlngState = 0 Or (mlngFun = 1 And mlngState = 2) Then
        Set rsTmp = GetPatiCriterion
        '一定有记录
        If mlngFun = 0 Then
            optResult(0).Value = rsTmp!状态 = 1
            optResult(1).Value = rsTmp!状态 <> 1
            
            If Not IsNull(rsTmp!未导入原因) Then
                i = vsVariation.FindRow(CStr(rsTmp!未导入原因)) '按编码查找rowdata
                If i > 0 Then
                    vsVariation.Row = i
                    vsVariation.TopRow = i
                    Call vsVariation_Click
                End If
            End If
            txtRemark.Text = "" & rsTmp!导入说明
        Else
            If rsTmp!时间进度 = -1 Then
                optDate(2).Value = True     '调用click事件，设置关联的optResult的可用性
            ElseIf rsTmp!时间进度 = 1 Then
                optDate(1).Value = True
            ElseIf rsTmp!时间进度 = 2 Then
                optDate(3).Value = True
            Else
                optDate(0).Value = True
            End If
            
            If rsTmp!评估结果 = -1 Then
                If mPP.病人路径状态 = 1 Then
                    optResult(1).Value = True   '变异并继续（变异后结束须先取消结束，此时相当于变异后继续）
                Else
                    optResult(2).Value = True   '变异退出
                End If
            Else
                optResult(0).Value = True
            End If
            
            txtRemark.Text = "" & rsTmp!评估说明
            Call SetFillTableByStr(vsPersonnel, rsTmp!评估人, 0)
        End If
            
        '加载指标结果
        If vsCriterion.Tag <> "没有评估指标记录" Then
            With vsCriterion
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    rsTmp.Filter = "评估指标='" & .TextMatrix(i, mcol("评估指标")) & "'"
                    If rsTmp.RecordCount > 0 Then
                        .TextMatrix(i, mcol("结果")) = "" & rsTmp!指标结果
                    Else
                        .TextMatrix(i, mcol("结果")) = ""
                    End If
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngFun = 0 And mblnOK = False And mlngType = 0 And mlngState <> 0 Then
        '导入评估时，隐藏了取消按钮，只能点确定
        Cancel = 1
        Exit Sub
    End If
    
    mbln项目评估结果 = False
    Set mrsCondition = Nothing
    Set mobjParent = Nothing
    Set mcolSQL = Nothing
End Sub

Private Sub optDate_Click(Index As Integer)
    If Index = 0 Then
        If optResult(0).Tag <> "禁止选择正常" Then optResult(0).Enabled = True
        optResult(1).Enabled = True
        optResult(2).Enabled = True
        optResult(3).Enabled = True
        If optResult(3).Caption = "提前完成(&3)" Then
            'optResult(3).Enabled = False
            If optResult(3).Value Then optResult(0).Value = True
        End If
        
        cmdTurn.Enabled = optResult(1).Value And Not mbln补录评估
    Else
    '时间变异时，只能选择变异后继续。
        optResult(0).Enabled = False
        optResult(1).Enabled = True
        '如果选择时间提前，允许使用提前结束功能。
        If Index = 1 And optResult(3).Caption = "提前完成(&3)" Then
            optResult(3).Enabled = True
            If optResult(0).Value Or optResult(2).Value Then optResult(1).Value = True
        Else
            optResult(3).Enabled = False
            optResult(1).Value = True
        End If
        optResult(2).Enabled = False
        
        
        cmdTurn.Enabled = False
    End If
End Sub

Private Sub optResult_Click(Index As Integer)
    If lblResult.Tag = "不处理" Then Exit Sub
    
    If mlngFun = 0 Then '导入
        Call InitVariation(0)
    Else
        If Index = 1 Or Index = 3 Then '变异继续或结束
            Call InitVariation(1)
            If Index = 3 And optResult(3).Caption = "提前完成(&3)" Then
                optDate(1).Value = True
            End If
        ElseIf Index = 2 Then   '变异退出
            Call InitVariation(2)
        End If
        
        cmdTurn.Enabled = (Index = 1 And optDate(0).Value) And Not mbln补录评估
    End If
    
    '评估正常时禁止用变异原因,查看导入评估时也禁用
    If Index = 0 Or mlngState = 0 Or mlngType = 1 Then
        vsVariation.Enabled = False
        vsVariation.BackColor = Me.BackColor
        vsVariation.Row = 0
        txtVariation.Enabled = False
        txtVariation.BackColor = Me.BackColor
    Else
        vsVariation.Enabled = True
        vsVariation.BackColor = &H80000005
        txtVariation.Enabled = True
        txtVariation.BackColor = &H80000005
        
        If vsVariation.Visible And vsVariation.Enabled Then vsVariation.SetFocus
    End If
End Sub

Private Sub txtRemark_GotFocus()
    Call zlControl.TxtSelAll(txtRemark)
End Sub

Private Sub vsCriterion_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mcol("结果") And mlngState = 1 Then    '修改评估结果时不再根据指结果设置总体结果
        Call SetResult
    End If
End Sub

Private Sub SetResult()
'功能：根据指标项目的结果设置总体的结果
    Dim i As Long, j As Long, strValue As String
    Dim blnValue As Boolean, blnThis As Boolean
    Dim blnFirst As Boolean
        
    blnFirst = True
    If mlngFun = 1 Then
        blnValue = mbln项目评估结果
    Else
        blnValue = True
    End If
    For i = 1 To vsCriterion.Rows - 1
        strValue = vsCriterion.TextMatrix(i, mcol("结果"))
        If mlngFun = 0 Then
            mrsCondition.Filter = "指标ID = " & vsCriterion.RowData(i)
        Else
            mrsCondition.Filter = "指标ID = " & vsCriterion.RowData(i) & " And 项目ID = 0"
        End If
        With mrsCondition
            For j = 1 To .RecordCount
                 Select Case !关系式
                    Case "="
                        blnThis = (strValue = !条件值)
                    Case "<>"
                        blnThis = (strValue <> !条件值)
                    Case ">"
                        blnThis = (strValue > !条件值)
                    Case ">="
                        blnThis = (strValue >= !条件值)
                    Case "<"
                        blnThis = (strValue < !条件值)
                    Case "<="
                        blnThis = (strValue <= !条件值)
                    Case "Like"
                        blnThis = (strValue Like "*" & !条件值 & "*")
                    Case Else
                        blnThis = True
                End Select
                
                If blnFirst And mlngFun = 0 Then
                    blnValue = blnThis
                    blnFirst = False
                Else
                    If !条件组合 = 1 Then
                        blnValue = (blnValue And blnThis)
                    Else
                        blnValue = (blnValue Or blnThis)
                    End If
                End If
                .MoveNext
            Next
        End With
    Next
    
    If mlngFun = 0 Then
        If blnValue Then
            optResult(0).Value = True
        Else
            If mblnImp Then
                optResult(1).Value = True
            Else
                optResult(0).Value = True
            End If
        End If
    Else
        If blnValue Then  '阶段评估，满足条件时表示变异
            optResult(1).Value = True
        Else
            If optResult(0).Enabled Then optResult(0).Value = True  '选择进度延后或提前时，以及超过标准住院日，不能再选择“正常”
        End If
    End If
End Sub

Private Sub vsCriterion_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Visible Then
        If NewCol = mcol("结果") And mlngState <> 0 Then
            Dim arrtmp As Variant
            
            With vsCriterion
                arrtmp = Split(.TextMatrix(NewRow, mcol("指标结果")), vbTab)
                .ColComboList(NewCol) = Replace(arrtmp(0), ",", "|")
            End With
        End If
    End If
End Sub

Private Sub vsCriterion_GotFocus()
'    vsCriterion.ForeColorSel = vbWhite
'    vsCriterion.BackColorSel = &H8000000D
End Sub

Private Sub vsCriterion_LostFocus()
'    vsCriterion.ForeColorSel = vbBlack
'    vsCriterion.BackColorSel = vbWhite
End Sub

Private Sub vsCriterion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsCriterion)
    End If
End Sub

Private Sub vsCriterion_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mcol("结果") Or mlngState = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsPersonnel_GotFocus()
    If vsPersonnel.Row = vsPersonnel.Rows - 1 Then
        With vsPersonnel
            If .TextMatrix(.Row, .Col) <> "" Then
                Call vsPersonnel_AfterEdit(.Row, .Col)
            End If
        End With
    End If
End Sub

Private Sub vsPersonnel_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除最后一行，或清除单元格内容
    If KeyCode = vbKeyDelete Then
        With vsPersonnel
            If .Row = .Rows - 1 And .Row > .FixedRows And .TextMatrix(.Row, 0) = "" Then '至少留一行
                .Rows = .Rows - 1
            ElseIf .Row > .FixedRows - 1 Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End With
    End If
End Sub

Private Sub vsPersonnel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：最后一行回车后自动加一行
    With vsPersonnel
        If Trim(.TextMatrix(Row, Col)) <> "" And Row = .Rows - 1 Then
            .Rows = .Rows + 1
            .Select .Rows - 1, .Col
        End If
    End With
End Sub

Private Sub vsPersonnel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsPersonnel)
    End If
End Sub

Private Sub vsPersonnel_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strtxt As String, strSql As String, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim vPoint As POINTAPI
    
    With vsPersonnel
        strtxt = Trim(.EditText)
        If strtxt = "" Then Exit Sub
        
        If zlCommFun.IsCharAlpha(strtxt) Then
            strtxt = UCase(strtxt)
            strSql = " And a.简码 like [1]"
        Else
            strSql = " And a.姓名 like [1]"
        End If
        strSql = "Select Distinct a.ID,a.编号 as 编码,a.姓名 From 人员表 a, 人员性质说明 b Where a.Id = b.人员id And b.人员性质 In ('医生', '护士')" & strSql
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "评估人", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, False, strtxt & "%")
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "根据输入内容未找到匹配的医生或护士。", vbInformation, gstrSysName
            End If
            Cancel = True
            Exit Sub
        End If
        For i = .FixedCols To .Rows - 1
            If .TextMatrix(i, 0) = rsTmp!姓名 And i <> .Row Then
                MsgBox "已经输入了相同姓名的人员。", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        Next
        
        .EditText = rsTmp!姓名
    End With
End Sub

Private Sub ResultEnterNextCell(vsthis As VSFlexGrid)
    With vsthis
        If .Col <= .Cols - 1 Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    If mlngState = 0 And mlngFun = 0 Then
        mblnOK = True
    Else
        mblnOK = False
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, str评估人 As String, strTmp As String
    Dim blnOver As Boolean, blnOK As Boolean, lngLen As Long
    Dim strSql As String, str审核人 As String, strVariation As String
    Dim rsTmp As ADODB.Recordset
    Dim lngMax As Long, lngMin As Long
    Dim str变异原因 As String
    Dim str跳转审核人 As String
    Dim blnTmp As Boolean
    
    '如果有数据，则必须选择一个变异原因，变异说明可以不输
    If optResult(0).Value = False And vsVariation.Rows > vsVariation.FixedRows Then
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, col变异选择) = 1 Then
                    strVariation = strVariation & "," & .RowData(i)
                    str变异原因 = str变异原因 & "," & Mid(.TextMatrix(i, col变异原因), InStr(.TextMatrix(i, col变异原因), "-") + 1)
                End If
            Next
            strVariation = Mid(strVariation, 2)
            If str变异原因 = "" And vsVariation.Enabled Then
                MsgBox "请选择一种变异原因。", vbInformation, gstrSysName
                If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
                Exit Sub
            End If
        End With
    End If
    
    '如果变异原因是其他则要求必须填写变异说明
    If InStr(str变异原因 & ",", ",其他,") > 0 Or InStr(str变异原因 & ",", ",其它,") > 0 Then
        If Trim(txtRemark.Text) = "" Then
            MsgBox "变异原因为其他的，必须填写评估备注。", vbInformation, gstrSysName
            If txtRemark.Enabled Then txtRemark.SetFocus
            Exit Sub
        End If
    End If
    
    If txtRemark.Text <> Trim(txtRemark.Text) Then txtRemark.Text = Trim(txtRemark.Text)
    If mlngFun = 0 Then
        lngLen = Sys.FieldsLength("病人临床路径", "导入说明")
    Else
        lngLen = Sys.FieldsLength("病人路径评估", "评估说明")
    End If
    If zlCommFun.ActualLen(txtRemark.Text) > lngLen Then
        Call MsgBox("备注信息不能超过最大长度" & lngLen, vbInformation, gstrSysName)
        txtRemark.SetFocus
        Exit Sub
    End If
    
    '评估指标
    If vsCriterion.Visible Then
        With vsCriterion
            For i = .FixedRows To .Rows - 1
                If InStr(.TextMatrix(i, mcol("评估指标")), "|") > 0 Then
                    MsgBox "第" & i & "行，评估指标中含有特殊字符:|，不能保存数据，请与系统管理员联系！", vbExclamation, gstrSysName
                    Exit Sub
                End If
                If .TextMatrix(i, mcol("结果")) = "" Then
                    MsgBox "第" & i & "行，评估指标未填写评估结果，请填写后再评估。", vbInformation, gstrSysName
                    .Select i, mcol("结果")
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If mlngFun = 1 Then
        With vsPersonnel
            For i = .FixedRows To .Rows - 1
                strTmp = Trim(.TextMatrix(i, 0))
                If strTmp <> "" Then
                    str评估人 = str评估人 & "," & strTmp
                End If
            Next
            str评估人 = Mid(str评估人, 2)
        End With
        If str评估人 = "" Then
            MsgBox "评估人未填写，请至少输入一名评估人。", vbInformation, gstrSysName
            Exit Sub
        ElseIf LenB(str评估人) > 50 Then
            MsgBox "评估人太多，超过最大长度50", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '当天是最后一个阶段的最后一天，则评估完成后自动结束路径
    If mlngFun = 1 Then
        If optResult(0).Value Or optResult(1).Value Then
            blnOver = IsLastDate(False, lngMin, lngMax, mPP.路径ID, mPP.版本号, mPP.当前阶段ID, mPP.当前阶段分支ID, mPP.当前天数)
            
            If InStr(GetInsidePrivs(p临床路径应用), ";结束路径;") = 0 Then
                blnOver = False
            End If
        End If
        
            
        '路径跳转时不结束
        If cmdTurn.Enabled And cmdTurn.Visible And cmdTurn.Tag <> "" Then
            blnOver = False
        End If
        
        If optResult(0).Value Or optResult(1).Value And optDate(0).Value Then
            '变异继续时，如果继续当前阶段或提前进入下一阶段，则不结束路径；变异结束不需再检查
            If blnOver Then
                '先判断该路径是否允许诊断不同完成路径，不允许则检查出院诊断是否和导入诊断相同
                If mPP.结束路径控制 = 0 Then
                    If Not CheckPathOutDiag(mPP.病人路径ID, mPati.病人ID, mPati.主页ID) Then
                        MsgBox "出院诊断不在适用病种范围内，不允许正常完成路径，只能变异退出路径。", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        MsgBox "注意：目前已达到或超过标准住院日，评估执行后将自动完成病人路径。", vbInformation, gstrSysName
                    End If
                Else
                    MsgBox "注意：目前已达到或超过标准住院日，评估执行后将自动完成病人路径。", vbInformation, gstrSysName
                End If
            End If
            
        ElseIf optResult(3).Value Then
            blnOver = True
            '先判断该路径是否允许诊断不同完成路径，不允许则检查出院诊断是否和导入诊断相同
            If mPP.结束路径控制 = 0 Then
                If Not CheckPathOutDiag(mPP.病人路径ID, mPati.病人ID, mPati.主页ID) Then
                    MsgBox "出院诊断不在适用病种范围内，不允许正常完成路径，只能变异退出路径。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Else
            blnOver = False
        End If
        
        '超过标准住院日变异继续,或变异退出，要审核
        If optResult(1).Value And mPP.当前天数 > lngMax Or optResult(2).Value Then
            If InStr(GetInsidePrivs(p临床路径应用), ";变异审核;") = 0 Then
                str审核人 = zlDatabase.UserIdentify(Me, "变异退出或超期继续需要审核。", glngSys, p临床路径应用, "变异审核")
                If str审核人 = "" Then Exit Sub
            Else
                str审核人 = UserInfo.姓名
            End If
        End If
    End If
    
    '检查合并路径
    If Not mrsMerge Is Nothing And mlngFun = 1 And mlngState = 1 Then
        If mrsMerge.RecordCount > 0 Then
            Set rsTmp = zlDatabase.CopyNewRec(mrsMerge)
            mrsMerge.MoveFirst
            Do While Not mrsMerge.EOF
                '检查不符合勾选要求的
                '未达到标准住院日的，但是又勾选了的
                If Val(mrsMerge!选择 & "") = 1 Then
                    If Not (Val(mrsMerge!状态 & "") = 1 Or Val(mrsMerge!状态 & "") = 2 Or optDate(1).Value Or optDate(1).Enabled = False) Then
                        MsgBox "您想完成的合并路径未到达标准住院日，如需提前完成，请选择下一阶段提前。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    '必须勾选但未勾选
                    If Val(mrsMerge!状态 & "") = 2 And Not optDate(2).Value Then
                        mrsMerge!选择 = 1
                        mrsMerge.Update
                        blnTmp = True
                    End If
                End If
                mrsMerge.MoveNext
            Loop
            mrsMerge.MoveFirst
            If blnTmp And lblMSG.Caption = "" Then
                '如果状态文字为Null则未选择过要停止的合并路径，但是又有必须要完成的合并路径（到了最后一天）,则提示
                If MsgBox("有合并路径达到了标准住院日，继续将自动完成合并路径，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                    Set mrsMerge = rsTmp
                    Exit Sub
                End If
            End If
        End If
    End If
        
    If blnOver Or optResult(2).Value Then
        '判断是否有未审核的阶段
        If InStr(GetInsidePrivs(p临床路径应用), ";跳转审核;") = 0 Then
            If CheckPathIsTurnAduit Then
                str跳转审核人 = zlDatabase.UserIdentify(Me, "前面阶段存在未审核的路径跳转，必须审核后才允许" & IIf(optResult(2).Value, "退出", "完成") & "。", glngSys, p临床路径应用, "跳转审核")
                If str跳转审核人 = "" Then Exit Sub
            End If
        Else
            str跳转审核人 = UserInfo.姓名
        End If
        If CheckPathOutLog Then
            blnOK = frmPathOutLog.ShowMe(Me, mPati.病人ID, mPati.主页ID, 0, mcolSQL, mPP.路径ID, mPP.病人路径ID)
            If blnOK = False Then
                i = Val(zlDatabase.GetPara("必须填写出径登记表", glngSys, p临床路径应用, "0"))
                If i = 1 Then Exit Sub
            End If
        End If
    End If
    
    
    '点了确定后将确定设置为不可用，等执行完了再启用，防止界面卡死用户多次点击。
    cmdOK.Enabled = False
    If mlngFun = 1 Then
    '阶段评估前调用接口
        blnTmp = True
        If CreatePlugInOK(p临床路径应用) Then
            On Error Resume Next
            blnTmp = gobjPlugIn.PathEvaluateBefore(glngSys, p临床路径应用, mPati.病人ID, mPati.主页ID, mPP.病人路径ID, mPP.当前阶段ID)
            '如果接口不存在，不影响原有逻辑
            If Not blnTmp And Err.Number <> 0 Then blnTmp = True
            Call zlPlugInErrH(Err, "PathEvaluateBefore")
            Err.Clear: On Error GoTo 0
        End If
        If Not blnTmp Then Unload Me: Exit Sub
    End If
    
    Call SaveData(blnOver, str审核人, str评估人, strVariation, str跳转审核人)
    
    If mlngFun = 1 Then
    '阶段评估后调用接口
        If CreatePlugInOK(p临床路径应用) Then
            On Error Resume Next
            Call gobjPlugIn.PathEvaluateAfter(glngSys, p临床路径应用, mPati.病人ID, mPati.主页ID, mPP.病人路径ID, mPP.当前阶段ID)
            Call zlPlugInErrH(Err, "PathEvaluateAfter")
            Err.Clear: On Error GoTo 0
        End If
    End If
    
    mblnOK = True
    cmdOK.Enabled = True
    Unload Me
End Sub

Private Sub SaveData(ByVal blnOver As Boolean, ByVal str审核人 As String, ByVal str评估人 As String, ByVal strVariation As String, ByVal str跳转审核人 As String)
'功能:保存数据
'参数:str审核人=变异退出或超期继续的审核人
'   blnOver=    最后一天评估时结束路径
'   strVariation=变异原因
    Dim strSql As String, str评估说明 As String, lng评估结果 As Long
    Dim strID As String, str符合导入 As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strTotal As String, strThis As String, dateInPath As Date
    Dim str时间进度 As String, str路径跳转 As String
    Dim rsTmp As Recordset, dateCur As Date
    Dim AddDate As Date
    Dim str合并路径记录IDs As String   '要结束的合并路径记录ID

    If mlngFun = 0 Then
        str符合导入 = IIf(optResult(0).Value = True, "1", "0")
        '如果是合并路径不符合，直接退出。
        If str符合导入 = "0" And mlngType = 1 Then Exit Sub
        str评估说明 = Trim(txtRemark.Text)
        dateCur = zlDatabase.Currentdate
        If mlngType = 0 Then
            strID = zlDatabase.GetNextId("病人临床路径")
        Else
            strID = zlDatabase.GetNextId("病人合并路径")
        End If
        If optStart(0).Value Then
            If CheckPathSend(mPati.病人ID, mPati.主页ID) Then
                dateInPath = dateCur
            Else
                dateInPath = GetPatiInDate(mPati)
            End If
            AddDate = dateCur
        Else
            dateInPath = dateCur
        End If
        
        strSql = "Zl_病人路径导入_Insert(" & mPati.病人ID & "," & mPati.主页ID & "," & mPati.科室ID & "," & _
                mPP.路径ID & "," & mPP.版本号 & "," & strID & ",'" & UserInfo.姓名 & "','" & str评估说明 & "'," & _
                str符合导入 & ",To_Date('" & Format(dateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                strVariation & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng疾病ID = 0, "NULL", mlng疾病ID) & "," & IIf(mlng诊断ID = 0, "NULL", mlng诊断ID)
                
        If vsCriterion.Visible = False Then
            colSQL.Add strSql & ",Null," & colSQL.count + 1 & "," & mlngType & "," & mlng首要路径记录ID & ")", "C" & colSQL.count + 1
        Else
            With vsCriterion
                For i = .FixedRows To .Rows - 1
                    strThis = .TextMatrix(i, mcol("评估指标")) & "|" & .TextMatrix(i, mcol("结果")) & "|" & .TextMatrix(i, mcol("指标类型")) & "||"
                    If LenB(strTotal & strThis) > 4000 Then
                        colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & "," & mlngType & "," & mlng首要路径记录ID & ")", "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                Next
                If strTotal <> "" Then
                    colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & "," & mlngType & "," & mlng首要路径记录ID & ")", "C" & colSQL.count + 1
                Else
                    colSQL.Add strSql & ",Null," & colSQL.count + 1 & "," & mlngType & "," & mlng首要路径记录ID & ")", "C" & colSQL.count + 1
                End If
            End With
        End If
        If mlngType = 0 And mblnPathSend = False Then
            '如果开始时间不是当前时间，则自动匹配并补齐之前的路径阶段和项目。
            If optStart(0).Value And optResult(0).Value Then
                '匹配路径项目
                Call CreatePathItem(dateCur, dateInPath, mPati, mPP, CLng(strID), colSQL)
            End If
        End If
    Else
        str评估说明 = Trim(txtRemark.Text)
        
        If cmdTurn.Enabled And cmdTurn.Visible Then str路径跳转 = cmdTurn.Tag '路径ID,版本号
        If str路径跳转 = "" Then str路径跳转 = "Null,Null,,Null"  '路径ID,版本号,跳转审核人,审核历史跳转
        
        If optDate(0).Value Then
            str时间进度 = "0"
        ElseIf optDate(1).Value Then
            str时间进度 = "1"      '下一阶段提前至今天
        ElseIf optDate(3).Value Then
            str时间进度 = "2"     '下一阶段提前至明天
        ElseIf optDate(2).Value Then
            str时间进度 = "-1"    '延后
        End If
        
        lng评估结果 = 0
        If optResult(1).Value Then
            lng评估结果 = 1
        ElseIf optResult(2).Value Then
            lng评估结果 = 2
        ElseIf optResult(3).Value Then
            lng评估结果 = 3
        End If
        If Not mrsMerge Is Nothing And mlngFun = 1 And mlngState = 1 Then
            If mrsMerge.RecordCount > 0 Then
                mrsMerge.MoveFirst
                Do While Not mrsMerge.EOF
                    '检查不符合勾选要求的
                    '未达到标准住院日的，但是又勾选了的
                    If Val(mrsMerge!选择 & "") = 1 Then
                        str合并路径记录IDs = str合并路径记录IDs & "," & mrsMerge!ID
                    End If
                    mrsMerge.MoveNext
                Loop
                str合并路径记录IDs = Mid(str合并路径记录IDs, 2)
            End If
        End If
        
        strSql = "Zl_病人路径评估_Insert(" & mlngState & "," & mPP.病人路径ID & "," & mPP.当前阶段ID & _
            ",To_Date('" & mPP.当前日期 & "','YYYY-MM-DD')," & mPP.当前天数 & ",'" & _
            str评估人 & "'," & lng评估结果 & ",'" & str评估说明 & "','" & UserInfo.姓名 & "','" & str审核人 & "','" & strVariation & "'," & str时间进度 & "," & Split(str路径跳转, ",")(0) & "," & Split(str路径跳转, ",")(1)
            
        With vsCriterion
            If .Visible Then    '可以不设置指标
                For i = .FixedRows To .Rows - 1
                    strThis = .TextMatrix(i, mcol("评估指标")) & "|" & .TextMatrix(i, mcol("结果")) & "|" & .TextMatrix(i, mcol("指标类型")) & "||"
                    If LenB(strTotal & strThis) > 4000 Then
                        colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & ",'" & Split(str路径跳转, ",")(2) & "'," & Split(str路径跳转, ",")(3) & ",'" & str合并路径记录IDs & "')", "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                Next
                If strTotal <> "" Then
                    colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & ",'" & Split(str路径跳转, ",")(2) & "'," & Split(str路径跳转, ",")(3) & ",'" & str合并路径记录IDs & "')", "C" & colSQL.count + 1
                Else
                    colSQL.Add strSql & ",Null," & colSQL.count + 1 & ",'" & Split(str路径跳转, ",")(2) & "'," & Split(str路径跳转, ",")(3) & ",'" & str合并路径记录IDs & "')", "C" & colSQL.count + 1
                End If
            Else
                colSQL.Add strSql & ",Null," & colSQL.count + 1 & ",'" & Split(str路径跳转, ",")(2) & "'," & Split(str路径跳转, ",")(3) & ",'" & str合并路径记录IDs & "')", "C" & colSQL.count + 1
            End If
        End With
        If blnOver Then
            strSql = "Zl_病人路径结束_Update(" & mPP.病人路径ID & ",'" & str跳转审核人 & "')"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        '执行出径登记表的SQL
        For i = 1 To mcolSQL.count
            Call zlDatabase.ExecuteProcedure(mcolSQL("C" & i), "出径登记表")
        Next
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "路径评估")
        Next
    gcnOracle.CommitTrans: blnTrans = False
    '消息发送
    strSql = ""
    For i = 1 To mcolSQL.count
        If InStr(UCase(mcolSQL("C" & i)), "Zl_病人路径生成_INSERT") > 0 Then
            strSql = "do"
            Exit For
        End If
    Next
    
    If strSql <> "" Then
        For i = 1 To colSQL.count
            If InStr(UCase(colSQL("C" & i)), "Zl_病人路径生成_INSERT") > 0 Then
                strSql = "do"
                Exit For
            End If
        Next
    End If
    
    If strSql <> "" Then
        Call ZLHIS_CIS_001(Nothing, mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID)
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsVariation_Click()
    Dim i As Long
    
    With vsVariation
        If .Row >= .FixedRows Then
            .Redraw = flexRDNone
            If mlngFun = 1 Then  '阶段评估
                If .Cell(flexcpData, .Row, col变异选择) = 0 Then
                    Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("Check").Picture
                    .Cell(flexcpData, .Row, col变异选择) = 1
                Else
                    Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("UnCheck").Picture
                    .Cell(flexcpData, .Row, col变异选择) = 0
                End If
            ElseIf mlngFun = 0 Then '导入评估
                If .Cell(flexcpData, .Row, col变异选择) = 0 Then
                    Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("Selected").Picture
                    .Cell(flexcpData, .Row, col变异选择) = 1
                    For i = .FixedRows To .Rows - 1
                        If i <> .Row Then
                            If .Cell(flexcpData, i, col变异选择) = 1 Then
                                Set .Cell(flexcpPicture, i, col变异选择) = imgNature.ListImages("UnSelected").Picture
                                .Cell(flexcpData, i, col变异选择) = 0
                            End If
                        End If
                    Next
                Else
                    Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("UnSelected").Picture
                    .Cell(flexcpData, .Row, col变异选择) = 0
                End If
            End If
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub vsVariation_GotFocus()
    If vsVariation.Row < vsVariation.FixedRows And vsVariation.Rows > vsVariation.FixedRows Then vsVariation.Row = vsVariation.FixedRows
End Sub

Private Sub vsVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call vsVariation_Click
    End If
End Sub

Private Sub txtVariation_GotFocus()
    Call zlControl.TxtSelAll(txtVariation)
End Sub

Private Sub txtVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim i As Long, strtxt As String
        strtxt = "*" & UCase(Trim(txtVariation.Text)) & "*"
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .RowData(i) Like strtxt Or .TextMatrix(i, col变异原因) Like strtxt Or .Cell(flexcpData, i, col变异原因) Like strtxt Then
                    .SetFocus
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
End Sub

Private Function IsLastDate(Optional ByVal blnEnd As Boolean, Optional ByRef lngMin As Long, Optional ByRef lngMax As Long, Optional ByVal lng路径ID As Long _
                                                                                                                            , Optional ByVal lng版本号 As Long, Optional ByVal lng当前阶段ID As Long, Optional ByVal lng当前阶段分支ID As Long, Optional ByVal lng当前天数 As Long, _
                            Optional ByVal blnBoth As Boolean, Optional ByRef lngState As Long, Optional ByVal lng合并路径记录ID As Long) As Boolean
'功能：判断是否退出路径
'      blnEnd=false:判断当前天数是否是路径最后阶段的最后一天，且没有后续阶段
'      blnEnd= true:是否允许变异退出（在标准住院日范围内都可退出）
'参数：blnBoth=合并路径检查时，同时检查最后一天或是达到标准住院日范围
'     lng首要路径阶段ID:blnBoth=true时，合并路径的起点阶段
'返回：lngMin，lngMax标准住院日
'      lngState :当blnBoth=true  返回0=未达到标准住院日，1=达到标准住院日，但为达到最后一天，2=标准住院日最后一天
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim arrtmp As Variant, lng实际天数 As Long, lng理论天数 As Long
    Dim blnIsLastDate As Boolean
    
    lngState = 0 'lngState为引用传值，初始为0。
    
    If lng当前阶段分支ID = 0 Then
        strSql = "Select 标准住院日 From 临床路径版本 Where 路径id = [1] And 版本号 = [2]"
    Else
        strSql = "Select 标准住院日 From 临床路径分支 Where 路径id = [1] And 版本号 = [2] And ID=[3]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, lng版本号, lng当前阶段分支ID)
    If Not IsNull(rsTmp!标准住院日) Then
        arrtmp = Split(rsTmp!标准住院日, "-")
        If UBound(arrtmp) > 0 Then
            lngMin = arrtmp(0)
            lngMax = arrtmp(1)
        Else
            lngMin = 1  '小于等于n天
            lngMax = arrtmp(0)
        End If

        If blnEnd Or blnBoth Then
            lng理论天数 = GetMustDay(mPP.病人路径ID, lng当前天数, , lng合并路径记录ID)
            If lng理论天数 > lngMax Then
                blnIsLastDate = True

            Else
                blnIsLastDate = Between(lng理论天数, lngMin, lngMax)
            End If
            If blnIsLastDate And blnBoth Then
                lngState = 1
            End If
        End If
        If blnIsLastDate Then IsLastDate = blnIsLastDate
        If Not blnEnd Or blnBoth Then
            lng实际天数 = GetMustDay(mPP.病人路径ID, lng当前天数, True, lng合并路径记录ID)
            If lng实际天数 >= lngMax Then
                blnIsLastDate = GetNextPhase(lng当前阶段ID, lng当前阶段分支ID) = 0
                If blnIsLastDate And blnBoth Then
                    lngState = 2
                End If
            End If
        End If
    End If
    If blnIsLastDate Then IsLastDate = blnIsLastDate
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
