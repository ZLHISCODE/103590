VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdviceEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmAdviceEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraMethod 
      BackColor       =   &H8000000E&
      Height          =   2175
      Left            =   1680
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdMethodOK 
         Caption         =   "确定"
         Height          =   300
         Left            =   1065
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMethod 
         Height          =   1815
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
         _cx             =   1993543209
         _cy             =   1993542785
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAdviceEditEx.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   Begin VB.PictureBox picSentence 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2880
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
      Begin VB.TextBox txtSentence 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   15
         TabIndex        =   2
         Top             =   30
         Width           =   930
      End
      Begin VB.Image imgSentence 
         Height          =   210
         Left            =   960
         Picture         =   "frmAdviceEditEx.frx":0048
         ToolTipText     =   "请按 * 号键选择"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   4
      Left            =   1470
      MousePointer    =   7  'Size N S
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   1155
      MousePointer    =   9  'Size W E
      TabIndex        =   17
      Top             =   2265
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   16
      Top             =   2250
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   495
      TabIndex        =   15
      Top             =   2535
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   495
      MousePointer    =   7  'Size N S
      TabIndex        =   14
      Top             =   2265
      Width           =   615
   End
   Begin VB.OptionButton optMode 
      Caption         =   "术中"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   3090
      TabIndex        =   11
      Top             =   2745
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.OptionButton optMode 
      Caption         =   "床旁"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   2415
      TabIndex        =   10
      Top             =   2745
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.OptionButton optMode 
      Caption         =   "常规"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   1740
      TabIndex        =   9
      Top             =   2745
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "…"
      Height          =   240
      Left            =   2225
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   3555
      Picture         =   "frmAdviceEditEx.frx":0572
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "取消(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4140
      _cx             =   1993546886
      _cy             =   1993542838
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceEditEx.frx":0AFC
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Begin MSComctlLib.ImageList img16 
         Left            =   1650
         Top             =   975
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
               Picture         =   "frmAdviceEditEx.frx":0BF7
               Key             =   "c0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":1191
               Key             =   "c1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":172B
               Key             =   "o0"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":1CC5
               Key             =   "o1"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo标本 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmAdviceEditEx.frx":225F
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "确认(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   870
      Left            =   135
      TabIndex        =   4
      Top             =   3015
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1535
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdviceEditEx.frx":27E9
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
   Begin VB.Line lin 
      Index           =   7
      X1              =   2475
      X2              =   3150
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   2475
      X2              =   3150
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   2475
      X2              =   3150
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   2475
      X2              =   3150
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   2475
      X2              =   3150
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   2475
      X2              =   3150
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   2475
      X2              =   3150
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   2475
      X2              =   3150
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblAppend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单据附项：(按~键可输入词句示范)"
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   2745
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "麻醉"
      Height          =   180
      Left            =   105
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmAdviceEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Const EM_POSFROMCHAR = &HD6
Private Const EM_EXGETSEL = (&H400 + 52)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
'=============================================================================================================
'入口参数：
Private mlngHwnd As Long '用于定位的控件句柄
Private mint期效 As Integer
Private mstr性别 As String
Private mint调用类型 As Integer  '1-门诊,2-住院
Private mint服务对象 As Integer '1-门诊,2-住院
Private mbytUseType As Byte      '0=医嘱下达,1-路径项目的医嘱生成,2-添加路径外项目

'0-检查组合,1-手术输入,4-检验组合，5-输血或治疗类或其他类需要填写申请附项的
Private mintType As Integer

'入:主诊疗项目ID
Private mlng项目ID As Long


'入/出:附加定义数据,新增时一般为空
'      检查="部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'      手术="手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
'      检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"
Private mstrExtData As String

'入/出:申请附项内容,新增时为空
'     格式="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
Private mstrAppend As String

'入：尚未保存的医嘱中已录入的附项内容,以最新的为准,用于新增时提取使用
'     格式=项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>...
Private mstrAdvItem As String

'入：尚未保存的医嘱已对应录入的诊断(不仅仅是当前医嘱所对应的)
Private mstrDiagnosis As String


'入:部份情况下需要,如检查申请取附项内容
Private mlng病人ID As Long
Private mvar就诊ID As Variant '主页ID或挂号单号
Private mint婴儿 As Integer
Private mlng病人科室id As Long

Private mint场合 As Integer  '0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Private mblnNew As Boolean  '判断是否是新开输入项目时进入，否则为点下箭头进入

'出：手术医嘱返回手术部位要素对应的值，以便拼到医嘱内容中
Private mstr手术部位 As String


'入：判断检验组合是否使用新版LIS的检验组合模式
Private mblnNewLIS As Boolean

'出口参数：
Private mblnOK As Boolean '出

'程序变量
Private mstrMatchMode As Boolean
Private mint简码 As Integer
Private mstrLike As String
Private mbln附项 As Boolean
Private mblnFirst As Boolean
Private mblnReturn As Boolean '是否了回车确认
Private mblnNotAddNew As Boolean '是否不允许增加
Private mbytSize As Byte '字体大小 0-小字体（9号），1-大字体（12号）
Private mstr手术等级 As String   '主刀医生的手术等级
Private mbln手术分级管理 As Boolean   '是否启用手术分级管理
Private mbln手术授权管理 As Boolean
Private mbln手术等级管理 As Boolean  '是否启用参数：主刀医师达到手术等级无需审核
Private mrsAppend As ADODB.Recordset
Private mbln检查部位 As Boolean    '该检查项目是否需要设置部位
Private mstr单选项目 As String
Private mbln技师站 As Boolean '是否是技师站调用
'模块号定义
Public Enum Enum_Program_Modual
    pm门诊医嘱下达 = 1252
    pm住院医嘱下达 = 1253
    pm门诊医生站 = 1260
    pm住院医生站 = 1261
    pm住院护士站 = 1262
    pm医技工作站 = 1263
End Enum

Private mblnChangeSel As Boolean
Private mstrPrivs As String             '权限
Private mfrmParent As Object
Private mobjEmrInterface As Object           '新版病历申请附项读取部件

Public Function ShowMe(ByVal frmParent As Object, ByVal lngHwnd As Long, ByRef t_Pati As TYPE_PatiInfoEx, ByVal int场合 As Integer, _
            ByVal intType As Integer, ByVal bytUseType As Byte, ByVal int期效 As Integer, ByVal int服务对象 As Integer, Optional ByVal int调用类型 As Integer, _
            Optional ByVal blnNewLIS As Boolean, Optional ByVal blnNew As Boolean, Optional ByVal lng项目id As Long, Optional ByRef strExtData As String, _
            Optional ByRef strAppend As String, Optional ByVal strAdvItem As String, Optional ByVal strDiagnosis As String, Optional ByRef str手术部位 As String, Optional ByVal bln技师站 As Boolean) As Boolean
'参数:
'     frmParent         父窗体
'     lngHwnd           用于定位的控件句柄,即调用该窗体的控件
'     t_Pati            病人信息
'     int场合           0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'     intType           0-检查组合,1-手术输入,4-检验组合，5-输血或治疗类需要填写申请附项的
'     bytUseType        0=医嘱下达,1-路径项目的医嘱生成,2-添加路径外项目
'     int期效           将要输入的医嘱期效 0-长嘱，1-临嘱
'     int服务对象       该医嘱要服务的病人性质 1-门诊（包括门诊病人，体检病人，外来病人等) 2-住院（只有住院病人）
'     int调用类型       调用该窗体的工作站类型 1-门诊医生工作站 2-住院医护工作站
'     blnNewLIS         判断检验组合是否使用新版LIS的检验组合模式
'     blnNew            判断是否是新开输入项目时进入，否则为点下箭头进入。 true-新开输入项目时进入， false-点下箭头进入（现在只针对检验，只在新版LIS中使用（blnNewLIS=true)）
'     lng项目id         主诊疗项目ID
'     strAdvItem        尚未保存的医嘱中已录入的附项内容,以最新的为准,用于新增时提取使用
'                       格式=项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>...
'    strDiagnosis       尚未保存的医嘱已对应录入的诊断(不仅仅是当前医嘱所对应的)
'返回：
'     strExtData        附加定义数据 , 新增时一般为空
'                       检查 = "部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'                       手术="手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
'                       检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"
'     strAppend         新增时为空，格式="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
'     str手术部位       手术医嘱，返回“手术部位”要素对应的值
'     bln技师站         当前调用方是技师工作站
    Set mfrmParent = frmParent
    mlngHwnd = lngHwnd
    With t_Pati
        mint婴儿 = .int婴儿
        mlng病人ID = .lng病人ID
        mlng病人科室id = .lng病人科室ID
        mvar就诊ID = IIF(.str挂号单 = "", .lng主页ID, .str挂号单)
        mstr性别 = .str性别
    End With
    mint场合 = int场合
    mintType = intType
    mbytUseType = bytUseType
    mint期效 = int期效
    mint服务对象 = int服务对象
    mint调用类型 = int调用类型
    mblnNewLIS = blnNewLIS
    mblnNew = blnNew
    mlng项目ID = lng项目id
    mstrExtData = strExtData
    mstrAppend = strAppend
    mstrAdvItem = strAdvItem
    mstrDiagnosis = strDiagnosis
    mbln技师站 = bln技师站
    mblnOK = False
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    
    strExtData = mstrExtData
    strAppend = mstrAppend
    str手术部位 = mstr手术部位
    
    
    ShowMe = mblnOK
End Function

Private Sub cbo标本_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo标本.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cbo标本.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo标本.ListCount > 0 Then lngIdx = 0
        cbo标本.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdMethodOK_Click()
    Call vsMethod_KeyPress(vbKeyReturn)
End Sub

Private Sub rtfAppend_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtSentence.Tag = "输入医生" Then
        If picSentence.Visible = False And KeyCode > 127 Then KeyCode = 0: Call rtfAppend_SelChange
        If txtSentence.Tag = "输入医生" And KeyCode = vbKeyBack Then KeyCode = 0: Call rtfAppend_SelChange
    End If
End Sub

Private Sub cmd_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim int性别 As Integer, strSQLItem As String, i As Long
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSamples As String, strPrivs As String
    
    If mstr性别 Like "*男*" Then
        int性别 = 1
    ElseIf mstr性别 Like "*女*" Then
        int性别 = 2
    End If
    
    On Error GoTo errH
    
    If mintType = 1 Then
        '输入附加手术:这里不是单独应用,因此不限制
        '"-1*主手术ID"是不排开主手术ID，以作为附加手术加收费用
        strSQLItem = _
            " From 诊疗项目目录 A Where A.类别='F' And A.ID<>-1*" & mlng项目ID & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " And A.服务对象 IN([1],3) And Nvl(A.执行频率,0) IN(0,[2]) And Nvl(A.适用性别,0) IN(0,[3])"
        
        strSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位,NULL as 规模" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
            " Group by ID,上级ID,编码,名称"
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
            strSQLItem & " Order By 末级,级ID Desc,编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "手术", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint服务对象, IIF(mint期效 = 0, 2, 1), int性别, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的手术项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set手术输入(vsExt.Row, rsTmp)
    ElseIf mintType = 4 Then
        '检验项目
        With Me.cbo标本
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        
        strSQL = "Select 0 as 末级,名称 as ID,-Null as 上级ID,编码,名称,' ' as 检验类型,' ' As 标本部位,NULL as 试管编码 From 诊疗检验类型" & _
            " Union ALL" & _
            " Select Distinct 1 as 末级,''||A.ID as ID,A.操作类型 as 上级ID,A.编码,A.名称,A.操作类型 as 检验类型,A.标本部位,A.试管编码 " & _
            " From 诊疗项目目录 A,检验项目参考 C,检验报告项目 D " & _
            " Where A.ID=D.诊疗项目id(+) And D.报告项目ID=C.项目id(+)" & _
            " And A.类别='C' " & _
            IIF(mint场合 = 2, "", " And Nvl(A.单独应用,0)=1 ") & _
            " And Nvl(A.适用性别,0) In (0,[2])" & _
            " And A.服务对象 IN([1],3" & IIF(mint场合 = 2, ",4", "") & ") " & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[3])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " And (C.标本类型 In (" & strSamples & ") Or C.标本类型 Is Null)" & _
            " Order By 末级,编码 "
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "检验项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint服务对象, int性别, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的检验项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
'        If rsTmp!检验类型 = "微生物" And vsExt.Rows > 2 Then
'            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
'                MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
'                Exit Sub
'            End If
'        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查检验类型，试管编码是否相同
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = NVL(rsTmp!检验类型) _
                    Or vsExt.TextMatrix(i, 1) = "" Or NVL(rsTmp!检验类型) = "") Then
                    MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Not (vsExt.Cell(flexcpData, i, 1) = CStr(NVL(rsTmp!试管编码)) _
                    Or vsExt.Cell(flexcpData, i, 1) = "" Or NVL(rsTmp!试管编码) = "") Then
                    MsgBox "请输入相同试管编码的项目，已输入项目的管编码为""" & vsExt.Cell(flexcpData, i, 1) & """。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '重新初始标本
        If Not InitCombox(rsTmp!ID, NVL(rsTmp!标本部位)) Then Exit Sub
        
        Call Set检验项目(vsExt.Row, rsTmp)
        If rsTmp("检验类型") = "微生物" Then
            mblnNotAddNew = False
'            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str性别 As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    Else
        str性别 = "0"
    End If
    
    If mintType = 1 Then
        '输入麻醉项目:这里不是单独应用,因此不限制
        strSQLItem = " From 诊疗项目目录 A Where A.类别='G'" & _
                " And A.服务对象 IN([2],3) And A.ID<>[1]" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[3])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"

        strSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位,NULL as 麻醉类型" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
            " Group by ID,上级ID,编码,名称"
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
            strSQLItem & " Order By 末级,级ID Desc,编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "麻醉项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng项目ID, mint服务对象, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '输入标本
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Get检查部位方法(str检查部位 As String, str检查方法 As String)
'功能：收集检查对应的部位及方法,用","号间隔
    Dim i As Long
    
    str检查部位 = "": str检查方法 = ""
    
    With vsExt
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 1) = 1 Then
                str检查部位 = str检查部位 & "," & .TextMatrix(i, 1)
                If .TextMatrix(i, 2) <> "" Then
                    str检查方法 = str检查方法 & "," & .TextMatrix(i, 2)
                End If
            End If
        Next
        str检查部位 = Mid(str检查部位, 2)
        str检查方法 = Mid(str检查方法, 2)
    End With
End Sub

Private Function GetMax手术等级(ByVal str手术项目 As String) As String
'功能：取得当前医嘱最大手术类型
'参数：str手术项目：手术项目ID用，分隔，lng手术等级返回最高的手术等级
    Dim strSQL As String, rsTmp As Recordset
    Dim str手术等级 As String
    
    On Error GoTo errH
    strSQL = "Select a.手术类型 From 疾病编码目录 A,疾病诊断对照 B Where a.ID=b.疾病ID And a.类别='S' And instr([1], b.手术id)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, str手术项目)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If decode(rsTmp!手术类型 & "", "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) > _
                decode(str手术等级, "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) Then
                str手术等级 = rsTmp!手术类型 & ""
            End If
            rsTmp.MoveNext
        Loop
    End If
    GetMax手术等级 = str手术等级
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim str手术类型 As String
    Dim str人员等级 As String
    Dim blnMsg As Boolean
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strAppend As String, strData As String
    
    If mintType = 0 Then '检查部位组合
        '收集部位及方法的情况
        With vsExt
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, 1) = 1 Then
                    If .TextMatrix(i, 2) = "" Then
                        .Row = i: .ShowCell .Row, .Col
                        MsgBox "没有为检查部位""" & .TextMatrix(i, 1) & """确定检查方法。", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    
                    strTmp = strTmp & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
                End If
            Next
            If strTmp = "" And vsExt.Editable <> flexEDNone Then
                MsgBox "请至少选择一个检查部位。", vbInformation, gstrSysName
                vsExt.SetFocus: Exit Sub
            End If
            strTmp = Mid(strTmp, 2) & vbTab & IIF(optMode(0).Value, 0, IIF(optMode(1).Value, 1, 2))
        End With
    ElseIf mintType = 1 Or mintType = 4 Then '附加手术及麻醉项目；检验项目及标本
        '确认所输入的项目
        If mintType = 1 Or mintType = 4 And mblnNewLIS = False Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 Then
                    If vsExt.RowData(i) = mlng项目ID And mintType = 1 Then
                        MsgBox "在所填附加手术中出现与主要手术相同的手术。", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            Next
        ElseIf mintType = 4 And mblnNewLIS Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And (Val(vsExt.Cell(flexcpChecked, i, 0)) = 1 Or Val(vsExt.TextMatrix(i, 3)) = 0) Then
                    strTmp = strTmp & IIF(Val(vsExt.TextMatrix(i, 3)) = 1, "|", ",") & vsExt.RowData(i)
                End If
            Next
        End If
        strTmp = Mid(strTmp, 2)
        If mintType = 1 And mbln手术分级管理 Then
            '检查手术等级
            str人员等级 = IIF(mstr手术等级 <> "", mstr手术等级, UserInfo.手术等级)
            str手术类型 = GetMax手术等级(mlng项目ID & "," & strTmp)
            If decode(str手术类型, "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) > _
                decode(str人员等级, "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) Then
                blnMsg = True
            End If
        End If
        If strTmp = "" And mintType = 4 Then
            MsgBox "至少要选择一个检验项目。", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo标本.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    End If
    
    '检查并收集附项输入情况，有的地方 rtfAppend.Find方法可能不支持，所以要用 Instr 再判断下
    If rtfAppend.Visible Then
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            strData = "": lngBegin = -1: lngEnd = -1
            lngBegin = rtfAppend.Find(mrsAppend!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngBegin = -1 Then
                lngBegin = InStr(rtfAppend.Text, mrsAppend!项目 & "：")
                lngBegin = lngBegin - 1
            End If
            If lngBegin <> -1 Then
                lngBegin = lngBegin + Len(mrsAppend!项目 & "：")
                If i = mrsAppend.RecordCount Then
                    lngEnd = Len(rtfAppend.Text)
                Else
                    mrsAppend.MoveNext
                    lngEnd = rtfAppend.Find(vbCrLf & mrsAppend!项目 & "：", lngBegin, , rtfNoHighlight Or rtfMatchCase)
                    If lngEnd = -1 Then
                        lngEnd = InStr(rtfAppend.Text, vbCrLf & mrsAppend!项目 & "：")
                        lngEnd = lngEnd - 1
                    End If
                    If lngEnd = -1 Then
                        lngEnd = InStr(rtfAppend.Text, mrsAppend!项目 & "：")
                        lngEnd = lngEnd - 1
                    End If
                    mrsAppend.MovePrevious
                End If
            End If
            If lngBegin <> -1 And lngEnd <> -1 Then
                'MID函数是以1为基础，rtf是以0为基础
                lngBegin = lngBegin + 1
                lngEnd = lngEnd + 1
                strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                '去掉为解决保护文本后第一个位置不能直接录入汉字所添加的空格
                If Left(strData, 1) = " " Then strData = Mid(strData, 2)
                If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)
                
                If Trim(strData) = "" And NVL(mrsAppend!必填, 0) = 1 Then
                    MsgBox "单据附项""" & mrsAppend!项目 & """的内容没有填写。", vbInformation, gstrSysName
                    If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
                        rtfAppend.SelStart = lngBegin
                    Else
                        rtfAppend.SelStart = lngBegin - 1
                    End If
                    rtfAppend.SetFocus: Exit Sub
                ElseIf zlCommFun.ActualLen(strData) > 4000 Then
                    MsgBox "单据附项""" & mrsAppend!项目 & """的内容过长，最多允许2000个汉字或4000个字符。", vbInformation, gstrSysName
                    If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
                        rtfAppend.SelStart = lngBegin
                    Else
                        rtfAppend.SelStart = lngBegin - 1
                    End If
                    If rtfAppend.SelText = " " Then rtfAppend.SelStart = lngBegin
                    rtfAppend.SetFocus: Exit Sub
                End If
            End If
            
            '没有输入内容的附项也进行了保存
            strAppend = strAppend & "<Split1>" & mrsAppend!项目 & "<Split2>" & NVL(mrsAppend!必填, 0) & "<Split2>" & NVL(mrsAppend!要素ID) & "<Split2>" & strData
            If mintType = 1 And mrsAppend!中文名 & "" = "手术部位" Then mstr手术部位 = strData
            mrsAppend.MoveNext
        Next
        strAppend = Mid(strAppend, Len("<Split1>") + 1)
    End If
    
    
    If blnMsg Then
         MsgBox "当前手术等级为" & str手术类型 & "，主刀医师" & IIF(mstr手术等级 = "", "不能开展手术。", "只能开展" & mstr手术等级 & "级手术。"), vbInformation, gstrSysName
    End If
    
    '如果启用了手术授权管理，则检查主刀医师执行权
    If mintType = 1 And mbln手术授权管理 And mint调用类型 = 2 Then
        If CheckDocEmpowerEx(mlng项目ID, strAppend) = False Then
            If Not mbln手术等级管理 Then
                MsgBox "主刀医生不具备此手术的执行权，不允许下达。", vbInformation, "手术授权管理"
                Exit Sub
            Else
                MsgBox "主刀医生不具备此手术的执行权。", vbInformation, "手术授权管理"
            End If
        End If
    End If
    
    
    mstrExtData = strTmp
    mstrAppend = strAppend
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst And vsExt.TabStop And vsExt.Enabled And vsExt.Visible And Not Me.ActiveControl Is vsExt Then
        mblnFirst = False: vsExt.SetFocus '？不清楚为什么自动定位到rtfAppend上面去了。
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        If fraMethod.Visible Then
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf picSentence.Visible Then
            Call HideWordInput(True) '隐藏词句输入
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub Form_Resize()
    Dim lngAppend As Long
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    Dim lngHeight As Long, lngTotalHeight As Long
    Call HideWordInput(True) '隐藏词句输入
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Top + fraBorder(0).Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2
    
    If mbln附项 Then
        lngTotalHeight = Me.ScaleHeight - fraBorder(4).Height * 3 - lblAppend.Height * 2 - (cbo标本.Height + 200) - IIF(vsExt.Visible, vsExt.Top, 0)
        vsExt.Height = lngTotalHeight * 0.618

        fraBorder(4).Left = fraBorder(3).Width
        fraBorder(4).Top = IIF(vsExt.Visible, vsExt.Top, 0) + IIF(vsExt.Visible, vsExt.Height, 0)
        fraBorder(4).Width = Me.ScaleWidth - fraBorder(3).Width * 2
        
        lblAppend.Left = fraBorder(3).Width * 2
        lblAppend.Top = fraBorder(4).Top + fraBorder(4).Height * 2
        
        rtfAppend.Left = fraBorder(3).Width
        rtfAppend.Top = lblAppend.Top + lblAppend.Height + fraBorder(4).Height
        rtfAppend.Width = Me.ScaleWidth - fraBorder(3).Width * 2
        rtfAppend.Height = Me.ScaleHeight - rtfAppend.Top - fraBorder(2).Height - (cbo标本.Height + 200)
        
        lngAppend = rtfAppend.Top + rtfAppend.Height - fraBorder(4).Top
    Else
        vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cbo标本.Height + 200)
    End If
    
    cbo标本.Top = Me.ScaleHeight - fraBorder(2).Height - ((Me.ScaleHeight - fraBorder(0).Height * 2 - IIF(vsExt.Visible, vsExt.Height, 0) - lngAppend) - cbo标本.Height) / 2 - cbo标本.Height
    txtData.Top = cbo标本.Top
    lblData.Top = cbo标本.Top + (cbo标本.Height - lblData.Height) / 2
    cmdOK.Top = cbo标本.Top + (cbo标本.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
        
    optMode(0).Top = cbo标本.Top + (cbo标本.Height - optMode(0).Height) / 2
    optMode(1).Top = optMode(0).Top: optMode(2).Top = optMode(0).Top
    optMode(0).Left = 500
    optMode(1).Left = optMode(0).Left + optMode(0).Width + 100
    optMode(2).Left = optMode(1).Left + optMode(1).Width + 100

    lblData.Left = 200
    cbo标本.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    txtData.Left = cbo标本.Left
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
        
    cbo标本.Width = cmdOK.Left - cbo标本.Left - 200

    txtData.Width = cbo标本.Width
    cmdData.Top = txtData.Top + 30
    cmdData.Left = txtData.Left + txtData.Width - cmdData.Width - 45

    Me.Refresh
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str方法 As String, i As Long, lngBaseHeight As Long
    
    Me.Height = 2325
    
    '边框设置
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
    '输入匹配
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    If mint服务对象 = 0 Then mint服务对象 = 2 '缺省为住院
    mblnOK = False
    mblnNotAddNew = False
    mblnFirst = True
    mstr手术部位 = ""
    mbln手术分级管理 = False
    If mintType = 1 Then
        '是否启用手术分级管理
        mbln手术分级管理 = Val(zlDatabase.GetPara(209, glngSys)) <> 0
        '是否启用手术按医师授权管理
        mbln手术授权管理 = Val(zlDatabase.GetPara(217, glngSys)) <> 0
        '是否启用参数：主刀医师达到手术等级无需审核
        mbln手术等级管理 = Val(zlDatabase.GetPara(254, glngSys)) <> 0
        '如果启用了授权管理，则手术分级管理的检查不使用,门诊也不使用
        If mbln手术授权管理 Or mint调用类型 = 1 Then mbln手术分级管理 = False
    End If
    If mint场合 = 0 Then
        If mint调用类型 = 1 Then
            mbytSize = zlDatabase.GetPara("字体", glngSys, pm门诊医生站, "0")
        Else
            mbytSize = zlDatabase.GetPara("字体", glngSys, pm住院医生站, "0")
        End If
    ElseIf mint场合 = 1 Then
        mbytSize = zlDatabase.GetPara("字体", glngSys, pm住院护士站, "0")
    Else
        mbytSize = zlDatabase.GetPara("字体", glngSys, pm医技工作站, "0")
    End If
    '字体设置，因为 RichTextBox的特殊性，字体设置放到前面进行调用
    Call SetControlFontSize(Me, mbytSize)
    '读取附项
    Call Init申请附项
    mbln附项 = Not mrsAppend Is Nothing
    If mbln附项 Then mbln附项 = mrsAppend.State = 1
    If mbln附项 Then mbln附项 = mrsAppend.RecordCount > 0
    
    '初始化表格样式
    If mintType = 0 Then
        If Not Init检查组合 Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "麻醉"
        If Not Init手术项目 Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "标本"
        With cbo标本
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init检验组合 Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    ElseIf mintType = 5 Then
        vsExt.Visible = False
        fraBorder(3).Visible = False
        fraBorder(4).Visible = False
        Me.Height = Me.Height - 800
    End If
    If mbln附项 Then
        Me.Height = Me.Height + (lblAppend.Height + rtfAppend.Height + fraBorder(4).Height * 3)
    End If
    If mbytUseType = 1 Then
        vsExt.Editable = flexEDNone
        If mintType = 4 Then cmd.Enabled = False  '121475
         '允许修改检查部位及方法 mbln检查部位-检查项目有效性检查时，会用到vsExt.Editable属性,对于不要求设置部位的检查项目,应该禁止其编辑
        If mintType = 0 And mbln检查部位 Then vsExt.Editable = flexEDKbdMouse
        txtData.Enabled = False
        cmdData.Visible = False
        cbo标本.Enabled = False
    End If
    
    '其他处理
    If mintType = 0 Then
        If vsExt.Rows = vsExt.FixedRows + 1 Then
            If vsExt.Editable = flexEDNone Then
                '没有设置部位时，如果不需要确定床旁术中，也不需要输入附项，则自动确认
                If Not mbln附项 And Not optMode(0).Enabled Then Call cmdOK_Click: Exit Sub
            ElseIf vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                '只有一个部位，且部位只有一个方法可选时，自动确认
                
                '只有一个部位，自动选中该部位
                vsExt.Cell(flexcpData, vsExt.FixedRows, 1) = 1
                Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1) = img16.ListImages("c1").Picture
                '如果没有默认方法，只有一个方法也选中
                str方法 = GetOnlyOneMethod(vsExt.Cell(flexcpData, vsExt.FixedRows, 2))
                If vsExt.TextMatrix(vsExt.FixedRows, 2) = "" And str方法 <> "" Then
                    vsExt.TextMatrix(vsExt.FixedRows, 2) = str方法
                End If
                If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" Then vsExt.TabStop = False
                
                '只有一个方法可选时，如果不需要输入申请附项，则界面也不弹出
                If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" And str方法 <> "" Then
                    If Not mbln附项 Then Call cmdOK_Click: Exit Sub
                End If
            End If
        End If
    ElseIf mintType = 4 Then
        '检验输入的特殊处理
        If Not mbln技师站 Then
            blnMulti = Val(zlDatabase.GetPara(84, glngSys)) = 1 '是否允许一条医嘱申请多个检验项目
            If Len(Trim(mstrExtData)) > 0 Then
                If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                    vsExt.Enabled = False
                    '如果只有一个标本则不显示本窗口
                    If cbo标本.ListCount < 2 And Not mbln附项 Then cmdOK_Click: Exit Sub
                End If
            End If
        End If
    End If
    
    Call Grid.SetFontSize(vsExt, IIF(mbytSize = 0, 9, 12))
    
    '恢复个性化
    lngBaseHeight = Me.Height
    Call RestoreWinState(Me, App.ProductName, mintType)
    
     '10.26.80增加规格显示后，以前个性化保存的高度可能不够
    If Me.Height < lngBaseHeight Then
        Me.Height = lngBaseHeight
    End If
    
    '窗体定位
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    
End Sub

Private Function Init手术项目() As Boolean
'功能：初始化手术表格格式及数据
'参数：mstrExtData=包含附加手术及麻醉项目的信息,其中可能没有附加手术；为空时表示新输入手术项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng麻醉ID As Long
    Dim arr手术IDs As Variant, str手术IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    str手术IDs = CStr(Split(strSQL, ";")(0))
    lng麻醉ID = Val(Split(strSQL, ";")(1))
    
    '附加手术
    If str手术IDs <> "" Then
        strSQL = "Select /*+ Rule*/ A.ID,A.编码,A.名称,A.操作类型" & _
            " From 诊疗项目目录 A" & _
            " Where A.类别='F' And A.ID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[2])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " Order by A.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str手术IDs, mlng病人科室id)
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "附加手术"
    vsExt.TextMatrix(0, 1) = "规模"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str手术IDs <> "" And i <> 0 Then
        arr手术IDs = Split(str手术IDs, ",") '按照原有输入顺序
        For i = 0 To UBound(arr手术IDs)
            rsTmp.Filter = "ID=" & CStr(arr手术IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '用于恢复显示
                vsExt.TextMatrix(j, 1) = NVL(rsTmp!操作类型, 0)
            End If
        Next
    End If
    
    '麻醉项目
    If lng麻醉ID <> 0 Then
        strSQL = "Select A.ID,A.编码,A.名称,操作类型 From 诊疗项目目录 A Where A.类别='G' And A.ID=[1]" & _
                " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[2])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng麻醉ID, mlng病人科室id)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text '用于恢复显示
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init手术项目 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检查组合() As Boolean
'功能：初始化检查部位表格格式及数据
'参数：mstrExtData=包含检查部位的信息,为空时表示新输入检查组合项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long, i As Integer
    Dim str类型 As String, str名称 As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    Dim Y As Long, str方法 As String
    
    On Error GoTo errH
    
    '读取检查项目基本信息
    strSQL = "Select 名称,操作类型,执行标记 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID)
    If mint服务对象 = 2 And NVL(rsTmp!执行标记, 0) = 1 Then
        '床旁术中执行标记
        optMode(0).Visible = True: optMode(1).Visible = True: optMode(2).Visible = True
        optMode(0).Enabled = True: optMode(1).Enabled = True: optMode(2).Enabled = True
        If UBound(Split(mstrExtData, vbTab)) >= 1 Then
            optMode(Val(Split(mstrExtData, vbTab)(1))).Value = True
        End If
    End If
    str类型 = rsTmp!操作类型
    str名称 = rsTmp!名称
        
    '读取检查部位信息
    strSQL = "Select B.分组,A.部位,A.方法,A.默认,B.备注,B.方法 as 检查方法 From 诊疗项目部位 A,诊疗检查部位 B" & _
        " Where A.类型=B.类型 And A.部位=B.名称 And A.项目ID=[1] And A.类型=[2] Order by B.分组,B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID, str类型)
    blnNone = rsTmp.EOF
    mbln检查部位 = Not blnNone
'    If rsTmp.EOF Then
'        '如果该检查项目还没有设置检查部位,则以所有的供选择
'        strSQL = "Select 分组,名称 as 部位,Null as 方法,Null as 默认,备注,方法 as 检查方法 From 诊疗检查部位 Where 类型=[1] Order by 分组,编码"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类型)
'        If rsTmp.EOF Then
'            MsgBox "该项目的检查类型""" & str类型 & """下面没有设置任何检查部位，请先到检查部位管理中进行设置。", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    With vsExt
        '显示基准的部位及默认方法
        If blnNone Then
            .HighLight = flexHighlightNever
            .Editable = flexEDNone
            .TabStop = False
        Else
            .HighLight = flexHighlightAlways
            .Editable = flexEDKbdMouse
        End If
        .WordWrap = True
        .FocusRect = flexFocusNone
        .BackColorSel = &HFFCC99
        .ForeColorSel = &H0&
        .FixedRows = 1: .FixedCols = 0
        .Rows = .FixedRows + 1: .Cols = 4
        .MergeCellsFixed = flexMergeFree: .MergeRow(0) = True
        .MergeCells = flexMergeFree: .MergeCol(0) = True
        
        If str类型 = "病理" Then
            .TextMatrix(0, 0) = "标本名称"
            .TextMatrix(0, 1) = "标本名称"
            .TextMatrix(0, 2) = "材料类别"
        Else
            .TextMatrix(0, 0) = "检查部位"
            .TextMatrix(0, 1) = "检查部位"
            .TextMatrix(0, 2) = "检查方法"
        End If
        
        .TextMatrix(0, 3) = "备注"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!部位 Then
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!分组)
                .TextMatrix(.Rows - 1, 1) = rsTmp!部位
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(NVL(rsTmp!检查方法)) '供方法选择器使用
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!备注)
            End If
            If NVL(rsTmp!默认, 0) = 1 Then '以"方法名1,方法名2,..."的方式显示部位检查方法
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & NVL(rsTmp!方法)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            rsTmp.MoveNext
        Loop
        
        '修改时套入已有的内容
        '  如果为空，也可能是以前的单部位检查项目，这时要以新增的方式重新选择部位
        '  或者对于以前的单部位项目，强行传入以前的部位(没有方法)，现还可能有同名部位
        If mstrExtData <> "" Then
            arrData = Split(Split(mstrExtData, vbTab)(0), "|")
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                str方法 = ""
                If lngIdx <> -1 Then
                    '检查方法有没有不存在的
                    For Y = 0 To UBound(Split(Split(arrData(i), ";")(1), ","))
                        If InStr(.Cell(flexcpData, lngIdx, 2), CStr(Split(Split(arrData(i), ";")(1), ",")(Y))) = 0 Then
                            strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0) & "(" & Split(Split(arrData(i), ";")(1), ",")(Y) & ")"
                        Else
                            str方法 = str方法 & "," & Split(Split(arrData(i), ";")(1), ",")(Y)
                        End If
                    Next
                    '该部位的方法:可能以前的数据只有部位没有方法
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Mid(str方法, 2)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    .Cell(flexcpData, lngIdx, 1) = 1 '表明该部位已选择
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                Else
                    '该部位设置已不存在
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        '确定表格尺寸
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        Me.Width = lngIdx + 90
        
        .Height = (.Rows - 1) * (.RowHeightMin + 15) + .RowHeight(0) + 60
        If Not blnNone Then
            If .Height < 1590 Then .Height = 1590 '最少5行部位
            If .Height > 2865 + 50 Then .Height = 2865 + 50 '最多10行部位
        End If
    End With
    
    Me.Height = (vsExt.Height + 90) + cmdOK.Height + (cmdOK.Height * 0.65)
    
    '已不存在的部位提示
    If strNoneRegion <> "" Then
        If str类型 = "病理" Then
            MsgBox "以下病理标本在项目设置中已不存在：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "以下检查部位或方法在项目设置中已不存在或不适用：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        End If
    End If
    
    Init检查组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Init申请附项() As Boolean
'功能：请取项目的单据申请附项
'返回：对应的单据定义了早请附项时返回True
    Dim strSQL As String, lngIdx As Long
    Dim arrData As Variant, strData As String
    Dim strNoneAppend As String, strHaveAppend As String
    Dim arrSub As Variant, i As Long
    Dim str主刀医生 As String
    Dim str主刀医生科室 As String
    Dim blnHave主刀医生 As Boolean
    Dim rsTmp As ADODB.Recordset
    
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    strSQL = "Select C.项目,C.内容,C.要素ID,C.必填,d.中文名,decode(D.表示法,4,D.数值域,NULL) as 数值域,c.只读" & _
        " From 病历单据应用 A,病历文件列表 B,病历单据附项 C,诊治所见项目 D" & _
        " Where A.诊疗项目ID=[1] And A.应用场合=[2]" & _
        " And A.病历文件ID=B.ID And B.种类=7 And B.ID=C.文件ID And c.要素id=d.id(+)" & _
        " Order by C.排列"
    
    On Error GoTo errH
    Set mrsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID, mint服务对象)
    If Not mrsAppend.EOF Then
    
        mrsAppend.Filter = "中文名='主刀医生'"
        If mrsAppend.RecordCount > 0 Then
            blnHave主刀医生 = True
            str主刀医生 = GetAppendItemValue(mrsAppend!项目, NVL(mrsAppend!要素ID, 0), mrsAppend!中文名 & "")
            str主刀医生 = Trim(str主刀医生)
        End If
        
        mrsAppend.Filter = "中文名='主刀医生科室'"
        If mrsAppend.RecordCount > 0 Then
            If str主刀医生 <> "" And blnHave主刀医生 Then
                strSQL = "Select a.名称, c.缺省 From 部门表 A, 人员表 B, 部门人员 C, 部门性质说明 D" & _
                    " Where a.Id = c.部门id And b.Id = c.人员id And a.Id = d.部门id And d.工作性质 = '临床' And" & _
                    " (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & _
                    " (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null) And b.姓名 = [1]"
                    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str主刀医生)
                
                If Not rsTmp.EOF Then
                    str主刀医生科室 = rsTmp!名称 & ""
                    rsTmp.Filter = "缺省 = 1"
                    If rsTmp.RecordCount > 0 Then str主刀医生科室 = rsTmp!名称 & ""
                End If
            End If
        End If
        
        mrsAppend.Filter = 0
        arrData = Split(mstrAppend, "<Split1>")
        With rtfAppend
            Do While Not mrsAppend.EOF
                '确定附项内容
                strData = ""
                If mrsAppend!数值域 & "" <> "" Then mstr单选项目 = mstr单选项目 & "," & mrsAppend!中文名
                If mstrAppend <> "" Then
                    '修改时，保持原有内容
                    For i = 0 To UBound(arrData)
                        arrSub = Split(arrData(i), "<Split2>")
                        If arrSub(0) = mrsAppend!项目 Then
                            strData = arrSub(3)
                            If strData = "" And UBound(arrSub) >= 4 Then
                                '当对复制或成套产生的医嘱进行修改时，申请附项也要取缺省值
                                If Val(arrSub(4)) = 1 Then
                                    If Not IsNull(mrsAppend!内容) Then
                                        strData = mrsAppend!内容
                                    ElseIf mlng病人ID <> 0 Then
                                        strData = GetAppendItemValue(mrsAppend!项目, NVL(mrsAppend!要素ID, 0), mrsAppend!中文名 & "")
                                    End If
                                End If
                            End If
                            
                            '存在的附项
                            strHaveAppend = strHaveAppend & "," & arrSub(0)
                            strNoneAppend = Replace(strNoneAppend & ",", "," & arrSub(0) & ",", ",")
                            If Right(strNoneAppend, 1) = "," Then strNoneAppend = Left(strNoneAppend, Len(strNoneAppend) - 1)
                        ElseIf InStr(strNoneAppend & ",", "," & arrSub(0) & ",") = 0 _
                             And InStr(strHaveAppend & ",", "," & arrSub(0) & ",") = 0 Then
                            strNoneAppend = strNoneAppend & "," & arrSub(0) '先记到没有的附项中
                        End If
                    Next
                Else
                    '新增时，使用预定义内容或从病人数据中提取
                    If Not IsNull(mrsAppend!内容) Then
                        strData = mrsAppend!内容
                    ElseIf mlng病人ID <> 0 Then
                        If mrsAppend!中文名 & "" = "主刀医生" Then
                            strData = str主刀医生
                        ElseIf mrsAppend!中文名 & "" = "主刀医生科室" And blnHave主刀医生 And str主刀医生 <> "" Then
                            strData = str主刀医生科室
                        Else
                            strData = GetAppendItemValue(mrsAppend!项目, NVL(mrsAppend!要素ID, 0), mrsAppend!中文名 & "")
                        End If
                    End If
                End If
                
                '将该项显示在RTF中:保护文本后第一个位置不能直接录入汉字,因此先多加一个不保护的空格
                .SelText = IIF(.Text = "", "", vbCrLf) & mrsAppend!项目 & "： " & strData
                lngIdx = .Find(mrsAppend!项目 & "：", , , rtfNoHighlight Or rtfMatchCase)
                '如果是主刀医生，则读取其手术登记
                If mrsAppend!中文名 & "" = "主刀医生" And mbln手术分级管理 Then
                    mstr手术等级 = GetDoctorLevel(strData)
                End If
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(mrsAppend!项目 & "：")
                    .SelBold = True
                    .SelIndent = 100
                    .SelProtected = True
                End If
                If Val(mrsAppend!只读 & "") = 1 And strData <> "" Then
                    lngIdx = lngIdx + Len(mrsAppend!项目 & "： ")
                    .SelStart = lngIdx
                    .SelLength = Len(strData)
                    .SelProtected = True
                End If
                .SelStart = Len(.Text)
                
                mrsAppend.MoveNext
            Loop
            mstr单选项目 = Mid(mstr单选项目, 2)
            
            '光标定位在第一个输入附项
            mrsAppend.MoveFirst
            lngIdx = .Find(mrsAppend!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(mrsAppend!项目 & "：") + 1
            
            '确定RTF控件尺寸
            .Height = (mrsAppend.RecordCount + 2) * 250 + 30
            If .Height < 3 * 265 + 30 Then .Height = 3 * 250 + 30 '最少3行
            If .Height > 8 * 265 + 30 Then .Height = 8 * 250 + 30 '最多8行
        End With
        
        lblAppend.Visible = True: rtfAppend.Visible = True: fraBorder(4).Visible = True
        Init申请附项 = True
    End If
    
    '已不存在的申请项目提示
    If strNoneAppend <> "" Then
        MsgBox "以下附项在项目对应的单据设置中已不存在：" & vbCrLf & Mid(strNoneAppend, 2), vbInformation, gstrSysName
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOrderInspectInfo(ByVal lng病人ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng就诊ID As Long) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng病人ID, lng就诊ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng病人ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Private Function GetAppendItemValue(ByVal str项目 As String, ByVal lng要素ID As Long, ByVal str中文名 As String) As String
'功能：获取指定的申请附项值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    Dim lng就诊ID As Long
    Dim intType As Integer '1-门诊，2－住院
    
    On Error GoTo errH
    
    If TypeName(mvar就诊ID) = "String" Then
        intType = 1
    Else
        intType = 2
    End If
    
    If intType = 1 And strText = "" Then
        '如果诊断，门诊从未保存的已录入诊断中提取
        If str项目 Like "*诊断" And strText = "" And mstrDiagnosis <> "" Then
            strText = mstrDiagnosis
        End If
    End If
    
    '就诊要素的提取门诊住院区别对待
    '门诊先取已有医嘱中的附项目，再取病历中的，2-住院先取病历，再取医嘱
    If intType = 1 And strText = "" Then
        '4.未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(TypeName(mvar就诊ID) = "String", " And A.挂号单=[2]", " And A.主页ID=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4" & _
            " Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, CStr(mvar就诊ID), Val(mvar就诊ID), mint婴儿, str项目)
        If Not rsTmp.EOF Then strText = nvl(rsTmp!内容)
    End If
    
    
    '如果有对应要素，从要素提取函数读取
    If lng要素ID <> 0 And strText = "" Then
        '先老版，再新版
        If TypeName(mvar就诊ID) = "String" Then '门诊
            strSQL = "Select Zl_Replace_Element_Value(B.中文名,[1],A.ID,1) as 内容" & _
                " From 病人挂号记录 A,诊治所见项目 B Where A.NO=[2] And B.ID=[3] And a.记录性质=1 And a.记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, CStr(mvar就诊ID), lng要素ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(中文名,[1],[2],2) as 内容 From 诊治所见项目 Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, Val(mvar就诊ID), lng要素ID)
        End If
        If Not rsTmp.EOF Then strText = nvl(rsTmp!内容)
        If strText = "" Then
            
            If TypeName(mvar就诊ID) = "String" Then
                strSQL = "select a.id From 病人挂号记录 A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(mvar就诊ID))
                lng就诊ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng就诊ID = Val(mvar就诊ID)
                intType = 2
            End If
            strText = GetOrderInspectInfo(mlng病人ID, str中文名, intType, lng就诊ID)
        End If
    End If
    
    '未取到或未对应要素的，从病人之前未保存的医嘱中提取,以最后填写的为准
    If strText = "" And mstrAdvItem <> "" Then
        arrItem = Split(mstrAdvItem, "<Split1>")
        For i = 0 To UBound(arrItem)
            If Split(arrItem(i), "<Split2>")(0) = str项目 Then
                strText = Split(arrItem(i), "<Split2>")(1): Exit For
            End If
        Next
    End If
    
    If strText = "" And intType = 2 Then
        '未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(TypeName(mvar就诊ID) = "String", " And A.挂号单=[2]", " And A.主页ID=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4" & _
            " Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, CStr(mvar就诊ID), Val(mvar就诊ID), mint婴儿, str项目)
        If Not rsTmp.EOF Then strText = nvl(rsTmp!内容)
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检验组合() As Boolean
'功能：初始化检验项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnLis As Boolean
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strTmp As String
    Dim colTmp As New Collection
    Dim strItemTmp As String
    Dim lng父ID As Long
    Dim Y As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = IIF(mlng项目ID <> 0, mlng项目ID, "") & ";"
    strItems = CStr(Split(strSQL, ";")(0))
    Me.txtData.Text = Split(strSQL, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        '判断是否是新版LIS模式的组合项目
        If Not gobjLIS Is Nothing Then
            blnLis = gobjLIS.CheckLisSate
        End If
        If mblnNewLIS And blnLis Then
            strLIS = " Union All" & vbNewLine & _
                    "       Select e.Id, e.编码, e.名称, e.操作类型, e.试管编码, 检验组合项目.编码 As 序号,检验组合项目.id as 父ID " & vbNewLine & _
                    "       From 检验组合项目, 检验报告项目 C, 检验报告项目 D, 诊疗项目目录 E" & vbNewLine & _
                    "       Where 检验组合项目.Id = c.诊疗项目id And c.报告项目id = d.报告项目id And d.诊疗项目id = e.Id And e.组合项目 <> 1 And 检验组合项目.Id <> e.Id"
            '分解子项
            For i = 0 To UBound(Split(strItems, ","))
                strTmp = Split(strItems, ",")(i)
                If InStr(strTmp, "|") > 0 Then
                    colTmp.Add Mid(strTmp, InStr(strTmp, "|") + 1), "_" & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                    strItemTmp = strItemTmp & "," & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                Else
                    strItemTmp = strItemTmp & "," & strTmp
                End If
            Next
            strItems = Mid(strItemTmp, 2)
            Me.Height = Me.Height + 1200
            vsExt.Height = vsExt.Height + 1200
        End If
        strSQL = "Select * From (With 检验组合项目 As (Select /*+ Rule*/ A.ID,A.编码,A.名称,A.操作类型,A.试管编码, a.编码 As 序号,null as 父ID  From 诊疗项目目录 A " & _
            " Where A.类别='C' " & _
            IIF(mint场合 = 2, "", " And Nvl(A.单独应用,0)=1") & _
            " And A.服务对象 IN(" & mint服务对象 & ",3" & IIF(mint场合 = 2, ",4", "") & ") " & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And A.ID In(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[2])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID)))" & _
            " Select * from 检验组合项目" & _
            strLIS & _
            ") Order by 序号,编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strItems, mlng病人科室id)
    End If
        
    vsExt.Clear
    If strItems <> "" Then
        vsExt.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    Else
        vsExt.Rows = 2
    End If
    vsExt.Cols = 4
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 2) = "检验项目"
    If mblnNewLIS Then
        vsExt.ColWidth(2) = 3700
        vsExt.ColWidth(0) = 300
    Else
        vsExt.ColWidth(2) = 4000
        vsExt.ColHidden(0) = True
    End If
    vsExt.ColHidden(1) = True
    vsExt.ColHidden(3) = True
    vsExt.FixedAlignment(2) = 4
    vsExt.ColAlignment(2) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If strItems <> "" Then
        If Not rsTmp.EOF Then
            arrItems = Split(strItems, ",") '按照原有输入顺序
            For i = 0 To UBound(arrItems)
                rsTmp.Filter = "ID=" & arrItems(i)
                If Not rsTmp.EOF Then
                    Y = vsExt.FindRow(CLng(rsTmp!ID))
                    '重复的指标不加入
                    If Y = -1 Then
                        j = j + 1
                        vsExt.RowData(j) = CLng(rsTmp!ID)
                        '主项默认勾选，且不能取消
                        vsExt.TextMatrix(j, 0) = " "
                        vsExt.Cell(flexcpBackColor, j, 0) = &H8000000F
                        vsExt.TextMatrix(j, 2) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                        vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                        vsExt.TextMatrix(j, 1) = NVL(rsTmp!操作类型)
                        vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!试管编码)) '用于同类输入限制
                        vsExt.TextMatrix(j, 3) = 0   '父项
    '                    If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                    End If
                    If mblnNewLIS Then
                        lng父ID = CLng(rsTmp!ID)
                        rsTmp.Filter = "父ID=" & CLng(rsTmp!ID)
                        Do While Not rsTmp.EOF
                            Y = vsExt.FindRow(CLng(rsTmp!ID))
                            '重复的指标不加入
                            If Y = -1 Then
                                j = j + 1
                                vsExt.RowData(j) = CLng(rsTmp!ID)
                                On Error Resume Next
                                strItemTmp = ""
                                strItemTmp = colTmp("_" & lng父ID)
                                On Error GoTo errH
                                If InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                ElseIf strItemTmp = "" And mblnNew Then  '第一次进入默认勾选
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                Else
                                    vsExt.Cell(flexcpChecked, j, 0) = 2
                                End If
                                '子项缩进
                                vsExt.TextMatrix(j, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                                vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                                vsExt.TextMatrix(j, 1) = NVL(rsTmp!操作类型)
                                vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!试管编码)) '用于同类输入限制
        '                       If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                                vsExt.TextMatrix(j, 3) = 1    '子项
                            Else
                                '如果重复的指标勾选了前面的指标未勾选，则删除前面的指标加载后面的指标
                                On Error Resume Next
                                strItemTmp = ""
                                strItemTmp = colTmp("_" & lng父ID)
                                On Error GoTo errH
                                If vsExt.Cell(flexcpChecked, Y, 0) = 1 And InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                    vsExt.RemoveItem Y
                                    vsExt.AddItem ""
                                    vsExt.RowData(j) = CLng(rsTmp!ID)
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                    '子项缩进
                                    vsExt.TextMatrix(j, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                                    vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                                    vsExt.TextMatrix(j, 1) = NVL(rsTmp!操作类型)
                                    vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!试管编码)) '用于同类输入限制
            '                       If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                                    vsExt.TextMatrix(j, 3) = 1    '子项
                                End If
                            End If
                            rsTmp.MoveNext
                        Loop
                    End If
                End If
            Next
        End If
        If j > 0 Then vsExt.Rows = j + 1
    End If
    
    vsExt.Row = 1: vsExt.Col = 2
    Init检验组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String, lngItemCount As Long
    
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSQL = "Select 名称 From 诊疗检验标本" & _
            "     Where (Instr(Nvl(适用性别,'不限'),'男')=0 And Instr(Nvl(适用性别,'不限'),'女')=0" & _
            "         Or Instr(Nvl([1],'不限'),'男')=0 And Instr(Nvl([1],'不限'),'女')=0" & _
            "         Or Instr([1],'男')>0 And Instr(适用性别,'男')>0 Or Instr([1],'女')>0 And Instr(适用性别,'女')>0)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr性别)
    Else
        strSQL = "Select /*+ Rule*/ 标本类型,Sum(1) From (" & _
            "   Select Distinct A.ID,B.名称 As 标本类型" & _
            "   From 诊疗项目目录 A,诊疗检验标本 B,检验项目参考 C,检验报告项目 D" & _
            "   Where A.ID=D.诊疗项目ID(+) And D.报告项目ID=C.项目ID(+)" & _
            "        And (NVL(C.标本类型,'') Is Null Or NVL( C.标本类型,'')=B.名称) " & _
            "       And A.ID In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            "       And (Instr(Nvl(B.适用性别,'不限'),'男')=0 And Instr(Nvl(B.适用性别,'不限'),'女')=0" & _
            "         Or Instr(Nvl([3],'不限'),'男')=0 And Instr(Nvl([3],'不限'),'女')=0" & _
            "         Or Instr([3],'男')>0 And Instr(B.适用性别,'男')>0 Or Instr([3],'女')>0 And Instr(B.适用性别,'女')>0)" & _
            "           And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " ) Group By 标本类型 Having Sum(1)=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strTmp, lngItemCount, mstr性别, mlng病人科室id)
    End If
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "未设置检验标本，请到字典管理工具中设置。", _
            lngItemCount = 1, "选取的检验项目未定义检验标本，请先到检验项目管理中设置", _
            lngItemCount > 1, "选取的检验项目的检验标本与其他项目的不一致，请先到检验项目管理中设置"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo标本
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mintType)

    mlngHwnd = 0
    mint婴儿 = 0
    mlng病人ID = 0
    mlng病人科室id = 0
    mvar就诊ID = Empty
    mstr性别 = ""
    mint场合 = 0
    mintType = 0
    mbytUseType = 0
    mint期效 = 0
    mint服务对象 = 0
    mint调用类型 = 0
    mblnNewLIS = False
    mblnNew = False
    mlng项目ID = 0
    mstrAdvItem = ""
    mstrDiagnosis = ""
    mstr手术等级 = ""
    Set mrsAppend = Nothing

    Set mfrmParent = Nothing
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + x < 4140 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        ElseIf Index = 4 Then
            If vsExt.Height + Y < 1000 Or vsExt.Height + Y > Me.Height * 0.7 Then Exit Sub
            vsExt.Height = vsExt.Height + Y
            Call Form_Resize
        End If
    End If
End Sub

Private Function CursorInItem(Optional ByRef str项目 As String, Optional ByRef bln大于 As Boolean) As Boolean
'功能：判断当前光标是否在某个项目标题后面
    Dim lngLoc As Long, i As Long
    
    With rtfAppend
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            lngLoc = .Find(mrsAppend!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngLoc = -1 Then
                lngLoc = InStr(rtfAppend.Text, mrsAppend!项目 & "：")
                lngLoc = lngLoc - 1
            End If
            If lngLoc <> -1 Then
                lngLoc = lngLoc + Len(mrsAppend!项目 & "：")
                If .SelStart >= lngLoc And InStr(Mid(.Text, lngLoc, IIF(.SelStart - lngLoc < 0, 0, .SelStart - lngLoc)), vbCrLf) = 0 Then
                    bln大于 = True
                    str项目 = NVL(mrsAppend!中文名)
                End If
                If .SelStart = lngLoc Then CursorInItem = True: Exit Function
            End If
            mrsAppend.MoveNext
        Next
    End With
End Function

Private Sub imgSentence_Click()
    Dim strSentence As String
    Dim str检查部位 As String, str检查方法 As String
    
    If txtSentence.Tag = "输入医生" Then
        Call FindDoctor("")
    ElseIf txtSentence.Tag = "输入科室" Then
        Call FindDept("")
    ElseIf txtSentence.Tag = "选择单选项" Then
        Call FindItem
    Else
        Call Get检查部位方法(str检查部位, str检查方法)
        
        strSentence = frmSentenceSel.ShowMe(Me, mint服务对象, mlng病人ID, mvar就诊ID, mlng项目ID, str检查部位, str检查方法, , , , mobjEmrInterface)
        If strSentence <> "" Then
            rtfAppend.SelText = strSentence
            Call HideWordInput(True)
        End If
    End If
End Sub

Private Sub rtfAppend_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub rtfAppend_KeyPress(KeyAscii As Integer)
    Dim str项目性质 As String
    Dim bln等于 As Boolean
    
    If KeyAscii = 13 Then
        If txtSentence.Tag <> "输入医生" Then
            '连续两次回车光标跳转
            With rtfAppend
                If .SelStart - 1 > 0 Then
                    If Mid(.Text, .SelStart - 1, 2) = vbCrLf Then
                        KeyAscii = 0
                        Call zlCommFun.PressKey(vbKeyBack)
                        If InStr(Mid(.Text, .SelStart + 1), "： ") > 0 Then
                            Call zlCommFun.PressKey(vbKeyDown)
                            Call zlCommFun.PressKey(vbKeyEnd)
                        Else
                            Call zlCommFun.PressKey(vbKeyTab)
                        End If
                    End If
                End If
            End With
        Else
            KeyAscii = 0
            bln等于 = CursorInItem(str项目性质, False)
            If str项目性质 = "助手医生" Or str项目性质 = "主刀医生" Then
                With rtfAppend
                    .SelStart = IIF(InStr(Mid(.Text, .SelStart + 1), "： ") > 0, .SelStart + 1 + InStr(Mid(.Text, .SelStart + 1), "： ") + 1, .SelStart - IIF(bln等于, 0, 1))
                    .SelLength = IIF(InStr(Mid(.Text, .SelStart), vbCrLf) - 1 > 0, InStr(Mid(.Text, .SelStart), vbCrLf) - 1, Len(.Text))
                End With
            End If
        End If
    ElseIf KeyAscii = 8 And txtSentence.Tag <> "输入医生" Then
        '不允许删除标题后的" "
        With rtfAppend
            If .SelLength = 0 And .SelStart > 0 Then
                If Mid(.Text, .SelStart, 1) = "：" Then
                    If CursorInItem Then
                         If Mid(.Text, .SelStart + 1, 1) <> " " Then .SelText = " "
                    End If
                End If
            End If
        End With
    Else
        If txtSentence.Tag = "输入医生" Then KeyAscii = 0: Call rtfAppend_SelChange
        If txtSentence.Tag = "选择单选项" Then KeyAscii = 0: Call rtfAppend_SelChange
    End If
End Sub

Private Sub rtfAppend_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub rtfAppend_SelChange()
    Dim str项目 As String
    Dim bln大于 As Boolean
    Dim bln等于 As Boolean
    Dim bytType  As Byte 'bytType 0-空为词句，1-输入人员，2-输入科室
    Dim str数值域 As String
    
    With rtfAppend
        If .Visible And .SelLength = 0 And .SelStart > 0 Then
            bln等于 = CursorInItem(str项目, bln大于)
            If str项目 = "助手医生" Or str项目 = "主刀医生" Then
                bytType = 1
            ElseIf str项目 = "主刀医生科室" Then
                bytType = 2
            ElseIf InStr("," & mstr单选项目 & ",", "," & str项目 & ",") > 0 And mstr单选项目 <> "" Then
                bytType = 3
                mrsAppend.Filter = "中文名='" & str项目 & "'"
                If mrsAppend.RecordCount > 0 Then mrsAppend.MoveFirst: str数值域 = mrsAppend!数值域 & ""
                 mrsAppend.Filter = 0
            End If
            
            If bln大于 And bytType <> 0 And Not picSentence.Visible Then
                .SelStart = IIF(InStrRev(Mid(.Text, 1, .SelStart + 1), "： ") > 0, InStrRev(Mid(.Text, 1, .SelStart + 1), "： ") + 1, .SelStart - IIF(bln等于, 0, 1))
                .SelLength = IIF(InStr(Mid(.Text, .SelStart), vbCrLf) - 1 > 0, InStr(Mid(.Text, .SelStart), vbCrLf) - 1, Len(.Text))
                Call ShowWordInput(bytType, str数值域)
            Else
                If Mid(.Text, .SelStart, 2) = "： " Then
                    '光标不允许定位到标题后的" "上
                    If CursorInItem() Then .SelStart = .SelStart + 1
                ElseIf Mid(.Text, .SelStart, 1) = "`" Then
                    '词句输入快捷特殊处理
                    '用vbBack达不到效果
                    .SelStart = .SelStart - 1
                    .SelLength = 1: .SelText = ""
                    Call ShowWordInput
                Else
                    If Not (str项目 = "助手医生" Or str项目 = "主刀医生") Then txtSentence.Tag = ""
                End If
            End If
        End If
    End With
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSQL As String, str性别 As String
    Dim strLike As String, blnCancel As Boolean
    
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    Else
        str性别 = "0"
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '手术可以不输入麻醉项目
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '优化
        strLike = mstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '输入麻醉项目
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='G' And A.服务对象 IN([3],3)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                    " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[5])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " Order by A.编码"
            vRect = zlControl.GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "麻醉项目", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint服务对象, mint简码 + 1, mlng病人科室id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '检验标本
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'功能：恢复显示原内容
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub txtSentence_GotFocus()
    Call zlControl.TxtSelAll(txtSentence)
End Sub

Private Function GetDoctorLevel(ByVal str姓名 As String) As String
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select 手术等级 From 人员表 Where 姓名=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, str姓名)
    If rsTmp.RecordCount > 0 Then
        GetDoctorLevel = rsTmp!手术等级 & ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindDoctor(ByVal strTmp As String) As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, str项目性质 As String, blnDo As Boolean
    Dim lngStart As Long
    Dim lng人员id As Long
    Dim str部门 As String
    
    On Error GoTo errH
    strInput = Trim(UCase(strTmp))   '传入的值存在前缀空格
    strSQL = "Select A.ID,A.编号,A.姓名,A.简码,A.手术等级" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编号"
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call CursorInItem(str项目性质, False)
            If str项目性质 = "助手医生" Then
                If MsgBox("没有找到匹配的医生，你确定要输入没有建立人员档案的医生吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                    strTmp = strInput
                Else
                    blnDo = False
                End If
            Else
                Call MsgBox("没有找到匹配的医生!", vbInformation, gstrSysName)
                blnDo = False
            End If
        End If
    Else
        blnDo = True
        strTmp = rsTmp!姓名 & ""
        lng人员id = rsTmp!ID
        Call CursorInItem(str项目性质, False)
        If str项目性质 = "主刀医生" And mbln手术分级管理 Then
            mstr手术等级 = rsTmp!手术等级 & ""
        End If
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        
        strSQL = "Select b.名称, a.缺省 From 部门人员 A, 部门表 B, 部门性质说明 C" & _
            " Where a.部门id = b.Id And b.Id = c.部门id And c.工作性质 = '临床' And a.人员id = [1]" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng人员id)
        If Not rsTmp.EOF Then
            str部门 = rsTmp!名称 & ""
            rsTmp.Filter = "缺省 = 1"
            If rsTmp.RecordCount > 0 Then str部门 = rsTmp!名称 & ""
        End If
         
        lngStart = rtfAppend.SelStart
        Call Do手术附项内容("主刀医生科室", True, str部门)
        rtfAppend.SelStart = lngStart
 
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindItem() As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, strTmp As String, blnDo As Boolean
    
    strInput = Replace(imgSentence.Tag, ";", ",")
    If strInput = "" Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select rownum as ID, Column_Value as 选择项 From Table(Cast(f_Str2List([1]) As zlTools.t_StrList))"
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            blnDo = False
        End If
    Else
        blnDo = True
        strTmp = rsTmp!选择项 & ""
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindDept(ByVal strTmp As String) As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, str项目性质 As String, blnDo As Boolean
    Dim strDoctor As String
    
    On Error GoTo errH
    strInput = Trim(UCase(strTmp))   '传入的值存在前缀空格
    strDoctor = Do手术附项内容("主刀医生")
    strDoctor = Trim(Replace(strDoctor, vbCrLf, "")) '去掉回车和空白
    strSQL = "Select Distinct A.ID,A.编码,A.名称 as 科室,A.简码 From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And a.Id = b.部门id And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.工作性质='临床'" & _
        IIF(strDoctor <> "", " And a.id in (select x.部门id from 部门人员 X, 人员表 Y where x.人员id=y.id and y.姓名=[3])", "") & _
        " Order by A.编码"
    
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", strDoctor)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的科室!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        strTmp = rsTmp!科室 & ""
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Do手术附项内容(ByVal str项目名称 As String, Optional ByVal blnSet As Boolean, Optional ByVal strValue As String) As String
'能功：设置或者是获取相应（str项目名称）的手术申请附项内容
'参数：str项目名称 申请项目的中文名
'      blnSet true 给项目赋值，false 取值然后返回
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strData As String
    
    If rtfAppend.Visible Then
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            If mrsAppend!中文名 = str项目名称 Then
                strData = "": lngBegin = -1: lngEnd = -1
                lngBegin = rtfAppend.Find(mrsAppend!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
                If lngBegin = -1 Then
                    lngBegin = InStr(rtfAppend.Text, mrsAppend!项目 & "：")
                    lngBegin = lngBegin - 1
                End If
                If lngBegin <> -1 Then
                    lngBegin = lngBegin + Len(mrsAppend!项目 & "：")
                    If i = mrsAppend.RecordCount Then
                        lngEnd = Len(rtfAppend.Text)
                    Else
                        mrsAppend.MoveNext
                        lngEnd = rtfAppend.Find(vbCrLf & mrsAppend!项目 & "：", lngBegin, , rtfNoHighlight Or rtfMatchCase)
                        If lngEnd = -1 Then
                            lngEnd = InStr(rtfAppend.Text, vbCrLf & mrsAppend!项目 & "：")
                            lngEnd = lngEnd - 1
                        End If
                        mrsAppend.MovePrevious
                    End If
                End If
                If lngBegin <> -1 And lngEnd <> -1 Then
                    'MID函数是以1为基础，rtf是以0为基础
                    lngBegin = lngBegin + 1
                    lngEnd = lngEnd + 1
                    strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                    '去掉为解决保护文本后第一个位置不能直接录入汉字所添加的空格
                    If Left(strData, 1) = " " Then strData = Mid(strData, 2)
                    If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)
                End If
                If blnSet Then
                    rtfAppend.SelStart = lngBegin
                    rtfAppend.SelLength = lngEnd - lngBegin
                    rtfAppend.SelText = strValue
                    Exit Function
                End If
            End If
            mrsAppend.MoveNext
        Next
    End If
    Do手术附项内容 = strData
End Function

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strSentence As String, blnCancel As Boolean
    Dim str检查部位 As String, str检查方法 As String
    Dim str项目性质 As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtSentence.Tag = "输入医生" Then
            If txtSentence.Text = "" Then
                rtfAppend.SelText = ""
                Call HideWordInput(True)
                Call CursorInItem(str项目性质, False)
                If str项目性质 = "主刀医生" And mbln手术分级管理 Then
                    mstr手术等级 = ""
                End If
            Else
                Call FindDoctor(txtSentence.Text)
            End If
        ElseIf txtSentence.Tag = "输入科室" Then
            If txtSentence.Text = "" Then
                rtfAppend.SelText = ""
                Call HideWordInput(True)
            Else
                Call FindDept(txtSentence.Text)
            End If
        Else
            Call Get检查部位方法(str检查部位, str检查方法)
            
            strSentence = frmSentenceSel.ShowMe(Me, mint服务对象, mlng病人ID, mvar就诊ID, mlng项目ID, str检查部位, str检查方法, txtSentence.Text, picSentence.Hwnd, blnCancel, mobjEmrInterface)
            If strSentence <> "" Then
                rtfAppend.SelText = strSentence
                Call HideWordInput(True)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的词句。", vbInformation, gstrSysName
                End If
                Call zlControl.TxtSelAll(txtSentence)
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call imgSentence_Click
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput(False) '隐藏词句输入
    End If
End Sub

Private Sub ShowWordInput(Optional ByVal bytType As Byte, Optional ByVal str数值域 As String)
'功能：显示词句输入
'参数：bytType 0-空为词句，1-输入人员，2-输入科室,3-选择单选项
    Dim vPos As PointAPI
    Dim blnLocked As Boolean
    
    imgSentence.Tag = ""
    If bytType = 1 Then
        txtSentence.Tag = "输入医生"
    ElseIf bytType = 2 Then
        txtSentence.Tag = "输入科室"
    ElseIf bytType = 3 Then
        txtSentence.Tag = "选择单选项"
        imgSentence.Tag = str数值域
        blnLocked = True
    Else
        txtSentence.Tag = ""
    End If
    txtSentence.Locked = blnLocked
    
    If rtfAppend.Visible And rtfAppend.Enabled Then
        vPos = GetCaretPos(rtfAppend.Hwnd)
        If vPos.x <> -1 And vPos.Y <> -1 Then
            If rtfAppend.Left + vPos.x + Screen.TwipsPerPixelX * 2 < rtfAppend.Left + rtfAppend.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX Then
                picSentence.Left = rtfAppend.Left + vPos.x + Screen.TwipsPerPixelX * 2
            Else
                picSentence.Left = rtfAppend.Left + rtfAppend.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX
            End If
            picSentence.Top = rtfAppend.Top + vPos.Y + Screen.TwipsPerPixelY
            If bytType <> 0 Then
                txtSentence.Text = rtfAppend.SelText
            Else
                txtSentence.Text = ""
            End If
            picSentence.Visible = True
            txtSentence.SetFocus
        End If
    End If
End Sub

Private Sub HideWordInput(ByVal blnFocus As Boolean)
'功能：隐藏词句输入
    picSentence.Visible = False
    txtSentence.Text = ""
    If blnFocus And rtfAppend.Visible And rtfAppend.Enabled Then
        rtfAppend.SetFocus
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能:显示选择按钮,并保证当前单元格可见
    Dim strKey As String, lng药名ID As Long
    
    If mblnChangeSel = True Then Exit Sub
    '保证当前单元格可见
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 1 Or mintType = 4 Then
        '显示/隐藏手术选择按钮
        If NewCol = 0 And mintType = 1 Or NewCol = 2 And mintType = 4 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(NewRow, 3) = "1" Then
                    cmd.Visible = False
                Else
                    cmd.Visible = True
                End If
            Else
                cmd.Visible = True
            End If
        Else
            cmd.Visible = False
        End If
        If cmd.Visible Then
            vsExt.FocusRect = flexFocusSolid
        Else
            vsExt.FocusRect = flexFocusLight
        End If
    End If
    
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'功能:限制某些列宽的范围
    If Row = -1 Then
        If mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见及调整按钮位置
        End If
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 0 Then
        If NewCol = 0 Or NewCol = 3 Then
            Cancel = True
            If NewRow <> OldRow Then vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
    If fraMethod.Visible Then fraMethod.Visible = False
End Sub

Private Function GetOnlyOneMethod(ByVal strMethod As String) As String
'功能：根据部位的方法定义，如果只有一个方法可选，则返回该方法
'注意：有3个符号为保留符号：< vbTab  ;  , >
    Dim strTmp As String
    
    If strMethod = "" Then Exit Function
    strTmp = strMethod
    
    strTmp = Replace(strTmp, vbTab, ";")
    strTmp = Replace(strTmp, ",", ";")
    strTmp = Replace(strTmp, ";;", ";")
    strTmp = "<spdel>" & strTmp & "<spdel>"
    strTmp = Replace(strTmp, "<spdel>;", "")
    strTmp = Replace(strTmp, ";<spdel>", "")
    strTmp = Replace(strTmp, "<spdel>", "")
    
    If InStr(strTmp, ";") = 0 Then GetOnlyOneMethod = Mid(strTmp, 2)        '去掉前首位造影标记
End Function

Private Sub vsExt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    Dim lngTmp As Long
    Dim k As Long
    Dim blnDo As Boolean

    strMethod = vsExt.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "该检查部位没有设置可供选择的检查方法。", vbInformation, gstrSysName
        Exit Sub
    End If
    With vsMethod
        .Rows = 0
        
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '表明是共选项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '第一位是造影剂标志
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    Else
                        '排斥项
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '表明是排斥项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '第一位是造影剂标志
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1为选中
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    End If
                Else
                    '共选子项
                    .RowData(.Rows - 1) = 3 '表明是共选子项
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)
                    If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        blnDo = True
                        '主项没有选择时,子项不能选择
                        For k = .Rows - 2 To 0 Step -1
                            If .RowData(k) <> 3 Then
                                If .Cell(flexcpData, k, 0) = 0 Then blnDo = False
                                Exit For
                            End If
                        Next
                    Else
                        blnDo = False
                    End If
                    
                    If blnDo Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                    End If
                End If
            Next
        Next

        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        If (vsExt.Width - 30) - (vsExt.CellLeft + 15) <= 0 Then
            For i = 0 To vsExt.Cols - 1
                lngTmp = vsExt.ColWidth(i) + lngTmp
            Next
            Me.Width = lngTmp
        End If
        
        .Width = (vsExt.Width - 30) - (vsExt.CellLeft + 15)
        
        .Left = vsExt.Left + vsExt.CellLeft + 15
        
        .Top = vsExt.Top + vsExt.CellTop + vsExt.CellHeight + 15
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        fraMethod.Top = .Top: .Top = 0
        fraMethod.Left = .Left: .Left = 0
        fraMethod.Width = .Width
        fraMethod.Height = .Height + cmdMethodOK.Height + 20
        cmdMethodOK.Top = .Height
        cmdMethodOK.Left = .Width - cmdMethodOK.Width - 20
        
        fraMethod.ZOrder
        If .Tag = "AutoPopup" Then
            fraMethod.Visible = .Rows > 1
        Else
            fraMethod.Visible = True
        End If
        If fraMethod.Visible Then .SetFocus
    End With
End Sub

Private Sub vsExt_DblClick()
    If mintType = 0 Then
        If vsExt.Editable <> flexEDNone And vsExt.MouseCol = 1 And vsExt.MouseRow >= vsExt.FixedRows Then
            Call vsExt_KeyPress(vbKeySpace)
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    If fraMethod.Visible Then fraMethod.Visible = False
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long, j As Long, k As Long, g As Long
    Dim intRow As Integer        '有效行
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strKey As String, lng药品ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As Recordset
    Dim lng药房ID As Long
    
    If KeyCode = vbKeyDelete Then
        If (mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            '如果是新版LIS组合项目模式，则不允许删除子项
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(vsExt.Row, 3) = "1" Then Exit Sub
            End If
            If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '如果组合项目模式，则同时删除子项
            If mintType = 4 And mblnNewLIS Then
                lngBegin = vsExt.Row + 1
                For j = vsExt.Row + 1 To vsExt.Rows - 1
                    If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
                    lngEnd = j
                Next
                For j = lngEnd To lngBegin Step -1
                    vsExt.RowData(j) = 0
                    For i = 0 To vsExt.Cols - 1
                        vsExt.TextMatrix(j, i) = ""
                        vsExt.Cell(flexcpData, j, i) = ""
                    Next
                    If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                        vsExt.RemoveItem j
                    End If
                Next
            End If
            
            vsExt.RowData(vsExt.Row) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(vsExt.Row, i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '重新初始标本
            If mintType = 4 Then InitCombox
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '定位到下一应输入单元格
        If mintType = 0 Then
            If vsExt.Col <= 1 Then
                vsExt.Col = vsExt.Col + 1
            ElseIf vsExt.Col = 2 And vsExt.Row <= vsExt.Rows - 2 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            ElseIf vsExt.Col = 2 And vsExt.Row = vsExt.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                If mintType = 1 Then
                    vsExt.Col = 0
                Else
                    vsExt.Col = 2
                End If
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Then
            If vsExt.Col = 2 Then
                Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If mintType = 0 Then
            If vsExt.Editable <> flexEDNone Then
                If vsExt.Col = 1 Then
                    If vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1 Then
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 0
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c0").Picture
                    Else
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c1").Picture
                        
                        '自动弹出方法选择器
                        vsExt.Col = 2
                        vsMethod.Tag = "AutoPopup"
                        Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                        vsMethod.Tag = ""
                    End If
                ElseIf vsExt.Col = 2 Then
                    Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                End If
            End If
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    Dim strPrivs As String, i As Long
    Dim strKey As String, lng药名ID As Long
    
    If Not mblnReturn Then
        If mintType = 1 Or mintType = 4 Then
            If Col = 0 And mintType = 1 Or Col = 2 And mintType = 4 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                            
                '重新初始标本
                If mintType = 4 Then InitCombox
                
            End If
        End If
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int性别 As Integer, str药品 As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String
    Dim strSamples As String, strPrivs As String
    Dim strKey As String, lng药名ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        KeyAscii = 0
        
        If mstr性别 Like "*男*" Then
            int性别 = 1
        ElseIf mstr性别 Like "*女*" Then
            int性别 = 2
        End If
        '优化
        strLike = mstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 1 Then
            '输入附加手术:这里不是单独应用,因此不限制
            '"-1*主手术ID"是不排开主手术ID，以作为附加手术加收费用
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='F' And A.ID<>-1*[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                    " And A.服务对象 IN([5],3) And Nvl(A.执行频率,0) IN(0,[6]) And Nvl(A.适用性别,0) IN(0,[7])" & _
                    " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[8])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " Order by A.编码"
            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "手术", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng项目ID, mint简码 + 1, mint服务对象, IIF(mint期效 = 0, 2, 1), int性别, mlng病人科室id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set手术输入(Row, rsTmp)
        ElseIf mintType = 4 Then
            '检验项目
            With Me.cbo标本
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSQL = "Select A.ID,A.编码,A.名称,A.操作类型,A.标本部位,A.试管编码" & _
                " From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID" & _
                " And (A.编码 Like [1] Or C.名称 Like [2] Or C.简码 Like [2]) And C.码类=[3]" & _
                " And A.类别='C' " & _
                IIF(mint场合 = 2, "", " And Nvl(A.单独应用,0)=1 ") & _
                " And Nvl(A.适用性别,0) In (0,[5])" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And A.服务对象 IN([4],3" & IIF(mint场合 = 2, ",4", "") & ") " & _
                " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[6])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
            If strLike = "" Then
                '当可以利用简码索引时(单向匹配),如果有(+)连接,则需要Group By一下(奇怪)
                strSQL = strSQL & " Group by A.ID,A.编码,A.名称,A.操作类型,A.标本部位,A.试管编码"
            End If
            
            strSQL = "Select Distinct A.ID,A.编码,A.名称,A.操作类型 as 检验类型,A.标本部位,A.试管编码" & _
                " From 检验项目参考 D,检验报告项目 E,(" & strSQL & ") A" & _
                " Where A.ID=E.诊疗项目id(+) And E.报告项目ID=D.项目id(+)" & _
                " And (D.标本类型 In (" & strSamples & ") Or D.标本类型 Is Null)" & _
                " Order by A.编码"

            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint简码 + 1, mint服务对象, int性别, mlng病人科室id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
'            If rsTmp!检验类型 = "微生物" And vsExt.Rows > 2 Then
'                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
'                    MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
'                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
'                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
'                    Exit Sub
'                End If
'            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查检验类型、试管编码是否相同
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = NVL(rsTmp!检验类型) _
                        Or vsExt.TextMatrix(i, 1) = "" Or NVL(rsTmp!检验类型) = "") Then
                        MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                        Exit Sub
                    End If
                    If Not (vsExt.Cell(flexcpData, i, 1) = CStr(NVL(rsTmp!试管编码)) _
                        Or vsExt.Cell(flexcpData, i, 1) = "" Or NVL(rsTmp!试管编码) = "") Then
                        MsgBox "请输入相同试管编码的项目，已输入项目的试管编码为""" & vsExt.Cell(flexcpData, i, 1) & """。", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                        Exit Sub
                    End If
                End If
            Next
            
            '重新初始标本
            If Not InitCombox(rsTmp!ID, NVL(rsTmp!标本部位)) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set检验项目(Row, rsTmp)
            If rsTmp!检验类型 = "微生物" Then
                mblnNotAddNew = False
'                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set手术输入(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '附加手术
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = NVL(rsInput!规模)

    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set检验项目(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    
    '检验项目
    '如果新LIS组合项目模式则先删除子项再路径
    '如果组合项目模式，则同时删除子项
    If mblnNewLIS Then
        lngBegin = lngRow + 1
        For j = lngRow + 1 To vsExt.Rows - 1
            If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
            lngEnd = j
        Next
        For j = lngEnd To lngBegin Step -1
            vsExt.RowData(j) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(j, i) = ""
                vsExt.Cell(flexcpData, j, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                vsExt.RemoveItem j
            End If
        Next
    End If
    
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 2) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 2) = vsExt.TextMatrix(lngRow, 2)
    vsExt.TextMatrix(lngRow, 1) = NVL(rsInput!检验类型)
    vsExt.Cell(flexcpData, lngRow, 1) = CStr(NVL(rsInput!试管编码))
    vsExt.TextMatrix(lngRow, 0) = " "
    vsExt.Cell(flexcpBackColor, lngRow, 0) = &H8000000F
    vsExt.TextMatrix(lngRow, 3) = 0 '父项
    
    If mblnNewLIS Then
        strSQL = "" & vbNewLine & _
            "       Select e.Id, e.编码, e.名称, e.操作类型, e.试管编码, a.编码 As 序号, a.Id As 父id" & vbNewLine & _
            "       From 诊疗项目目录 a, 检验报告项目 C, 检验报告项目 D, 诊疗项目目录 E" & vbNewLine & _
            "       Where a.Id = c.诊疗项目id And c.报告项目id = d.报告项目id And d.诊疗项目id = e.Id And e.组合项目 <> 1 And a.Id <> e.Id and a.id=[1]" & vbNewLine & _
            "       Order By 序号, 编码"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsInput!ID))
        Do While Not rsTmp.EOF
            i = vsExt.FindRow(CLng(rsTmp!ID))
            '重复的指标不加入
            If i = -1 Then
                If vsExt.RowData(vsExt.Rows - 1) & "" <> "" Then vsExt.AddItem ""
                vsExt.RowData(vsExt.Rows - 1) = CLng(rsTmp!ID)
                vsExt.Cell(flexcpChecked, vsExt.Rows - 1, 0) = 1
                '子项缩进
                vsExt.TextMatrix(vsExt.Rows - 1, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 2) = vsExt.TextMatrix(vsExt.Rows - 1, 2) '用于恢复显示
                vsExt.TextMatrix(vsExt.Rows - 1, 1) = NVL(rsTmp!操作类型)
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 1) = CStr(NVL(rsTmp!试管编码)) '用于同类输入限制
    '                       If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                vsExt.TextMatrix(vsExt.Rows - 1, 3) = 1  '子项
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExt_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim strTip As String
    
    If mintType = 0 Then
        lngRow = vsExt.MouseRow: lngCol = vsExt.MouseCol
        If Between(lngRow, 0, vsExt.Rows - 1) And Between(lngCol, 0, vsExt.Cols - 1) Then
            If vsExt.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            Else
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 - 240 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            End If
        End If
        vsExt.ToolTipText = strTip
    End If
End Sub

Private Sub vsExt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mintType = 0 Then
        If vsExt.Col = 1 And vsExt.MouseCol = 1 Then
            If x <= vsExt.CellLeft + 250 Then
                Call vsExt_KeyPress(vbKeySpace)
            End If
        End If
    End If
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不允许编辑(该事件后于BeforeEdit,在EditText赋值之前)
    mblnReturn = False
        
    If mintType = 0 Then
        '只允许选择检查方法
        If Col <> 2 Then Cancel = True
    ElseIf mintType = 1 Or mintType = 4 Then
        '只允许编辑附加手术
        If cmd.Visible Then cmd.Visible = False '开始编辑了则隐藏按钮
        If Col <> 0 And mintType = 1 Or Col <> 2 And Col <> 0 And mintType = 4 Then Cancel = True
        '如果开启了新版LIS的组合项目模式则子项不允许输入
        If mblnNewLIS And mintType = 4 And Col = 2 Then
            If vsExt.TextMatrix(Row, 3) = "1" Then Cancel = True
        ElseIf mblnNewLIS And mintType = 4 And Col = 0 Then
            If Val(vsExt.TextMatrix(Row, 3)) = 0 Then Cancel = True
        End If
    End If
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_Click()
    If fraMethod.Visible And vsMethod.Row >= 0 And vsMethod.Col >= 0 Then Call vsMethod_KeyPress(vbKeySpace)
End Sub

Private Sub ConfirmMethod()
'功能：检查方法的确认
    Dim strMethod As String, i As Long
        
    With vsMethod
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strMethod = strMethod & "," & .TextMatrix(i, 1)
            End If
        Next
        vsExt.TextMatrix(vsExt.Row, 2) = Mid(strMethod, 2)
        
        '方法设置后，自动选中该部位
        If vsExt.TextMatrix(vsExt.Row, 2) <> "" Then
            vsExt.Cell(flexcpData, vsExt.Row, 1) = 1
            Set vsExt.Cell(flexcpPicture, vsExt.Row, 1) = img16.ListImages("c1").Picture
        End If
    End With
End Sub
    
Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            Call ConfirmMethod
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '检查方法的选择与取消
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '单选项目前也允许取消选择
                .Cell(flexcpData, .Row, 0) = 0
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                '同时取消该单选项的子项
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '主项没有选择时,子项不能选择
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    If .RowData(.Row) = 1 Then '单选项选中时，取消其他单选项
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 '同时取消该单选项的子项
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            
            Call ConfirmMethod
        End If
    End With
End Sub

Private Function GetCaretPos(ByVal lngHwnd As Long) As PointAPI
'功能：返回编辑控件中当前光标的坐标
'参数：lngHwnd=Edit控件的句柄
'返回：坐标值，基于Edit控件,以Twip为单位
'      如果坐标在控件范围之外，则返回(-1,-1)坐标
    Dim lngPos As Long
    Dim vSel As CHARRANGE
    Dim vPos As PointAPI
    Dim vRect As RECT
    
    SendMessage lngHwnd, EM_EXGETSEL, 0, vSel
    lngPos = SendMessage(lngHwnd, EM_POSFROMCHAR, vSel.cpMin, 0)
    
    vPos.x = lngPos Mod 2 ^ 16
    vPos.Y = lngPos \ 2 ^ 16
    
    '超范围判断
    GetWindowRect lngHwnd, vRect
    If vPos.x >= 0 And vPos.x <= vRect.Right - vRect.Left + 1 _
        And vPos.Y >= 0 And vPos.Y <= vRect.Bottom - vRect.Top + 1 Then
        vPos.x = vPos.x * Screen.TwipsPerPixelX
        vPos.Y = vPos.Y * Screen.TwipsPerPixelY
    Else
        vPos.x = -1: vPos.Y = -1
    End If
    
    GetCaretPos = vPos
End Function

Private Sub SetControlFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：frmMe=需要设置字体的窗体对象
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：1.如果涉及到VsFlexGrid等表格控件，需要根据所在的环境重新调整列宽和行高
'      2.如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control, objrptCol As ReportColumn
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    
    lngFontSize = IIF(bytSize = 0, 9, IIF(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            On Error Resume Next
            If InStr(1, strOther, "," & objCtrol.Container.Name & ",") > 0 Then
                 blnDo = False
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("字") + 20
                        'Label宽度需要自行调整
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("字体" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("字") + IIF(bytSize = 0, 100, 120)
                Case "TextBox", "RichTextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("字")
                        If objCtrol.Name = "txtSentence" Then
                            imgSentence.Width = imgSentence.Width * dblRate
                            imgSentence.Height = imgSentence.Height * dblRate
                            imgSentence.Left = objCtrol.Width + objCtrol.Left
                            picSentence.Width = picSentence.Width * dblRate
                            picSentence.Height = picSentence.Height * dblRate
                        End If
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                        objCtrol.Height = frmMe.TextHeight("字")
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                        
            End Select
        End If
    Next
End Sub
