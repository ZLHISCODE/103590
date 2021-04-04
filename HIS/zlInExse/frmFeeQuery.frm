VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeeQuery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picNum 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8520
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   33
      Top             =   30
      Width           =   1095
      Begin VB.ComboBox cboNum 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   -30
         Width           =   1185
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshInsure 
      Height          =   1095
      Left            =   4980
      TabIndex        =   27
      Top             =   945
      Width           =   3660
      _cx             =   6456
      _cy             =   1931
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   5
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
   Begin VSFlex8Ctl.VSFlexGrid mshDepost 
      Height          =   1260
      Left            =   30
      TabIndex        =   26
      Top             =   915
      Width           =   4545
      _cx             =   8017
      _cy             =   2222
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   5
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
   Begin VB.PictureBox picLR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   5040
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1215
      ScaleWidth      =   45
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox pic费用信息 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   105
      ScaleHeight     =   315
      ScaleWidth      =   7035
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   315
      Width           =   7035
      Begin VB.Label lbl费用信息 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费用信息："
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   135
         TabIndex        =   21
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   19
      Top             =   2130
      Width           =   7275
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   60
      ScaleHeight     =   780
      ScaleWidth      =   10830
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2235
      Width           =   10830
      Begin VB.Frame fraDeptMode 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton optDeptMode 
            Caption         =   "执行科室"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optDeptMode 
            Caption         =   "开单科室"
            Height          =   255
            Index           =   0
            Left            =   100
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame fraTypeMode 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton optTypeMode 
            Caption         =   "收据费目"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton optTypeMode 
            Caption         =   "收入项目"
            Height          =   255
            Index           =   0
            Left            =   100
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新"
         Height          =   315
         Left            =   6420
         TabIndex        =   16
         Top             =   345
         Width           =   630
      End
      Begin VB.ComboBox cbo日期 
         Height          =   300
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   367
         Width           =   1230
      End
      Begin VB.CheckBox chkAdivce 
         Caption         =   "非医嘱费用"
         Height          =   300
         Left            =   6075
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboFeeType 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.TabStrip tabTime 
         Height          =   315
         Left            =   7560
         TabIndex        =   11
         Top             =   -15
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Style           =   2
         TabFixedHeight  =   526
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   882
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "所有费用"
               Key             =   "All"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   2130
         TabIndex        =   14
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271253507
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271253507
         CurrentDate     =   36257.9999884259
      End
      Begin VB.CheckBox chkNotCheckFee 
         Caption         =   "不包含体检费用"
         Height          =   210
         Left            =   9675
         TabIndex        =   32
         Top             =   412
         Width           =   1575
      End
      Begin VB.CheckBox chk仅显示销帐单据 
         Caption         =   "仅显示销帐单据"
         Height          =   210
         Left            =   8850
         TabIndex        =   17
         Top             =   412
         Width           =   1665
      End
      Begin VB.CheckBox chk按规格进行小计 
         Caption         =   "按规格进行小计"
         Height          =   210
         Left            =   7035
         TabIndex        =   31
         Top             =   412
         Width           =   2115
      End
      Begin VB.Label lbl日期范围 
         AutoSize        =   -1  'True
         Caption         =   "2009-01-01 00:00:00至2009-02-02 23:59:59"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2070
         TabIndex        =   30
         Top             =   420
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.Label lbl至 
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3885
         TabIndex        =   29
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lbl发生时间 
         AutoSize        =   -1  'True
         Caption         =   "发生时间"
         Height          =   180
         Left            =   75
         TabIndex        =   12
         Top             =   420
         Width           =   720
      End
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   10410
         Picture         =   "frmFeeQuery.frx":0000
         ToolTipText     =   "选择需要显示的列(ALT+C)"
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 费用查询方式▼"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   -15
         TabIndex        =   18
         ToolTipText     =   "F2:选择费用清单类型"
         Top             =   60
         Width           =   1350
      End
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   315
      Left            =   90
      TabIndex        =   22
      Top             =   0
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   556
      Style           =   2
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "未结费用"
            Key             =   "Main"
            Object.ToolTipText     =   "未结费用清单"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFee 
      Height          =   1260
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   4545
      _cx             =   8017
      _cy             =   2222
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblDepost 
      BackColor       =   &H00808080&
      Caption         =   " 预交款清单"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   690
      Width           =   4050
   End
   Begin VB.Label lblInsure 
      BackColor       =   &H00808080&
      Caption         =   " 保险预结费用"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   2010
   End
End
Attribute VB_Name = "frmFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '自已激活时
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字

Private mint场合 As Integer '0-费用查询，1-护士站调用
Private mlng科室ID As Long, mbln补费 As Boolean '33744

Private mcbsMain As CommandBars
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1

Private mstrUnitIDs As String   '操作员所属的病区集

Private msngScale As Single
Private Const mlngModul = 1139
Private mrsList As ADODB.Recordset '记录当前费用清单
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr住院号 As String
Private mlng病区ID As Long
Private mintInsure As Integer
Private mblnDateMoved As Boolean
Private mbln出院 As Boolean
Private mbln结清 As Boolean
Private mbln门诊留观病人 As Boolean
Private mblnHavePara As Boolean

Private mintPreCard As Integer
Private mintPreTime As Integer
Private mintPreTimeIndex As Integer

Private mbytList As Byte                '未结费用清单和结帐清单查询类型序号
Private mblnClinicOrNurse As Boolean     '当前操作员的缺省部门是否是临床或护理部门

Private mbytDateType As Byte '1-发生时间,2-登记时间
Private mblnPreBalance As Boolean   '是否允许预结医保
Private mblnUnBilling As Boolean '是否可销帐
Private mstrPrivs As String
Private mstr截止日期 As String
Private mblnNotClick As Boolean '不执行相关的选择条件

Private Enum ListType
    C0费用清单 = 0
    C1分科室明细 = 1
    C2分项目明细 = 2
    C3分类别明细 = 3
    C4分类分项明细 = 4
    
    C5分项目汇总 = 5
    C6分类别汇总 = 6
    C7分月分类汇总 = 7
    C8逐日单据汇总 = 8
    C9逐日费目汇总 = 9
End Enum

Private Const conTab未结 = 1

Private Type t_ViewState
    ReBalance As Boolean
    ZeroFee As Boolean
    CheckFee As Boolean
End Type
Private mobjInPati As Object
Private mvs As t_ViewState
Private mbytFontSize As Byte
Private mblnFisrtSetFontSize As Boolean '第一次设置字体大小
Private mcllBalaceNums As Collection
Private mstrRestoreFeeCons As String
Private mblnContainOutFee As Boolean '是否包含门诊费用

Public Sub SetFontSize(ByVal bytSize As Byte)
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub
Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置字体大小
    '编制:刘兴洪
    '日期:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("刘") + 20
        Case UCase("VsFlexGrid")
            Call zlControl.VSFSetFontSize(objCtrl, mbytFontSize)
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("刘") * 1.5
            
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    Call Form_Resize
    Call picDetail_Resize
    '问题:55392
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False, , mblnHavePara
 End Sub

Private Sub InitBaseData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化基础数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '问题:24913
    '日期:2009-08-17 09:47:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType  As String, strBeginDate As String, strEndDate As String, varData As Variant, intTYPE As Integer, i As Long
    With cbo日期
        .AddItem "所有"
        .ItemData(.NewIndex) = 0
        .AddItem "今日"
        .ItemData(.NewIndex) = 1
        .AddItem "昨日"
        .ItemData(.NewIndex) = 2
        .AddItem "本周"
        .ItemData(.NewIndex) = 3
        .AddItem "本月"
        .ItemData(.NewIndex) = 4
        .AddItem "自定义范围"
        .ItemData(.NewIndex) = 5
    End With
    strType = zlDatabase.GetPara("费用查询范围", glngSys, mlngModul, "所有", Array(cbo日期), InStr(1, mstrPrivs, ";参数设置;") > 0)
    mblnHavePara = InStr(1, mstrPrivs, ";参数设置;") > 0
       
    varData = Split(strType & "|", "|"): strType = varData(0)
    intTYPE = Switch(strType = "所有", 0, strType = "今日", 1, strType = "昨日", 2, strType = "本周", 3, strType = "本月", 4, True, 5)
    If intTYPE = 5 Then
       varData = Split(varData(1) & ",", ",")
       If varData(0) <> "" And IsDate(varData(0)) Then dtpBegin.Value = Format(CDate(varData(0)), "yyyy-mm-dd 00:00:00")
       If varData(1) <> "" And IsDate(varData(1)) Then dtpEnd.Value = CDate(Format(CDate(varData(1)), "yyyy-mm-dd") & " 23:59:59")
    End If
    For i = 0 To cbo日期.ListCount - 1
        If cbo日期.ItemData(i) = intTYPE Then
            cbo日期.ListIndex = i: Exit For
        End If
    Next
    If cbo日期.ListIndex < 0 Then cbo日期.ListIndex = 0
    '46646
    chkNotCheckFee.Value = IIf(Val(zlDatabase.GetPara("不含体检费用", glngSys, mlngModul, "0", Array(chkNotCheckFee), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
    
    
    strType = zlDatabase.GetPara("明细仅显销帐单据", glngSys, mlngModul, "0", Array(chk仅显示销帐单据), InStr(1, mstrPrivs, ";参数设置;") > 0)
    chk仅显示销帐单据.Value = IIf(Val(strType) = 0, 0, 1)
   
    chk仅显示销帐单据.Visible = False
    strType = zlDatabase.GetPara("按规格分类统计", glngSys, mlngModul, "0", Array(chk按规格进行小计), InStr(1, mstrPrivs, ";参数设置;") > 0)
    '问题:41673
    chk按规格进行小计.Value = IIf(Val(strType) = 0, 0, 1)
    chk按规格进行小计.Visible = mbytList = ListType.C2分项目明细
    
    Select Case mbytList
        Case ListType.C0费用清单, ListType.C1分科室明细, ListType.C2分项目明细, ListType.C3分类别明细, ListType.C4分类分项明细 '明细清单,分科明细,项目明细,分类明细,(按收入项目(或收据费目),收费项目,明细分级查询)
            chk仅显示销帐单据.Visible = True
            
        Case ListType.C5分项目汇总  '项目汇总
        Case ListType.C6分类别汇总  '分类汇总
        Case ListType.C7分月分类汇总  '分月汇总
        Case ListType.C8逐日单据汇总  '逐日费用
        Case ListType.C9逐日费目汇总  '逐日费目
    End Select
End Sub
Private Sub SetDateVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置日期控件的visible特性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '问题:24913
    '日期:2009-08-17 10:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String, intTYPE As Integer, blnDateVisible As Boolean, strBeginDate As String, strEndDate As String
    
    strType = cbo日期.Text
    intTYPE = Switch(strType = "所有", 0, strType = "今日", 1, strType = "昨日", 2, strType = "本周", 3, strType = "本月", 4, True, 5)
    blnDateVisible = intTYPE = 5
    dtpBegin.Visible = blnDateVisible: dtpEnd.Visible = blnDateVisible: lbl至.Visible = blnDateVisible
    cmdRefresh.Visible = blnDateVisible
    
    lbl日期范围.Visible = (Not blnDateVisible) And intTYPE <> 0
    
    If lbl日期范围.Visible Then
        zlGetDateRange , strBeginDate, strEndDate
        lbl日期范围.Caption = strBeginDate & "至" & strEndDate
    End If
End Sub

Private Sub zlGetDateRange(Optional ByVal blnOnlyDate As Boolean = False, Optional ByRef strBeginDate As String, Optional ByRef strEndDate As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取日期范围
    '入参:blnOnlyDate-true:仅为日期(2009-01-01),否则为带分秒的日期(20009-01-01 23:59:59)
    '出参:strBeginDate-开始日期,strEndDate-结束日期
    '返回:
    '编制:刘兴洪
    '日期:2009-08-17 10:23:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    Dim blnTime As Boolean
    strType = cbo日期.Text
    blnTime = False
    Select Case strType
    Case "所有"
        strBeginDate = "": strEndDate = ""
    Case "今日"
         strBeginDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
         strEndDate = strBeginDate
    Case "昨日"
         strBeginDate = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-mm-dd"): strEndDate = strBeginDate
    Case "本周"
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd"): strBeginDate = Format(DateAdd("d", -1 * (Weekday(CDate(strEndDate), vbSunday) - 1), CDate(strEndDate)), "yyyy-mm-dd")
    Case "本月"
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd"): strBeginDate = Format(CDate(strEndDate), "yyyy-mm") & "-01"
    Case Else
        strBeginDate = Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"): strEndDate = Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
        blnTime = True
    End Select
    If blnOnlyDate = False And strBeginDate <> "" And Not blnTime Then
          strBeginDate = strBeginDate & " 00:00:00": strEndDate = strEndDate & " 23:59:59"
    End If
End Sub

Private Sub RefreshAllData()
'功能：刷新数据
    Dim blnMCPatient As Boolean, i As Integer
    
    '根据是否医保病人显示保险模拟结算费用
    If tabClass.SelectedItem.Index = conTab未结 Then blnMCPatient = mintInsure <> 0
    
            
    '读取保险模拟结算费用清单
    '刘兴洪:加入        mblnPreBalance = True:原因是control.enabled可能不为true
    '25657:
    If blnMCPatient Then
        Call ReadInsureMoney(mlng病人ID, mlng主页ID)
        mblnPreBalance = True
    Else
        mshInsure.Clear
        mshInsure.Rows = 2
        mblnPreBalance = True
    End If
    lblInsure.Visible = blnMCPatient
    mshInsure.Visible = blnMCPatient
    picLR.Visible = blnMCPatient
    Call Form_Resize
    
    Call LoadPatientBaby(cboBaby, mlng病人ID, mlng主页ID)
    cboBaby.Visible = cboBaby.ListCount > 1
    If cboBaby.ListCount > 1 Then
        cboBaby.AddItem "病人和婴儿"
        cboBaby.ItemData(cboBaby.NewIndex) = 999
        Call zlControl.CboSetIndex(cboBaby.hWnd, cboBaby.NewIndex)
    End If
    zlControl.CboSetWidth cboBaby.hWnd, cboBaby.Width * 2
    
    
    If LoadPatiClass Then '初始化选项卡
        Call ChangeList(False)  '启动时
        If mstrRestoreFeeCons <> "" Then
            For i = 1 To tabClass.Tabs.Count
                If tabClass.Tabs(i).Key = Nvl(Split(mstrRestoreFeeCons, "|")(1)) Then tabClass.Tabs(i).Selected = True
            Next
        End If
        If i = 0 Then tabClass.Tabs(1).Selected = True '调用tabClass_Click显示费用
    End If
          
    Call SetCondition '主要是因为婴儿费可能变化
    If mstrRestoreFeeCons <> "" Then
        If zlRestorePosition(mlng病人ID) = False Then Exit Sub
    End If
End Sub


Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str住院号 As String, ByVal lng病区ID As Long, _
    ByVal intInsure As Integer, ByVal blnDateMoved As Boolean, ByVal bln出院 As Boolean, ByVal bln结清 As Boolean, _
    blnOnlyRefreshVar As Boolean, _
    Optional bln补费 As Boolean = False, Optional lng科室ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:bln出院-是否出院病人
    '       bln补费-对转科，转病区的病人进行补费
    '       lng科室ID-当补费为true时,则传入本次需要补费的科室ID
    '       lng病区ID-当补费为true时,则传入本次需要补费的病区ID
    '       blnOnlyRefreshVar-仅刷新内部变量
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-10 14:43:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr住院号 = str住院号
    mlng病区ID = lng病区ID
    mlng科室ID = lng科室ID
    mbln补费 = bln补费
    mintPreTimeIndex = 0
    mintPreTime = 0
    mintInsure = intInsure
    mblnDateMoved = blnDateMoved
    mbln出院 = bln出院
    mbln结清 = bln结清
    mbln门诊留观病人 = ZlIsOutpatientObserve(lng病人ID, lng主页ID)
    If blnOnlyRefreshVar Then Exit Sub
    tabClass.Tag = ""
    Call RefreshAllData
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strBillingPrivs As String
    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub
    blnVisible = True
    
    strBillingPrivs = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    Select Case Control.ID
        Case conMenu_File_PrintPageSet
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p费用查询), ";病人帐页") > 0
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintSingleBill
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p费用查询), ";催款单打印") > 0
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p费用查询), ";预结所有病人") > 0
        Case conMenu_Edit_Billing, conMenu_Edit_Copy, conMenu_Edit_Billing_Mulit
            '54274
            blnVisible = InStr(strBillingPrivs, "住院记帐") > 0
        Case conMenu_Edit_CardBackMoney
            blnVisible = InStr(";" & GetInsidePrivs(9000), ";在院病人余额退款;") > 0 Or InStr(";" & GetInsidePrivs(9000), ";出院病人余额退款;") > 0
        Case conMenu_Edit_ReBilling
            '55380
            blnVisible = InStr(strBillingPrivs, ";药品销帐;") > 0 _
                Or InStr(strBillingPrivs, ";诊疗销帐;") > 0 _
                Or InStr(strBillingPrivs, ";卫材销帐;") > 0
        Case conMenu_Edit_ReBillingApply
            '55380
            blnVisible = (InStr(strBillingPrivs, ";药品销帐申请;") > 0 _
                Or InStr(strBillingPrivs, ";诊疗销帐申请;") > 0 _
                Or InStr(strBillingPrivs, ";卫材销帐申请;") > 0) _
                And InStr(strBillingPrivs, "部分销帐") > 0
                
        Case conMenu_Edit_ReBillingAudit
            blnVisible = InStr(strBillingPrivs, "销帐审核") > 0
        Case conMenu_Edit_ReBillingButton
            '55380
            blnVisible = InStr(strBillingPrivs, "销帐审核") > 0 _
                Or ((InStr(strBillingPrivs, ";药品销帐申请;") > 0 _
                        Or InStr(strBillingPrivs, ";诊疗销帐申请;") > 0 _
                        Or InStr(strBillingPrivs, ";卫材销帐申请;") > 0) And InStr(strBillingPrivs, "部分销帐") > 0)
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Private Function GetPatiInsure() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select A.登记时间, B.险类, E.密码, Nvl(E.医保号, D.信息值) 医保号" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病案主页从表 D, 医保病人档案 E, 医保病人关联表 F" & vbNewLine & _
            "Where B.病人id = [1] And B.主页id = [2] And A.病人id = B.病人id And B.病人id = D.病人id(+) And B.主页id = D.主页id(+) And D.信息名(+) = '医保号' And" & vbNewLine & _
            "      A.病人id = F.病人id(+) And F.标志(+) = 1 And F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+)"
    On Error GoTo errH
    Set GetPatiInsure = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ExecPreBalance()
    Dim int险类 As Integer
    Dim str医保号 As String, str密码 As String
    Dim rsTmp As ADODB.Recordset, str结算费用 As String
    Dim blnDateMoved As Boolean, dat登记时间 As Date
    
    Set rsTmp = GetPatiInsure
    If rsTmp.RecordCount > 0 Then
    With rsTmp
        int险类 = Val(!险类)
        str医保号 = "" & !医保号
        str密码 = "" & !密码
        dat登记时间 = !登记时间
    End With
    End If
    If int险类 = 0 Then
        MsgBox "读取病人医保相关信息失败!", vbExclamation, gstrSysName
        Exit Sub
    End If
    If gclsInsure.GetCapability(support结帐_结帐设置后调用接口, mlng病人ID, mintInsure) Then
        MsgBox "该医保接口不支持结帐设置前预结算!", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    blnDateMoved = zlDatabase.DateMoved(dat登记时间, , , Caption)
    
    Screen.MousePointer = 11
    Set rsTmp = GetVBalance(1, "住院费用结帐", int险类, mlng病人ID, , , , , blnDateMoved)
    Screen.MousePointer = 0
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人没有未结帐的保险项目费用!", vbInformation, gstrSysName
    Else
        str结算费用 = gclsInsure.WipeoffMoney(rsTmp, mlng病人ID, str医保号, "0", int险类, "|0") '当成中途结算
        MsgBox "预结算成功!" & str结算费用, vbInformation, gstrSysName '可报销金额串:"报销方式;金额;是否允许修改|...."
        Call RefreshAllData
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim i As Long, objControl As CommandBarControl
    Select Case Control.ID
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Preview
            PrintList 2
        Case conMenu_File_Print
            PrintList 1
        Case conMenu_File_Excel
            PrintList 3
        Case conMenu_File_PrintBedCard
            Call zlPrintBedCard(Me, mlng病人ID, mlng主页ID)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_PrintSingleBill
            Call zlExecPrintSingleBill(Me, mlng病人ID, mstr截止日期)
        Case conMenu_File_PrintDayDetail
            Call zlPrintDayDetail(Me, mint场合, mlng病人ID, mlng病区ID, mvs.ReBalance, mvs.ZeroFee, mbytDateType = 1, mlng主页ID)
        Case conMenu_File_PrintPageSet  '打印帐页设置
            Call zlPrintAccountPage(Me)
        Case conMenu_Edit_PreBalance     '预结算
            If zlPreBalance(Me, mlng病人ID, mlng主页ID) = True Then RefreshAllData
        Case conMenu_Edit_PreBalanceAll      '预结算所有
            Call zlPreBalanceAll(Me, mlng病区ID)
        Case conMenu_Edit_PatiMemo '修改病人备注信息
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, mlng病人ID, mlng主页ID, mobjInPati, False)
        Case conMenu_Edit_Billing   '记帐
            '问题:33744
            If zlExecBilling(IIf(mint场合 = 0, 6, mint场合), gfrmMain, mlng病区ID, mlng病人ID, mbln出院, mbln结清, _
                mstrUnitIDs, mlng主页ID, mbln补费, mlng科室ID, , mbln门诊留观病人) Then Call RefreshAllData
        Case conMenu_Edit_Billing_Mulit '批量记帐
            If zlExecBilling_Mulit(IIf(mint场合 = 0, 6, mint场合), gfrmMain, mlng病区ID, mlng病人ID, mbln出院, mbln结清, _
                mstrUnitIDs, mlng主页ID, mbln补费, mlng科室ID) Then Call RefreshAllData
        Case conMenu_Edit_Copy '复制记账单
            '54274
            If zlCopyBill(IIf(mint场合 = 0, 6, mint场合), gfrmMain, mlng病区ID, mlng病人ID, mbln出院, mbln结清, _
                mstrUnitIDs, mlng主页ID, mlng科室ID, mbln门诊留观病人) Then Call RefreshAllData
        Case conMenu_Edit_ReBilling '销帐
            Call ExecUnBilling
        Case conMenu_Edit_CardBackMoney '余额退款
            Call NurseDeposit(mfrmParent, mlng病人ID, mlng主页ID, True, IIf(mbln门诊留观病人, 1, 2))
        Case conMenu_Edit_ReBillingApply
            If vsfFee.ColIndex("单据号") = -1 Then
                If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng病区ID, mlng病人ID, Control.ID = conMenu_Edit_ReBillingApply) = True Then
                    RefreshAllData
                End If
            Else
                If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng病区ID, mlng病人ID, Control.ID = conMenu_Edit_ReBillingApply, vsfFee.TextMatrix(vsfFee.Row, vsfFee.ColIndex("单据号"))) = True Then
                    RefreshAllData
                End If
            End If
        Case conMenu_Edit_ReBillingAudit
            If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng病区ID, mlng病人ID, Control.ID = conMenu_Edit_ReBillingApply) = True Then RefreshAllData
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2 '时间模式
            mbytDateType = Control.ID - conMenu_View_DateType * 10
            lbl发生时间.Caption = IIf(mbytDateType = 1, "发生时间", "登记时间")
            Call LoadCardData(False, False, True)
        
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9 '查询方式'
            
            '保存上次选择的结果
            zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
            
            mbytList = Control.ID - conMenu_View_DetailType * 10
            chk仅显示销帐单据.Visible = False
            chk按规格进行小计.Visible = mbytList = ListType.C2分项目明细
            Select Case mbytList
                Case ListType.C0费用清单, ListType.C1分科室明细, ListType.C2分项目明细, ListType.C3分类别明细, ListType.C4分类分项明细 '明细清单,分科明细,项目明细,分类明细,(按收入项目(或收据费目),收费项目,明细分级查询)
                    chk仅显示销帐单据.Visible = True
                Case ListType.C5分项目汇总  '项目汇总
                Case ListType.C6分类别汇总  '分类汇总
                Case ListType.C7分月分类汇总  '分月汇总
                Case ListType.C8逐日单据汇总  '逐日费用
                Case ListType.C9逐日费目汇总  '逐日费目
            End Select
            Call ChangeList(True)
            Call picDetail_Resize
            
        Case conMenu_View_ReBalance '显示结帐作废
            Control.Checked = Not Control.Checked: mvs.ReBalance = Control.Checked
            Call RefreshAllData
        Case conMenu_View_ZeroFee   '显示零费用
            Control.Checked = Not Control.Checked: mvs.ZeroFee = Control.Checked
            Call LoadCardData(False, False, True)
        Case conMenu_View_CheckFee  '显示体检费用
            Control.Checked = Not Control.Checked: mvs.CheckFee = Control.Checked
            Call LoadCardData(True, True, True)
        Case conMenu_View_TurnToWardFeeQuery '转病区费用变动查询
            If CreatePublicExpenseBillOperation() Then
                Call gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 3, mlng病人ID, mlng主页ID)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            For i = 1 To cbsTools.Count
                cbsTools(i).Visible = Not cbsTools(i).Visible
            Next
            cbsTools.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            Control.Checked = Not Control.Checked
            For i = 1 To cbsTools.Count
                For Each objControl In cbsTools(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsTools.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            cbsTools.Options.LargeIcons = Not cbsTools.Options.LargeIcons
            cbsTools.RecalcLayout
        Case conMenu_View_PatInfor  '查看病人卡片
            Call ShowPatiCard
        Case conMenu_View_Billing   '查看记帐单
            Call vsfFee_DblClick
        Case conMenu_View_Refresh
            mintPreCard = 0: mintPreTime = 0
            Call RefreshAllData
        Case conMenu_Tool_Option    '记帐选项
            frmSetExpence.mlngModul = 1133
            frmSetExpence.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
            frmSetExpence.mbytInFun = 0
            frmSetExpence.mbytUseType = 1   '当成住院记帐管理模块调用
            frmSetExpence.Show 1, Me
        Case conMenu_View_ContainOutFee
            Control.Checked = Not Control.Checked
            mblnContainOutFee = IIf(Control.Checked, 1, 0)
            Call LoadCardData(True, True, True)
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                '执行发布到当前模块的报表
                If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then '催款表(即使没有显示病人也可以使用)
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                ElseIf Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then '住院科室日报(护士站调用才有)
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                             "病区=" & mlng病区ID, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID)
                Else
                    If mlng病人ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "病区=" & mlng病区ID)
                    Else
                        Dim lng结帐ID As Long
                        lng结帐ID = Val(IIf(tabClass.SelectedItem.Index = 1, 0, tabClass.SelectedItem.Tag))
                        If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then  '病人帐页
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "住院号=" & mstr住院号, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, _
                             "病区=" & mlng病区ID, "结帐ID=" & lng结帐ID)
                        Else
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "住院号=" & mstr住院号, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, _
                             "病区=" & mlng病区ID, "结帐ID=" & lng结帐ID)
                        End If
                    End If
                End If
          End If
    End Select
End Sub

Private Sub ShowPatiCard()
    frmDegreeCard.mlng病人ID = mlng病人ID
    frmDegreeCard.mlng主页ID = mlng主页ID
    frmDegreeCard.Show 1, Me
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByRef cbsMain As CommandBars, _
    ByVal int场合 As Integer, Optional ByVal blnChildToolBar As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
    '入参:int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    '       CommandBars=仅用于查看时可以不传(传入Nothing)
    '       blnChildToolBar = True表示工具栏添加在自己的窗体内部
    '出参:
    '返回:
    '说明:
    '   定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)，如果bln内部工具栏为假，则不再主界面上创建工具栏（菜单仍然要创建），
    '   而需要在自己的界面上创建工具栏，因此对于自己界面上已经存在工具栏的程序，应避免关键字重复。
    '注意:
    '         添加工具栏时注意各个功能按钮的主键不要重复
    '         病人事务处理模块无菜单：conMenu_ManagePopup，因此子程序在处理时需要检查，无此对象时添加到自已的菜单中
    '         如果是添加内部工具栏，先删除活动工具栏后再添加
    '         未使用工具栏的模块需要添加初始化工具栏代码
    '         工具条的功能状态的变化，通过主程序调用zlUpdateCommandBars来统一更新
    '编制:刘兴洪
    '日期:2010-10-29 15:14:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Set mfrmParent = frmParent
    Set mcbsMain = cbsMain
    mint场合 = int场合
        
    Err = 0: On Error GoTo ErrHand:
        
    '文件菜单
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_FilePopup, True, False)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(xtpControlButton, conMenu_File_Excel, True, False) '输出到Excel之后
        If mint场合 = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "打印床头卡(&K)…", objControl.Index + 1) '打印床头卡
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintPageSet, "打印帐页设置(&A)…", objControl.Index + 1)
    End With
    
    '有编辑菜单时，放在编辑菜单下(费用查询模块)，否则放在管理菜单(主窗体可能没有)或文件菜单后面
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_EditPopup, True, False)
    If objMenu Is Nothing Then
        Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ManagePopup, True, False)
        If objMenu Is Nothing Then
            Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_FilePopup, True, False)
        End If
        ''0-费用查询，1-护士站调用
        '快键:C;E:63630
        Set objMenu = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, IIf(mint场合 = 1, "费用(&C)", "编辑(&E)"), objMenu.Index + 1, False)
        objMenu.ID = conMenu_EditPopup
    End If
    With objMenu.CommandBar.Controls
        If mint场合 = 1 Then
            '问题:40900
            Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalanceAll, "预结所有病人(&I)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结当前病人(&W)")
            objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "记帐(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing_Mulit, "批量记帐(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBilling, "销帐(&D)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "销帐申请(&L)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingAudit, "销帐审核(&U)", objControl.Index + 1)
        If mint场合 = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_View_Billing, "查看记帐单(&D)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "余额退款(&F)"): objControl.BeginGroup = True
        End If
        '54274
       Set objControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制记账单(&F)"): objControl.BeginGroup = True
       If mint场合 = 0 Then .Add(xtpControlButton, conMenu_Edit_PatiMemo, "病人备注信息(&M)").BeginGroup = True
    End With
             
    
    '查看菜单
    '-----------------------------------------------------
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(xtpControlButton, conMenu_View_StatusBar, True, False) '状态栏项后
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_DetailType, "清单类型(&M)", objControl.Index + 1): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 0, "费用清单(&0)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 1, "分科室明细(&1)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 2, "分项目明细(&2)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 3, "分类别明细(&3)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 4, "分类分项明细(&4)", -1, False
                        
            .Add(xtpControlButton, conMenu_View_DetailType * 10 + 5, "分项目汇总(&5)", -1, False).BeginGroup = True
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 6, "分类别汇总(&6)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 7, "分月分类汇总(&7)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 8, "逐日单据汇总(&8)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 9, "逐日费目汇总(&9)", -1, False
        End With
                
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_DateType, "查询时间(&E)", objPopup.Index + 1)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_DateType * 10 + 1, "发生时间(&H)", -1, False
            .Add xtpControlButton, conMenu_View_DateType * 10 + 2, "登记时间(&A)", -1, False
        End With
        
        If mint场合 = 1 Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
            With objMenu.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "查看病人详细信息(&K)"): objControl.BeginGroup = True
            End With
         End If
         
        Set objControl = .Add(xtpControlButton, conMenu_View_ReBalance, "显示结帐作废(&Q)", objPopup.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_ZeroFee, "显示零费用(&Z)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_CheckFee, "显示体检费用(&C)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_TurnToWardFeeQuery, "转病区费用变动查询(&T)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_ContainOutFee, "包含门诊费用(&B)", objControl.Index + 1)
    End With
    
    
    '报表菜单
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ReportPopup, True, False)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ReportPopup, True, False)
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "打印一日清单(&D)…", 1)
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "打印催款单(&C)…", objControl.Index + 1)
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "记帐操作选项(&O)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Parameter
    End With
    
    
    If mint场合 <> 1 Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_EditPopup, True, False)
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "查看病人详细信息(&K)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Billing, "查看记帐单(&D)", objControl.Index + 1)
        End With
    End If
        
    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    If blnChildToolBar Then
        Set objBar = CreateChildTools
    Else
        cbsTools.DeleteAll
        Set objBar = mcbsMain(2)
    End If
    
    If blnChildToolBar = False Then
        For Each objControl In objBar.Controls '先求出前面的最后一个Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
            End If
        Next
    End If
    Dim intIndex As Integer
    With objBar.Controls
        Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
        If objControl Is Nothing Then
            intIndex = 0
        Else
            intIndex = objControl.Index
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "催款", intIndex + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "一日", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结", objControl.Index + 1)
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "记帐", objControl.Index + 1)
        objControl.BeginGroup = True
        intIndex = objControl.Index
        For Each objControl In objBar.Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
        
        Set objPopup = .Add(xtpControlPopup, conMenu_Edit_ReBillingButton, "销帐", intIndex + 1)
        objPopup.ID = conMenu_Edit_ReBillingButton
        objPopup.IconId = conMenu_Edit_ReBillingButton
        objPopup.Style = xtpButtonIconAndCaption
                
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add 0, VK_F4, conMenu_Edit_Billing
        .Add 0, VK_F6, conMenu_Edit_ReBilling
        .Add 0, vbKeyF11, conMenu_Tool_Option '记帐选项
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options   '如果隐藏了，控件在菜单第一次显示时没有调用update事件
'        .AddHiddenCommand conMenu_View_ReBalance
'        .AddHiddenCommand conMenu_View_ZeroFee
'        .AddHiddenCommand conMenu_View_CheckFee
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function CreateChildTools() As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建子工具栏
    '编制:刘兴洪
    '日期:2010-10-29 15:59:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsTools.DeleteAll
    Set CreateChildTools = cbsTools.Add("费用操作", xtpBarTop)
    CreateChildTools.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    CreateChildTools.ModifyStyle XTP_CBRS_GRIPPER, 0
    
End Function


Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_ReBillingButton '费用销帐
        With CommandBar.Controls
            .DeleteAll
            .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "销帐申请(&L)").BeginGroup = True
            .Add xtpControlButton, conMenu_Edit_ReBillingAudit, "销帐审核(&U)"
        End With
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnSelect As Boolean, lngColTmp As Long, blnEnabled As Boolean
    
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    blnSelect = mlng病人ID <> 0
    Select Case Control.ID
        '文件
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel
            Control.Enabled = vsfFee.Rows > vsfFee.FixedRows
        Case conMenu_File_PrintSingleBill
            Control.Enabled = blnSelect
            
        '编辑
        Case conMenu_Edit_PreBalance
            Control.Enabled = blnSelect
            If blnSelect Then
                Dim blnMCPatient As Boolean
                If tabClass.SelectedItem.Index = conTab未结 Then
                    blnMCPatient = mintInsure <> 0
                End If
                Control.Enabled = blnMCPatient
            End If
            mblnPreBalance = Control.Enabled
        Case conMenu_Edit_PatiMemo   '修改备注信息
           ' Control.Visible = InStr(1, mstrPrivs, ";病人备注编辑;")
            Control.Enabled = mlng病人ID > 0 And Control.Visible
        Case conMenu_Edit_Billing
            Control.Enabled = blnSelect
        Case conMenu_Edit_Billing_Mulit '批量记帐
        
        Case conMenu_Edit_Copy
            '54274
            Control.Enabled = blnSelect
            With vsfFee
                If Control.Enabled Then
                    If mbytList = ListType.C0费用清单 Or mbytList = ListType.C1分科室明细 Or mbytList = ListType.C2分项目明细 Or mbytList = ListType.C3分类别明细 Or mbytList = ListType.C4分类分项明细 Then
                        '.row>=1:61895
                        Control.Enabled = .ColIndex("单据号") >= 0 And .ColIndex("记录状态") >= 0 And .ColIndex("记录性质") >= 0 And .Row >= 1
                        If Control.Enabled Then
                            If Trim(.TextMatrix(.Row, .ColIndex("单据号"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 2 Or Val(.TextMatrix(.Row, .ColIndex("记录性质"))) = 3 Then
                                Control.Enabled = False
                            End If
                        End If
                     Else
                        Control.Enabled = False
                    End If
                End If
            End With
        Case conMenu_Edit_ReBilling '销帐
            Control.Enabled = blnSelect
            With vsfFee
                If Control.Enabled Then
                    If mbytList = ListType.C0费用清单 Or mbytList = ListType.C1分科室明细 Or mbytList = ListType.C2分项目明细 Or mbytList = ListType.C3分类别明细 Or mbytList = ListType.C4分类分项明细 Then
                        lngColTmp = VsfGetColNum(vsfFee, "记录状态")
                        If lngColTmp = -1 Or .Row < 1 Then
                            Control.Enabled = False
                        Else
                            lngColTmp = Val(.TextMatrix(.Row, lngColTmp))
                            Control.Enabled = (lngColTmp = 1 Or lngColTmp = 3)
                        End If
                    Else
                        Control.Enabled = False
                    End If
                End If
            End With
            mblnUnBilling = Control.Enabled
            
       '查看
        Case conMenu_View_PatInfor
            Control.Enabled = blnSelect
        Case conMenu_View_Billing
            Control.Enabled = vsfFee.Rows > vsfFee.FixedRows
            
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2
            Control.Checked = (Control.ID - conMenu_View_DateType * 10) = mbytDateType
            Control.Enabled = blnSelect
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9
            Control.Checked = (Control.ID - conMenu_View_DetailType * 10) = mbytList
            Control.Enabled = blnSelect
            
        Case conMenu_View_ZeroFee
            Control.Enabled = tabClass.SelectedItem.Index = conTab未结
            Control.Checked = mvs.ZeroFee
        Case conMenu_View_CheckFee
            Control.Checked = mvs.CheckFee
        Case conMenu_View_ReBalance
            Control.Checked = mvs.ReBalance
        Case conMenu_View_ContainOutFee
            Control.Checked = mblnContainOutFee
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then Control.Enabled = blnSelect  '病人帐页
            End If
        
    End Select
End Sub

Private Sub cboNum_Click()
    If mblnNotClick Then Exit Sub
    Call LoadPages(cboNum.ListIndex + 1)
End Sub

Private Sub cbo日期_Click()
    If mblnNotClick Then Exit Sub
    Call SetDateVisible
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub cbo日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
        Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
   If CommandBar.Parent Is Nothing Then Exit Sub
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsTools_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
       Call Form_Resize
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
       Call zlUpdateCommandBars(Control)
End Sub

Private Sub chkAdivce_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub


Private Sub cboBaby_Click()
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub

Private Sub cboFeeType_Click()
    If mblnNotClick Then Exit Sub
    Call FilterDetail
End Sub

Private Sub cboDept_Click()
    '问题64817:刘尔旋,2013-10-28,切换到汇总列表时查询的汇总内容并没有按照科室条件进行过滤
    If cboDept.Tag = "不刷新" Or mblnNotClick Then Exit Sub
    Call FilterDetail
End Sub

Private Sub FilterDetail()
'功能:根据选择的费目或科室过滤明细费用
'参数:
    Dim arrTotal(2) As Currency
    Dim strFilter As String
        
    Select Case mbytList
        Case ListType.C0费用清单, ListType.C1分科室明细, ListType.C2分项目明细, ListType.C3分类别明细, ListType.C4分类分项明细
            If mrsList Is Nothing Then Exit Sub
            If mrsList.State = adStateClosed Then Exit Sub
            
            If cboFeeType.ListIndex > 0 Then strFilter = "费目='" & cboFeeType.Text & "'"
            If cboDept.ListIndex > 0 Then strFilter = IIf(strFilter = "", "", strFilter & " And") & " 开单科室='" & cboDept.Text & "'"
            
            mrsList.Filter = strFilter
            Set vsfFee.DataSource = mrsList
            
            Call SetVsffeeFormat
            vsfFee.AutoSize 0, vsfFee.Cols - 1
            
            '保存个性化设置
            zl_vsGrid_Para_Restore mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
            
        Case ListType.C5分项目汇总, ListType.C6分类别汇总, ListType.C7分月分类汇总, ListType.C8逐日单据汇总, ListType.C9逐日费目汇总
            Call LoadCardData(False, False, True, False)
    End Select
End Sub

Private Function LoadPatiClass() As Boolean
'功能：设置病人的费用选项卡
    Dim strSQL As String, i As Long, intPage As Integer, intCount As Integer
    Dim rsTmp As ADODB.Recordset
    Dim cllPage As Collection
    Dim str(3) As String
    
    mintPreCard = 0
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
        
    '如果当前病人的入院时间在转出时间之前,则需要联结后备数据表查询
    If mblnDateMoved Then
        strSQL = zlGetFullFieldsTable("病人结帐记录")
    Else
        strSQL = "病人结帐记录 A"
    End If
    
    strSQL = "Select A.ID,A.NO,A.收费时间 as 日期,A.记录状态" & _
        " From " & strSQL & " " & _
        " Where A.病人ID = [1]" & _
        " And A.记录状态 IN (1" & IIf(mvs.ReBalance, ",3", "") & ")" & _
        " Order by A.ID Desc"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID)
        
    cboNum.Visible = False
    picNum.Visible = False
    Set mcllBalaceNums = New Collection
    If rsTmp.RecordCount >= 20 Then
        cboNum.Visible = True
        picNum.Visible = True
    End If
    
    Set cllPage = New Collection
    Do While Not rsTmp.EOF
        If rsTmp.RecordCount < 20 Then
            tabClass.Tabs.Add , "_" & rsTmp!NO, Format(rsTmp!日期, "yyyy-MM-dd") & IIf(rsTmp!记录状态 = 1, " 结帐", " 退费")
            tabClass.Tabs(tabClass.Tabs.Count).Tag = rsTmp!ID '记录结帐ID,加快速度
            tabClass.Tabs(tabClass.Tabs.Count).ToolTipText = "结帐时间:" & Format(rsTmp!日期, "yyyy-MM-dd hh:mm:ss")
        Else
            str(0) = Val(rsTmp!ID)
            str(1) = Nvl(rsTmp!NO)
            str(2) = Format(rsTmp!日期, "yyyy-MM-dd") & IIf(rsTmp!记录状态 = 1, " 结帐", " 退费")
            str(3) = "结帐时间:" & Format(rsTmp!日期, "yyyy-MM-dd hh:mm:ss")
            cllPage.Add str
            'cllPage.Add Array(rsTmp!ID, rsTmp!NO, Format(rsTmp!日期, "yyyy-MM-dd") & IIf(rsTmp!记录状态 = 1, " 结帐", " 退费"), "结帐时间:" & Format(rsTmp!日期, "yyyy-MM-dd hh:mm:ss"))
            intPage = intPage + 1
            intCount = intCount + 1
            If intPage >= 5 Or rsTmp.RecordCount = intCount Then
                mcllBalaceNums.Add cllPage
                Set cllPage = New Collection
                intPage = 0
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '加载分页数据
    If cboNum.Enabled And cboNum.Visible Then
        cboNum.Clear
        For i = 1 To mcllBalaceNums.Count
            cboNum.AddItem "第" & i & "页"
            cboNum.ItemData(cboNum.NewIndex) = i
            If i = 1 Then cboNum.ListIndex = cboNum.NewIndex
        Next
        If mstrRestoreFeeCons <> "" Then
            For i = 0 To cboNum.ListCount - 1
                If Nvl(Split(mstrRestoreFeeCons, "|")(4)) = cboNum.List(i) Then cboNum.ListIndex = i: Exit For
            Next
        End If
    End If
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    Else
        tabClass.Width = Me.ScaleWidth
    End If
    
    LoadPatiClass = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPatiTime() As Boolean
'功能：设置病人住院次数选项卡
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
      
    With tabTime
        mintPreTime = 0
        For i = .Tabs.Count To 2 Step -1    '保留第一个
            .Tabs.Remove i
        Next
        .Visible = tabClass.SelectedItem.Index = conTab未结
        If Not (tabClass.SelectedItem.Index = conTab未结) Then LoadPatiTime = True: Exit Function
        
        On Error GoTo errH
        strSQL = "Select 主页ID,入院日期,出院日期 From 病案主页 Where Nvl(主页ID,0)<>0 And 病人ID=[1] Order by 主页ID Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID)
      
        Do While Not rsTmp.EOF
            .Tabs.Add , "_" & rsTmp!主页ID, "第" & rsTmp!主页ID & "次"
            .Tabs((.Tabs.Count)).Tag = rsTmp!主页ID
            .Tabs((.Tabs.Count)).ToolTipText = "入院:" & Format(rsTmp!入院日期, "yyyy-MM-dd") & _
                                                IIf(Not IsNull(rsTmp!出院日期), ",出院:" & Format(rsTmp!出院日期, "yyyy-MM-dd"), "")
            '问题号:53136 修改人:刘兴洪,修改时间:2012-12-10 13:26:07
            If Val(Nvl(rsTmp!主页ID)) = mlng主页ID Then
                .Tag = "1"
                .Tabs((.Tabs.Count)).Selected = True
                .Tag = ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    LoadPatiTime = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintList(bytStyle As Byte)
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow, strTmp As String, bytR As Byte, lngTmp As Long
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng主页ID As Long
            
    On Error GoTo errH
    Set objOut = New zlPrint1Grd
    '空行
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objOut.UnderAppRows.Add objRow
    
    If tabClass.SelectedItem.Index = conTab未结 Then
        If tabTime.SelectedItem.Index = 1 Then
            '当前病人清单中的信息
            lng主页ID = mlng主页ID
        Else
            '指定住院次数的信息
            lng主页ID = Val(tabTime.SelectedItem.Tag)
        End If
        strSQL = "" & _
        "   Select Nvl(b.姓名, a.姓名) As 姓名,A.住院号,B.出院病床 as 床号," & _
        "           Nvl(b.性别, a.性别) As 性别,Nvl(b.年龄,a.年龄) As 年龄,B.入院日期,B.出院日期,C.名称 as 科室" & _
        "   From 病人信息 A,病案主页 B,部门表 C" & _
        "   Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
        "           And A.病人ID=[1] And B.主页ID=[2] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID, lng主页ID)
        If rsTmp.EOF Then Exit Sub
        
        Set objRow = New zlTabAppRow
        objRow.Add "病人:" & rsTmp!姓名 & "    住院号:" & rsTmp!住院号 & "    床号:" & rsTmp!床号 & "    性别:" & rsTmp!性别 & "    年龄:" & rsTmp!年龄
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "第 " & lng主页ID & " 次住院    科室:" & rsTmp!科室 & _
            "    入院日期:" & Format(rsTmp!入院日期, "yyyy-MM-dd") & "    出院日期:" & Format(Nvl(rsTmp!出院日期), "yyyy-MM-dd")
        objOut.UnderAppRows.Add objRow
    Else
        strSQL = "Select Max(主页ID) as 主页ID From 住院费用记录 Where 结帐ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, Val(tabClass.SelectedItem.Tag))
        If Not rsTmp.EOF Then lng主页ID = Nvl(rsTmp!主页ID, 0)
        If lng主页ID <> 0 Then
            '以结帐最大住院次数为准
            strSQL = "" & _
            "   Select Nvl(b.姓名, a.姓名) As 姓名,A.住院号,B.出院病床 as 床号," & _
            "           Nvl(b.性别, a.性别) As 性别,Nvl(b.年龄, a.年龄) as 年龄,B.入院日期,B.出院日期,C.名称 as 科室" & _
            "   From 病人信息 A,病案主页 B,部门表 C" & _
            "   Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
            "           And A.病人ID=[1] And B.主页ID=[2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID, lng主页ID)
            If rsTmp.EOF Then Exit Sub
        
            Set objRow = New zlTabAppRow
            objRow.Add "病人:" & rsTmp!姓名 & "    住院号:" & rsTmp!住院号 & "    床号:" & rsTmp!床号 & "    性别:" & rsTmp!性别 & "    年龄:" & rsTmp!年龄
            objOut.UnderAppRows.Add objRow
        
            Set objRow = New zlTabAppRow
            objRow.Add "第 " & lng主页ID & " 次住院    科室:" & rsTmp!科室 & _
                "    入院日期:" & Format(rsTmp!入院日期, "yyyy-MM-dd") & "    出院日期:" & Format(Nvl(rsTmp!出院日期), "yyyy-MM-dd")
            objOut.UnderAppRows.Add objRow
        Else
            '结的仅是门诊费用
            strSQL = "Select A.姓名,A.住院号,A.性别,A.年龄 From 病人信息 A Where A.病人ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID)
            If rsTmp.EOF Then Exit Sub
        
            Set objRow = New zlTabAppRow
            objRow.Add "病人:" & rsTmp!姓名 & "    住院号:" & rsTmp!住院号 & "    性别:" & rsTmp!性别 & "    年龄:" & rsTmp!年龄
            objOut.UnderAppRows.Add objRow
        End If
    End If
    
    '空行
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objOut.UnderAppRows.Add objRow
    
    '费用概况情况
    Set objRow = New zlTabAppRow
    objRow.Add lbl费用信息.Caption
    objOut.UnderAppRows.Add objRow
    
    objOut.Title.Font.Size = 16
    If tabClass.SelectedItem.Index = conTab未结 Then
        '费用清单类型
        Dim objControl As CommandBarControl
        Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
        If objControl Is Nothing Then
            strTmp = ""
        Else
            strTmp = objControl.Caption
        End If
        objOut.Title.Text = GetUnitName & "病人未结费用" & Left(strTmp, Len(strTmp) - 4)
    Else
        objOut.Title.Text = GetUnitName & "病人结帐明细清单"
        '空行
        Set objRow = New zlTabAppRow
        objRow.Add ""
        objOut.UnderAppRows.Add objRow
        '结帐情况
        Set objRow = New zlTabAppRow
        objRow.Add "方式:" & Right(tabClass.SelectedItem.Caption, 3)
        objRow.Add "结帐日期:" & Left(tabClass.SelectedItem.Caption, Len(tabClass.SelectedItem.Caption) - 3)
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "备注:"
    objOut.BelowAppRows.Add objRow
    
    If vsfFee.FixedCols = 1 Then
        vsfFee.Redraw = flexRDNone
        vsfFee.OutlineBar = flexOutlineBarNone
        lngTmp = vsfFee.ColWidth(0)
        vsfFee.ColWidth(0) = 0
    End If
    Set objOut.Body = vsfFee
    
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    If vsfFee.FixedCols = 1 Then
        vsfFee.OutlineBar = flexOutlineBarComplete
        vsfFee.ColWidth(0) = lngTmp
        vsfFee.Redraw = flexRDDirect
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadFeeOutline(ByVal lng病人ID As Long, ByVal blnDateMoved As Boolean, ByVal lng结帐ID As Long)
    '功能:显示病人费用概况:预交余额,费用余额,预结费用,或某次结帐费用概况
    '参数:lng病人ID:0-表示仅显示初始信息.
    Dim rsTmp As ADODB.Recordset, strWhere As String
    Dim strSQL As String, strInfo As String, strTmp As String, i As Long
    Dim lngColor As Long
    Dim dblYbMoney As Double, dblTotal As Double
     
    lngColor = ForeColor

    On Error GoTo errH
    'a.未结费用概况
    If lng结帐ID = 0 Then
        If lng病人ID > 0 Then
            If mblnContainOutFee = False Then strWhere = " And 类型=2"
            strSQL = _
                " Select Nvl(预交余额,0) As 预交余额,Nvl(费用余额,0) As 费用余额,0 as 预结费用" & _
                " From 病人余额" & _
                " Where 性质=1 And 病人ID=[1]" & strWhere
            
            If mblnPreBalance Then
                strSQL = strSQL & " Union ALL " & _
                    " Select 0 as 预交余额,0 as 费用余额,Sum(B.金额) as 预结费用" & _
                    " From 病人信息 A,保险模拟结算 B" & _
                    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
            End If
            
            strSQL = _
                " Select Nvl(Sum(预交余额),0) as 预交余额,Nvl(Sum(费用余额),0) as 费用余额,Nvl(Sum(预结费用),0) as 预结费用" & _
                " From (" & strSQL & ")"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng病人ID)
            If rsTmp.RecordCount > 0 Then
                '自付费用:43575
                With rsTmp
                    strInfo = "费用信息：预交款:" & Format(!预交余额, "0.00") & Space(4) & "未结费用:" & Format(!费用余额, gstrDec) & _
                            IIf(mblnPreBalance, Space(4) & "预结费用:" & Format(!预结费用, gstrDec) & Space(4) & "自付费用:" & Format(Val(Nvl(!费用余额)) - Val(Nvl(!预结费用)), gstrDec), "") & _
                            Space(4) & "剩余款:" & Format((!预交余额 - !费用余额 + !预结费用), "0.00")
                            If (Val(Nvl(!预交余额)) - Val(Nvl(!费用余额)) + Val(Nvl(!预结费用))) < 0 Then
                                    lngColor = vbRed
                            End If
                End With
            Else
                strInfo = "费用信息：预交款:0.00" & Space(4) & "未结费用:" & gstrDec & Space(4) & "剩余款:0.00"
            End If
            
            strTmp = GetPatientDue(lng病人ID)
            If Val(strTmp) <> 0 Then strInfo = strInfo & Space(4) & "应收款:" & Format(strTmp, "0.00")
            
            strSQL = "Select 担保人,担保额 From 病人信息 Where 病人ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng病人ID)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!担保人) Or Not IsNull(rsTmp!担保额) Then
                    strInfo = strInfo & "担保人:" & rsTmp!担保人
                    strInfo = strInfo & "担保额:" & Format(Nvl(rsTmp!担保额, 0), "0.00")
                End If
            End If
        Else
            strInfo = "费用信息：预交款:0.00" & Space(4) & "未结费用:" & gstrDec & Space(4) & "剩余款:0.00"
        End If
        
    'b.结帐概况
    Else
        
        strInfo = "单据号:" & Mid(tabClass.SelectedItem.Key, 2)
        
        strSQL = _
            " Select nvl(结帐金额,0) as 结帐金额 From 门诊费用记录 where 结帐ID=[1]" & _
            " Union ALL " & _
            " Select nvl(结帐金额,0) as 结帐金额 From 住院费用记录 where 结帐ID=[1]  "
        If blnDateMoved Then
            strSQL = strSQL & " UNION ALL " & Replace(Replace(strSQL, "门诊费用记录", "H门诊费用记录"), "住院费用记录", "H住院费用记录")
        End If
        
        strSQL = "Select Nvl(Sum(结帐金额),0) as 金额 From  (" & strSQL & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
        
        strInfo = strInfo & Space(4) & "结帐金额:" & Format(rsTmp!金额, gstrDec)
        dblTotal = Val(Nvl(rsTmp!金额))
        strSQL = _
            " Select Decode(Substr(A.记录性质,Length(A.记录性质),1),1,'冲预交',A.结算方式) as 结算方式,Sum(Nvl(冲预交,0)) as 金额," & _
            "               Decode(Substr(A.记录性质,Length(A.记录性质),1),1,0,B.性质,3,1,b.性质,4,1,0) as 医保 " & _
            " From " & IIf(blnDateMoved, zlGetFullFieldsTable("病人预交记录"), "病人预交记录 A") & " ,结算方式 B " & _
            " Where A.结算方式=B.名称(+) And A.结帐ID=[1]" & _
            " Group by Decode(Substr(A.记录性质,Length(A.记录性质),1),1,'冲预交',A.结算方式)," & _
            "       Decode(Substr(a.记录性质, Length(a.记录性质), 1), 1, 0, b.性质, 3, 1, b.性质, 4, 1, 0)" & _
            " Order by 结算方式"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
       
        dblYbMoney = 0
        For i = 1 To rsTmp.RecordCount
            If rsTmp!结算方式 = "冲预交" Then
                strInfo = strInfo & Space(4) & "冲预交:" & Format(rsTmp!金额, "0.00")
            Else
                strInfo = strInfo & Space(4) & IIf(rsTmp!金额 < 0, "退", "收") & rsTmp!结算方式 & ":" & Format(Abs(rsTmp!金额), "0.00")
            End If
            If Val(Nvl(rsTmp!医保)) = 1 Then dblYbMoney = dblYbMoney + Val(Nvl(rsTmp!金额))
            rsTmp.MoveNext
        Next
        If dblYbMoney <> 0 Then '43575
            strInfo = strInfo & Space(4) & "自付金额:" & Format(dblTotal - dblYbMoney, "0.00")
        End If
        strInfo = "结帐信息：" & strInfo
    End If
    
    lbl费用信息.Caption = strInfo
    lbl费用信息.ForeColor = lngColor
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDeposit(ByVal lng病人ID As Long, ByVal blnDateMoved As Boolean, ByVal lng结帐ID As Long)
'功能:显示未冲完的预交款清单或结帐冲预交明细
'参数:lng病人ID:0-表示仅显示初始信息.
    Dim rsTmp As ADODB.Recordset, strWhere As String
    Dim strSQL As String, strDepost As String, i As Long
    
    On Error GoTo errH
    
    'a.未冲完的预交款清单
    If lng结帐ID = 0 Then
        lblDepost.Caption = " 预交款清单"
        If mblnContainOutFee = False Then
            strWhere = " And Nvl(A.预交类别, 2) = 2"
        End If
        
        '日期,单据号,科室,结算方式,结算号码,预交金额,结帐金额,摘要
        strSQL = "Select To_Char(Max(Decode(a.记录性质, 1, a.收款时间, Null)), 'YYYY-MM-DD') As 日期, " & vbNewLine & _
                "       a.NO As 单据号, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, b.名称, Null)) As 科室, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, a.结算方式, Null)) As 结算方式, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, a.结算号码, Null)) as 结算号码, " & vbNewLine & _
                "       To_Char(Sum(Nvl(a.金额, 0)), 'FM9999999990.00') As 预交金额," & vbNewLine & _
                "       To_Char(Sum(Nvl(a.冲预交, 0)), 'FM9999999990.00') As 结帐金额," & vbNewLine & _
                "       To_Char(Sum(Nvl(a.金额, 0)) - Sum(Nvl(a.冲预交, 0)), 'FM9999999990.00') As 剩余金额, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, a.摘要, Null)) as 摘要, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, a.实际票号, Null)) as 实际票号, " & vbNewLine & _
                "       Max(Decode(a.记录性质, 1, a.操作员姓名, Null)) as 操作员姓名," & vbNewLine & _
                "       Decode(a.预交类别,1,'门诊预交','住院预交') As 预交类别" & vbNewLine & _
                " From 病人预交记录 A, 部门表 B" & vbNewLine & _
                " Where A.科室id = B.ID(+) And A.记录性质 In (1, 11) And A.病人id = [1]" & strWhere & vbNewLine & _
                " Group By A.NO,a.预交类别" & vbNewLine & _
                " Having Sum(Nvl(A.金额, 0)) - Sum(Nvl(A.冲预交, 0)) <> 0" & vbNewLine & _
                " Order By 预交类别 Desc,日期, 单据号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng病人ID)
        strDepost = "4,4,1,4,1,7,7,7,1,1,1"
        
     'b.结帐冲预交明细
     Else
        lblDepost.Caption = " 预交使用清单"
        '日期、单据号、科室、结算方式、结算号码、结帐金额、摘要
        strSQL = "Select To_Char(c.收款时间, 'YYYY-MM-DD') As 日期, c.No As 单据号, b.名称 As 科室, c.结算方式, " & vbNewLine & _
                "       c.结算号码, LTrim(To_Char(Nvl(a.冲预交, 0), '9999999990.00')) As 冲预交金额, " & vbNewLine & _
                "       c.摘要, c.实际票号, a.操作员姓名 As 结帐操作员, c.操作员姓名 As 预交收款操作员, " & vbNewLine & _
                "       Decode(c.预交类别,1,'门诊预交','住院预交') As 预交类别" & vbNewLine & _
                " From " & IIf(blnDateMoved, zlGetFullFieldsTable("病人预交记录"), "病人预交记录 A") & ", 部门表 B, " & vbNewLine & _
                        IIf(blnDateMoved, zlGetFullFieldsTable("病人预交记录", , , , "C"), "病人预交记录 C") & vbNewLine & _
                " Where a.No = c.No And c.记录性质 = 1 And c.记录状态 In (1, 3) And c.科室id = b.Id(+)" & vbNewLine & _
                "       And a.记录性质 In (1, 11) And Nvl(a.冲预交, 0) <> 0 And a.结帐id = [1] " & vbNewLine & _
                " Order By 预交类别 Desc,日期, 单据号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng结帐ID)
        strDepost = "4,4,4,1,1,7,1,1,1,1"
    End If
    
    With mshDepost
        .Redraw = flexRDNone
        
        Set .DataSource = rsTmp       '使用此方式,下次有数据时会出现行定位错位,不能显示合计行
        If rsTmp.RecordCount = 0 Then .Rows = 2
        '格式控制
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            .ColKey(i) = Trim(.TextMatrix(0, i))
        Next
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        For i = 0 To UBound(Split(strDepost, ","))
            .ColAlignment(i) = Split(strDepost, ",")(i)
        Next
        zl_vsGrid_Para_Restore mlngModul, mshDepost, Me.Name, "mshDepost", False
        .Redraw = flexRDBuffered
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function LoadCardData(Optional blnHead As Boolean = True, _
    Optional blnDeposit As Boolean = True, Optional blnMoney As Boolean = True, Optional blnLoadDept As Boolean = True) As Boolean
    '功能：根据当前选择的病人费用项目卡片，读取并设置费用清单
    '参数：blnHead=只处理概况部份
    '      blnDeposit=只处理预交款部份
    '      blnMoney=只处理费用部份
    '      blnLoadDept=汇总查询时，是否重新读取科室列表
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str病人费用 As String, strIF As String, strWhere主页 As String
    Dim lng主页ID As Long, lng结帐ID As Long
    Dim strStartDate As String, strEndDate As String, strWhere As String, strDept As String
    Dim strBaby As String, strDateMode As String
    Dim blnDateMoved As Boolean '记录当前选择的日期是否是在后备数据表中
    Dim str科室类型 As String, str分类方式 As String, str金额 As String, str金额合计 As String, str结帐条件 As String, str超期收回 As String
    Dim lng开单部门ID As Long
    Dim lngPre科室ID As Long    '上次科室ID
    Dim strPre科室Text As String, blnNotCheckFee As Boolean
    Dim strWhereCheckFee As String  '体检费用过滤条件
    
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在加载费用信息,请稍候 ...", Me)
    If mlng病人ID <> 0 Then lng主页ID = Val(tabTime.SelectedItem.Tag)
    
    '不含体检费用:46646
    strWhereCheckFee = IIf(chkNotCheckFee.Value = 1, " And nvl(A.门诊标志,0)<>4 ", "")
    
    strDateMode = IIf(mbytDateType = 1, "发生", "登记")
    '最后一项是:病人和婴儿
    If cboBaby.Visible And cboBaby.ListIndex < cboBaby.ListCount - 1 Then strBaby = " And Nvl(A.婴儿费,0)=" & cboBaby.ItemData(cboBaby.ListIndex)
        
    If blnMoney Then
        If mbytList = ListType.C1分科室明细 Then
            If optDeptMode(0).Value = True Then
                str科室类型 = optDeptMode(0).Caption
            Else
                str科室类型 = optDeptMode(1).Caption
            End If
        ElseIf mbytList = ListType.C4分类分项明细 Then
            If optTypeMode(0).Value = True Then
                str分类方式 = optTypeMode(0).Caption
            Else
                str分类方式 = optTypeMode(1).Caption
            End If
        End If
    End If
    
    
    If mlng病人ID <> 0 Then
        If tabClass.SelectedItem.Index = conTab未结 Then '默认显示未结费用,其它情况是某次结帐的费用
            '如果当前病人的入院时间在转出时间之前,则需要联结后备数据表查询
            '入院时间,如果当前选择的是结帐日期,在本条件语句的else中,取当前所选的结帐日期判断
            blnDateMoved = mblnDateMoved
        Else
            blnDateMoved = zlDatabase.DateMoved(Format(Mid(tabClass.SelectedItem.ToolTipText, InStr(1, tabClass.SelectedItem.ToolTipText, ":") + 1), "yyyy-MM-dd hh:mm:ss"), , , Caption)
            
            lng结帐ID = Val(tabClass.SelectedItem.Tag)
        End If
    End If
    
    '费用概况
    If blnHead Then Call LoadFeeOutline(mlng病人ID, blnDateMoved, lng结帐ID)
    '预交款清单
    If blnDeposit Then Call LoadDeposit(mlng病人ID, blnDateMoved, lng结帐ID)
    
    On Error GoTo errH
    '刘兴洪:24913,  mbytDateType:1-发生时间,2-登记时间
    zlGetDateRange , strStartDate, strEndDate
    If strStartDate <> "" Then
        strWhere = IIf(mbytDateType = 1, " And  (A.发生时间 between [4] and [5] ) ", " And  (A.登记时间 between [4] and [5] )")
    Else
        strStartDate = "1901-01-01": strEndDate = "3000-01-01"
    End If
   
    Select Case mbytList
        Case ListType.C0费用清单, ListType.C1分科室明细, ListType.C2分项目明细, ListType.C3分类别明细, ListType.C4分类分项明细 '明细清单,分科明细,项目明细,分类明细,(按收入项目(或收据费目),收费项目,明细分级查询)
            strWhere = strWhere & IIf(chk仅显示销帐单据.Value = 0, "", " And Exists (Select 1 From 住院费用记录 Where NO = a.No And Mod(记录性质, 10) = Mod(a.记录性质, 10) And 序号 = a.序号 And 记录状态 = 3)")
        Case ListType.C5分项目汇总, ListType.C6分类别汇总, ListType.C7分月分类汇总, ListType.C8逐日单据汇总, ListType.C9逐日费目汇总
            If cboDept.ListIndex > 0 And Not blnLoadDept Then   '0是所有部门
                strDept = " And A.开单部门id = [6]"
                strWhere = strWhere & strDept
                lng开单部门ID = cboDept.ItemData(cboDept.ListIndex)
            End If
    End Select
    
    lngPre科室ID = 0
    strPre科室Text = cboDept.Text   ''43494
    If cboDept.ListIndex >= 0 Then
        lngPre科室ID = cboDept.ItemData(cboDept.ListIndex)
    End If
    If blnLoadDept Then
        cboDept.Clear
        cboDept.AddItem "所有科室"
    End If
    
    
    If blnMoney Then
        If lng结帐ID = 0 Then '默认显示未结费用,其它情况是某次结帐的费用
            
            strIF = " And A.记录状态<>0 And A.记帐费用=1" & strBaby & _
                    IIf(mvs.CheckFee, "", " And A.门诊标志<>4") & _
                    IIf(chkAdivce.Value = 0, "", " And A.医嘱序号 is Null") & strWhere
            
            strWhere主页 = IIf(tabTime.SelectedItem.Index = 1, "", " And A.主页ID=[2]")
            
            If mvs.ZeroFee Or (chk仅显示销帐单据.Value = 1 And chk仅显示销帐单据.Visible) Then
                '61527        Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and Sum(Nvl(A.结帐金额,0)) =0 And (Mod(Count(*),2)=0 or  sum(decode(a.结帐ID,null,0,0,1)) = 0)) " & _
                '       :不能用count(*)=1来代替,因为存在结帐一次时,也要显示,不应显示才对,现调整为:sum(decode(a.结帐ID,null,0,0,1)=0
            
                str病人费用 = _
                    "  Select Mod(A.记录性质,10) as 记录性质,A.记录状态,A.发生时间,A.登记时间,A.NO,A.收费细目ID,A.收据费目,A.收费类别,A.开单人,A.开单部门ID,A.执行部门ID,A.计算单位,Max(A.摘要) as 摘要,Max(A.保险编码) as 保险编码," & _
                    "       A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型,Decode(Nvl(A.医嘱序号,0),0,0,(Decode(Sign(A.数次),-1,1,0))) 超期收回,Nvl(A.价格父号,A.序号) as 序号,A.执行状态 as 执行状态" & _
                    "  From 住院费用记录 A" & _
                    "  Where A.病人ID=[1]" & strIF & strWhere主页 & _
                    "           And (Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0) Or Nvl(A.结帐金额, 0)=0)" & _
                    "  Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0 " & _
                    "           Or (Sum(Nvl(A.结帐金额, 0)) = 0 And (Mod(Count(*),2)=0 Or sum(decode(a.结帐ID,null,0,0,1))=0))" & _
                    "  Group by A.NO,Mod(A.记录性质,10),Nvl(A.价格父号,A.序号),A.发生时间,A.登记时间,A.记录状态,A.收费细目ID,A.收据费目,A.收费类别,A.执行状态," & _
                    "          A.开单人,A.开单部门ID,A.执行部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型,Decode(Nvl(A.医嘱序号,0),0,0,(Decode(Sign(A.数次),-1,1,0))),a.医嘱序号 "
                
                If mblnContainOutFee Then
                    str病人费用 = str病人费用 & "  Union ALL " & Replace(str病人费用, "住院费用记录", "门诊费用记录")
                End If
              
            Else
                strSQL = _
                    " Select Distinct NO,Mod(记录性质,10) as 记录性质" & _
                    " From 住院费用记录 A" & _
                    " Where 病人ID=[1]" & strIF & strWhere主页 & _
                    " Group by NO,Mod(记录性质,10),序号" & _
                    " Having Nvl(Sum(实收金额),0)-Nvl(Sum(结帐金额),0)<>0"
                    
                str病人费用 = _
                    " Select /*+ optimizer_features_enable('10.2.0.4') */ Mod(A.记录性质,10) as 记录性质,A.记录状态,A.发生时间,A.登记时间,A.NO,A.收费细目ID,A.收据费目,A.收费类别,A.开单人,A.开单部门ID,A.执行部门ID,A.计算单位,Max(A.摘要) as 摘要,Max(A.保险编码) as 保险编码," & _
                    "        A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型,Decode(Nvl(A.医嘱序号,0),0,0,(Decode(Sign(A.数次),-1,1,0))) 超期收回,Nvl(A.价格父号,A.序号) as 序号,A.执行状态 as 执行状态" & _
                    " From 住院费用记录 A," & _
                    "      (" & strSQL & ") B" & _
                    " Where A.NO=B.NO And Mod(A.记录性质,10)=B.记录性质 " & _
                    "       And A.病人ID+0=[1]" & strIF & strWhere主页 & _
                    "       And Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0)" & _
                    " Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0" & _
                    " Group by A.NO,Mod(A.记录性质,10),Nvl(A.价格父号,A.序号),A.发生时间,A.登记时间,A.记录状态,A.收费细目ID,A.收据费目,A.收费类别,A.执行状态 ," & _
                    "          A.开单人,A.开单部门ID,A.执行部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型,Decode(Nvl(A.医嘱序号,0),0,0,(Decode(Sign(A.数次),-1,1,0))) "
                 
                 If mblnContainOutFee Then
                    str病人费用 = str病人费用 & " Union ALL " & Replace(str病人费用, "住院费用记录", "门诊费用记录")
                 End If
            End If
            
            str结帐条件 = ""
            str金额 = " Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 未结金额,"
            str金额合计 = " Ltrim(To_Char(Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0),'999999999" & gstrDec & "')) as 未结金额,"
            str超期收回 = "A.超期收回,"
        Else
            str病人费用 = "" & _
            " Select 结帐ID,发生时间,NO,序号,医嘱序号,价格父号,记录状态,执行状态,开单部门ID,执行部门id,收费细目ID,收入项目id,开单人,付数,数次,计算单位,标准单价,结帐金额,记录性质,费用类型,收据费目,收费类别,操作员姓名,登记时间,保险编码,摘要  " & _
            " From 住院费用记录 A" & _
            " Where A.结帐ID=[1]  " & strBaby & IIf(chkAdivce.Value = 0, "", " And A.医嘱序号 is Null") & strWhere
            
            str病人费用 = str病人费用 & vbCrLf & " UNION ALL " & vbCrLf & Replace(str病人费用, "住院费用记录", "门诊费用记录")
            If blnDateMoved Then
                str病人费用 = str病人费用 & vbCrLf & " UNION ALL " & vbCrLf & _
                    Replace(Replace(str病人费用, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
            End If
            
            str结帐条件 = " And A.结帐ID=[1]" & strBaby & IIf(chkAdivce.Value = 0, "", " And A.医嘱序号 is Null") & strWhere
            str金额 = " Ltrim(To_Char(A.结帐金额,'999999999" & gstrDec & "')) as 结帐金额,"
            str金额合计 = " Ltrim(To_Char(Nvl(Sum(A.结帐金额),0),'999999999" & gstrDec & "')) as 结帐金额,"
            str超期收回 = "Decode(Nvl(A.医嘱序号,0),0,0,Decode(Sign(A.数次),-1,1,0)) 超期收回,"
        End If
        
        '28078:case  when trunc(A.数次)=0 then  case when A.数次>=0 then '0' else '-0' end when nvl(A.数次,0)<0 then '-' else '' end||abs(A.数次)
        '主要是格式.5的情况.显示类式于0.5或-.5显示为-0.5
        
        Select Case mbytList
            Case ListType.C0费用清单  '明细清单
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.开单人," & _
                "       B.名称 as 开单科室,E.名称 as 执行科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||case  when trunc(A.数次)=0 then  case when A.数次>=0 then '0' else '-0' end when nvl(A.数次,0)<0 then '-' else '' end||abs(A.数次)||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 标准金额," & str金额 & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,N.名称 医保大类,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & str超期收回 & _
                "        Mod(A.记录性质,10) as 记录性质,A.记录状态,A.序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & _
                " FROM (" & str病人费用 & ") A,部门表 B,收费项目目录 C,收费项目别名 D,部门表 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.开单部门ID=B.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=C.ID " & _
                "       And C.ID=M.收费细目ID(+) And M.险类(+)=" & IIf(lng结帐ID = 0, "[3]", "[2]") & " And M.大类ID=N.ID(+)" & vbNewLine & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by 发生日期,单据号,费目"
                
            Case ListType.C1分科室明细  '分科明细
                strSQL = _
                "SELECT " & IIf(str科室类型 = "开单科室", " B.名称 as 开单科室,", " E.名称 as 执行科室,") & "To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.开单人," & _
                        IIf(str科室类型 = "开单科室", " E.名称 as 执行科室,", " B.名称 as 开单科室,") & "Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||case  when trunc(A.数次)=0 then  case when A.数次>=0 then '0' else '-0' end when nvl(A.数次,0)<0 then '-' else '' end||abs(A.数次)||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 标准金额," & str金额 & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,N.名称 医保大类,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & str超期收回 & _
                "       Mod(A.记录性质,10) as 记录性质,A.记录状态,A.序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & _
                " FROM (" & str病人费用 & ") A,部门表 B,收费项目目录 C,收费项目别名 D,部门表 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.开单部门ID=B.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=C.ID " & _
                " And C.ID=M.收费细目ID(+) And M.险类(+)=" & IIf(lng结帐ID = 0, "[3]", "[2]") & " And M.大类ID=N.ID(+)" & vbNewLine & _
                " And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by " & str科室类型 & ",发生日期,单据号"
                
            Case ListType.C2分项目明细  '项目明细
                strSQL = _
                " SELECT Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ') 规格,To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.开单人," & _
                "        B.名称 as 开单科室,E.名称 as 执行科室,A.收据费目 as 费目," & _
                "        Nvl(A.付数,1)*A.数次 as 数量,A.计算单位," & _
                "        Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "        Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 标准金额," & str金额 & _
                "        Nvl(A.费用类型,C.费用类型) as 类型,N.名称 医保大类,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & str超期收回 & _
                "       Mod(A.记录性质,10) as 记录性质,A.记录状态,A.序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & _
                " FROM (" & str病人费用 & ") A,部门表 B,收费项目目录 C,收费项目别名 D,部门表 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.开单部门ID=B.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=C.ID " & _
                "       And C.ID=M.收费细目ID(+) And M.险类(+)=" & IIf(lng结帐ID = 0, "[3]", "[2]") & " And M.大类ID=N.ID(+)" & vbNewLine & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by 项目,规格,发生日期,单据号"
                
            Case ListType.C3分类别明细  '分类明细
                strSQL = _
                " SELECT A.收据费目 as 费目,To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.开单人," & _
                "       B.名称 as 开单科室,E.名称 as 执行科室,Nvl(D.名称,C.名称) as 项目,C.规格," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||case  when trunc(A.数次)=0 then  case when A.数次>=0 then '0' else '-0' end when nvl(A.数次,0)<0 then '-' else '' end||abs(A.数次)||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 标准金额," & str金额 & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,N.名称 医保大类,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & str超期收回 & _
                "       Mod(A.记录性质,10) as 记录性质,A.记录状态,A.序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & _
                " FROM (" & str病人费用 & ") A,部门表 B,收费项目目录 C,收费项目别名 D,部门表 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.开单部门ID=B.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=C.ID " & _
                "       And C.ID=M.收费细目ID(+) And M.险类(+)=" & IIf(lng结帐ID = 0, "[3]", "[2]") & " And M.大类ID=N.ID(+)" & vbNewLine & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by 费目,发生日期,单据号"
            Case ListType.C4分类分项明细 '按收入项目(或收据费目),收费项目,明细分级查询
                If str分类方式 = "收入项目" Then
                    If lng结帐ID = 0 Then
                        str病人费用 = " With PatiFee as (" & str病人费用 & ")"
                        
                        str病人费用 = str病人费用 & "" & _
                        " Select  distinct A.发生时间,A.NO,A.开单人,A.收据费目," & _
                        "        A.付数*decode(A.记录性质,12,0,13,0,1) as 付数,A.数次*decode(A.记录性质,12,0,13,0,1) as 数次, " & _
                        "       A.计算单位,A.标准单价,A.实收金额,A.结帐金额,A.费用类型,A.操作员姓名,A.登记时间," & _
                        "       A.医嘱序号,A.价格父号,A.保险编码 ,A.序号,A.记录状态,A.执行状态,A.开单部门ID,A.执行部门id,A.收费细目ID,A.收入项目id,A.记录性质,A.摘要  " & _
                        " From 住院费用记录 A ,PatiFee G" & _
                        " Where A.NO = G.NO And Mod(A.记录性质,10)=G.记录性质 And Nvl(A.价格父号,A.序号)=G.序号 And A.记录状态<>0 "
                        If mblnContainOutFee Then
                            str病人费用 = str病人费用 & " Union ALL " & _
                            " Select distinct  A.发生时间,A.NO,A.开单人,A.收据费目, " & _
                            "        A.付数*decode(A.记录性质,12,0,13,0,1) as 付数,A.数次*decode(A.记录性质,12,0,13,0,1) as 数次, " & _
                            "       A.计算单位,A.标准单价,A.实收金额,A.结帐金额,A.费用类型,A.操作员姓名,A.登记时间," & _
                            "       A.医嘱序号,A.价格父号,A.保险编码 ,A.序号,A.记录状态,A.执行状态,A.开单部门ID,A.执行部门id,A.收费细目ID,A.收入项目id,A.记录性质,A.摘要  " & _
                            " From 门诊费用记录 A ,PatiFee G" & _
                            " Where A.NO = G.NO And Mod(A.记录性质,10)=G.记录性质 And Nvl(A.价格父号,A.序号)=G.序号 And A.记录状态<>0 " & strWhereCheckFee & _
                            ""
                        End If
                        
                        '结帐作废后再销帐，记录性质为12的记录状态是1,记录性质为2的记录状态是3,要汇总起来需要Decode
                        strSQL = "Select F.名称 收入项目, Nvl(D.名称, C.名称) As 收费项目, To_Char(A.发生时间, 'YYYY-MM-DD') As 发生日期, A.NO As 单据号,A.开单人," & vbNewLine & _
                            "       B.名称 As 开单科室, E.名称 As 执行科室, C.规格, A.收据费目 As 费目," & vbNewLine & _
                            "       sum(Nvl(A.付数,1)*A.数次)  as 数量,A.计算单位," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.标准单价, 0), '999999999" & gstrFeePrecisionFmt & "')) As 标准单价," & vbNewLine & _
                            "       LTrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) As 标准金额," & str金额合计 & vbNewLine & _
                            "       Nvl(A.费用类型, C.费用类型) As 类型,N.名称 医保大类, A.操作员姓名 As 操作员," & vbNewLine & _
                            "       To_Char(A.登记时间, 'YYYY-MM-DD HH24:MI:SS') As 登记时间,Decode(Nvl(A.医嘱序号,0),0,0,Decode(Sign(A.数次),-1,1,0)) 超期收回," & _
                            "       Mod(A.记录性质,10) as 记录性质,Decode(A.记录状态,3,1,A.记录状态) as 记录状态,Nvl(A.价格父号,A.序号) 序号,max(C.编码) as 项目编码, max(A.保险编码) 保险编码," & _
                            "       max(C.说明) as 项目说明,max(A.摘要) as 摘要,Decode(Decode(A.记录状态,3,1,A.记录状态),2,'',Decode(Max(A.执行状态),0,'未执行','已执行')) as 执行状态" & vbNewLine & _
                            "From  (" & str病人费用 & ") A, 部门表 B, 收费项目目录 C, 收费项目别名 D, 部门表 E, 收入项目 F,保险支付项目 M,保险支付大类 N" & vbNewLine & _
                            "Where   A.开单部门id = B.ID(+) And A.执行部门id = E.ID(+) And A.收费细目id = C.ID And A.收入项目id = F.ID " & vbNewLine & _
                            "      And C.ID=M.收费细目ID(+) And M.险类(+)=[3] And M.大类ID=N.ID(+)" & vbNewLine & _
                            "      And A.收费细目id = D.收费细目id(+) And D.码类(+) = 1 And D.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & vbNewLine & _
                            " Group By F.名称, Nvl(D.名称, C.名称), To_Char(A.发生时间, 'YYYY-MM-DD'), A.NO,A.开单人,B.名称, E.名称, C.规格, " & vbNewLine & _
                            "          A.收据费目,A.计算单位,Nvl(A.标准单价, 0),Nvl(A.费用类型, C.费用类型),N.名称, A.操作员姓名," & vbNewLine & _
                            "          A.登记时间,Decode(Nvl(A.医嘱序号,0),0,0,Decode(Sign(A.数次),-1,1,0)),Mod(A.记录性质,10),Decode(A.记录状态,3,1,A.记录状态),Nvl(A.价格父号,A.序号)" & vbNewLine & _
                            " Order By 收入项目,收费项目,发生日期"
                    Else
                        strSQL = "Select F.名称 收入项目, Nvl(D.名称, C.名称) As 收费项目, To_Char(A.发生时间, 'YYYY-MM-DD') As 发生日期, A.NO As 单据号,A.开单人," & vbNewLine & _
                            "       B.名称 As 开单科室, E.名称 As 执行科室, C.规格, A.收据费目 As 费目," & vbNewLine & _
                            "       Nvl(A.付数,1)*A.数次 as 数量,A.计算单位," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.标准单价, 0), '999999999" & gstrFeePrecisionFmt & "')) As 标准单价," & vbNewLine & _
                            "       LTrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) As 标准金额," & str金额 & vbNewLine & _
                            "       Nvl(A.费用类型, C.费用类型) As 类型,N.名称 医保大类, A.操作员姓名 As 操作员," & vbNewLine & _
                            "       To_Char(A.登记时间, 'YYYY-MM-DD HH24:MI:SS') As 登记时间," & str超期收回 & _
                            "       Mod(A.记录性质,10) as 记录性质,A.记录状态,Nvl(A.价格父号,A.序号) 序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & vbNewLine & _
                            "From (" & str病人费用 & ") A, 部门表 B, 收费项目目录 C, 收费项目别名 D, 部门表 E, 收入项目 F,保险支付项目 M,保险支付大类 N" & vbNewLine & _
                            "Where A.开单部门id = B.ID(+) And A.执行部门id = E.ID(+) And A.收费细目id = C.ID And A.收入项目id = F.ID " & vbNewLine & _
                            "      And C.ID=M.收费细目ID(+) And M.险类(+)=[2] And M.大类ID=N.ID(+)" & vbNewLine & _
                            "      And A.收费细目id = D.收费细目id(+) And D.码类(+) = 1 And D.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & vbNewLine & _
                            "Order By 收入项目,收费项目,发生日期"
                    End If
                Else
                    strSQL = "Select A.收据费目 As 费目, Nvl(D.名称, C.名称) As 收费项目, To_Char(A.发生时间, 'YYYY-MM-DD') As 发生日期, A.NO As 单据号,A.开单人," & vbNewLine & _
                            "       B.名称 As 开单科室, E.名称 As 执行科室, C.规格, " & vbNewLine & _
                            "       Nvl(A.付数,1)*A.数次 as 数量,A.计算单位," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.标准单价, 0), '999999999" & gstrFeePrecisionFmt & "')) As 标准单价," & vbNewLine & _
                            "       LTrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) As 标准金额," & str金额 & vbNewLine & _
                            "       Nvl(A.费用类型, C.费用类型) As 类型,N.名称 医保大类, A.操作员姓名 As 操作员," & vbNewLine & _
                            "       To_Char(A.登记时间, 'YYYY-MM-DD HH24:MI:SS') As 登记时间," & str超期收回 & _
                            "       Mod(A.记录性质,10) as 记录性质,A.记录状态,序号,C.编码 as 项目编码,A.保险编码,C.说明 as 项目说明,A.摘要,Decode(A.记录状态,2,'',Decode(A.执行状态,0,'未执行','已执行')) as 执行状态" & vbNewLine & _
                            "From (" & str病人费用 & ") A, 部门表 B, 收费项目目录 C, 收费项目别名 D, 部门表 E,保险支付项目 M,保险支付大类 N" & vbNewLine & _
                            "Where A.开单部门id = B.ID(+) And A.执行部门id = E.ID(+) And A.收费细目id = C.ID " & vbNewLine & _
                            "      And C.ID=M.收费细目ID(+) And M.险类(+)=" & IIf(lng结帐ID = 0, "[3]", "[2]") & " And M.大类ID=N.ID(+)" & vbNewLine & _
                            "      And A.收费细目id = D.收费细目id(+) And D.码类(+) = 1 And D.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & str结帐条件 & vbNewLine & _
                            "Order By 费目,收费项目,发生日期"
                End If
                
            Case ListType.C5分项目汇总  '项目汇总
                strSQL = "case  when trunc(Sum(Nvl(A.数次,1)*Nvl(A.付数,1)))=0 then  case when Sum(Nvl(A.数次,1)*Nvl(A.付数,1))>0 then  '0' when Sum(Nvl(A.数次,1)*Nvl(A.付数,1))=0 then '' else '-0' end when Sum(Nvl(A.数次,1)*Nvl(A.付数,1))<0 then '-' else '' end||abs(Sum(Nvl(A.数次,1)*Nvl(A.付数,1)))"
                
                strSQL = _
                " SELECT nvl(Q.类别,A.收费类别) as 收费类别, Nvl(D.名称,C.名称) as 项目,C.规格," & strSQL & "||Max(A.计算单位) 数量," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 标准金额," & Mid(str金额合计, 1, Len(str金额合计) - 1) & _
                " FROM (" & str病人费用 & ") A,收费项目目录 C,收费项目别名 D, 收费类别 Q" & _
                " Where A.收费细目ID=C.ID And A.收费细目ID=D.收费细目ID(+) And A.收费类别=Q.编码(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Group by nvl(Q.类别,A.收费类别), Nvl(D.名称,C.名称),规格" & _
                " Order by 收费类别,项目,规格"
                    
            Case ListType.C6分类别汇总  '分类汇总
                'If str结帐条件 <> "" Then str结帐条件 = " Where " & Mid(str结帐条件, InStr(1, str结帐条件, "And") + 3)
                strSQL = _
                " SELECT A.收据费目 as 费目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 标准金额," & Mid(str金额合计, 1, Len(str金额合计) - 1) & _
                " FROM (" & str病人费用 & ") A " & _
                " Group by A.收据费目 " & _
                " Order by 费目"
            Case ListType.C7分月分类汇总  '分月汇总
                strSQL = _
                " SELECT B.期间,A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 标准金额," & Mid(str金额合计, 1, Len(str金额合计) - 1) & _
                " FROM (" & str病人费用 & ") A,期间表 B" & _
                " Where A." & strDateMode & "时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                " Group by B.期间,A.收据费目" & _
                " Order by 期间,费目"
                    
            Case ListType.C8逐日单据汇总  '逐日费用
                'If str结帐条件 <> "" Then str结帐条件 = " Where " & Mid(str结帐条件, InStr(1, str结帐条件, "And") + 3)
                strSQL = _
                " SELECT TO_Char(A." & strDateMode & "时间,'YYYY-MM-DD') as " & strDateMode & "日期,A.NO as 单据号,A.收据费目 as 费用项目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 标准金额," & str金额合计 & _
                "       A.操作员姓名 as 操作员,A.记录性质" & _
                " FROM (" & str病人费用 & ") A" & _
                " Group by TO_Char(A." & strDateMode & "时间,'YYYY-MM-DD'),A.NO,A.记录性质,A.收据费目,A.操作员姓名" & _
                " Order by " & strDateMode & "日期,记录性质 desc,单据号,费用项目"
            
            Case ListType.C9逐日费目汇总  '逐日费目
                'If str结帐条件 <> "" Then str结帐条件 = " Where " & Mid(str结帐条件, InStr(1, str结帐条件, "And") + 3)
                strSQL = _
                " SELECT TO_Char(A." & strDateMode & "时间,'YYYY-MM-DD') as " & strDateMode & "日期,A.收据费目 as 费用项目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 标准金额," & Mid(str金额合计, 1, Len(str金额合计) - 1) & _
                " FROM (" & str病人费用 & ") A" & _
                " Group by TO_Char(A." & strDateMode & "时间,'YYYY-MM-DD'),A.收据费目" & _
                " Order by " & strDateMode & "日期,费用项目"
        End Select
                    
        If lng结帐ID = 0 Then
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID, lng主页ID, mintInsure, CDate(strStartDate), CDate(strEndDate), lng开单部门ID)
        Else
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Caption, lng结帐ID, mintInsure, 0, CDate(strStartDate), CDate(strEndDate), lng开单部门ID)
        End If
        If mbytList = ListType.C0费用清单 Or mbytList = ListType.C2分项目明细 Then Call LoadCbo费目
        
        If mbytList <> ListType.C1分科室明细 Then
            If mbytList = ListType.C0费用清单 Or mbytList = ListType.C2分项目明细 Or mbytList = ListType.C3分类别明细 Or mbytList = ListType.C4分类分项明细 Then
                mblnNotClick = True
                Call LoadCbo开单科室(mrsList, False)
                If strPre科室Text <> "" Then
                    Call zlControl.CboLocate(cboDept, strPre科室Text)
                End If
                mblnNotClick = False
            ElseIf blnLoadDept Then    '分类汇总，重读开单部门列表(在当前查询类别，选择一个部门时，不用重读)
                strSQL = "Select Distinct B.名称 as 开单科室,a.开单部门ID From (" & Replace(str病人费用, strDept, "") & ") A,部门表 B Where a.开单部门ID = b.ID"
                If lng结帐ID = 0 Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng病人ID, lng主页ID, mintInsure, CDate(strStartDate), CDate(strEndDate), lng开单部门ID)
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng结帐ID, mintInsure, 0, CDate(strStartDate), CDate(strEndDate), lng开单部门ID)
                End If
                mblnNotClick = True
                Call LoadCbo开单科室(rsTmp, True)
                If strPre科室Text <> "" Then
                    Call zlControl.CboLocate(cboDept, strPre科室Text)
                End If
                mblnNotClick = False
            End If
        End If
        
        With vsfFee
            .Redraw = flexRDNone
            .Clear
                        
            '必须在绑定数据前设置,有树型大纲时要设置为1
            If mbytList = ListType.C0费用清单 Or mbytList = ListType.C5分项目汇总 Or mbytList = ListType.C6分类别汇总 Then
                .FixedCols = 0
            Else
                .FixedCols = 1
                .OutlineCol = 0
                .OutlineBar = flexOutlineBarComplete
            End If
            Set .DataSource = mrsList
            Call SetVsffeeFormat
            
            '恢复个性化设置
            zl_vsGrid_Para_Restore mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
            
            If mbytList = ListType.C0费用清单 Or mbytList = ListType.C1分科室明细 Or mbytList = ListType.C2分项目明细 Or mbytList = ListType.C3分类别明细 Or mbytList = ListType.C4分类分项明细 Then
            
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                .ColData(.ColIndex("记录性质")) = "-1||2"
                .ColData(.ColIndex("记录状态")) = "-1||2"
                .ColData(.ColIndex("序号")) = "-1||2"
                .ColData(.ColIndex("超期收回")) = "-1||2"
                
                .ColWidth(.ColIndex("记录性质")) = 0    '如果用ColHidden方式,打印预览仍可见
                .ColWidth(.ColIndex("记录状态")) = 0
                .ColWidth(.ColIndex("序号")) = 0
                .ColWidth(.ColIndex("超期收回")) = 0
                
                .ColHidden(.ColIndex("记录性质")) = True
                .ColHidden(.ColIndex("记录状态")) = True
                .ColHidden(.ColIndex("序号")) = True
                .ColHidden(.ColIndex("超期收回")) = True
                If .ColIndex("开单人") >= 0 Then
                    '问题:35710
                    If InStr(1, mstrPrivs, ";医生查询;") = 0 Then
                        .ColHidden(.ColIndex("开单人")) = True
                        .ColWidth(.ColIndex("开单人")) = 0
                        .ColData(.ColIndex("开单人")) = "-1||2"
                    End If
                End If
                
                If mintInsure = 0 Then
                    .ColWidth(.ColIndex("医保大类")) = 0
                    .ColHidden(.ColIndex("医保大类")) = True
                    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                    .ColData(.ColIndex("医保大类")) = "-1||2"
                End If
            ElseIf mbytList = ListType.C8逐日单据汇总 Then
                .ColWidth(.ColIndex("记录性质")) = 0
                .ColHidden(.ColIndex("记录性质")) = True
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                .ColData(.ColIndex("记录性质")) = "-1||2"
            End If
            'If .Rows = 1 Then .Rows = .FixedRows + 1
            
            .Redraw = flexRDDirect
        End With
    End If
    If mstrRestoreFeeCons <> "" Then
        If zlRestoreFeeControls(mlng病人ID) = False Then
            Call zlCommFun.StopFlash
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    '55107
    Select Case mbytList
    Case ListType.C0费用清单, ListType.C1分科室明细, ListType.C2分项目明细, ListType.C3分类别明细, ListType.C4分类分项明细
        Call FilterDetail
    End Select
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    vsfFee.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetVsffeeFormat()
    Dim i As Long, j As Long, lng超期收回 As Long, lng标准金额 As Long, lng未结金额 As Long, lng数量 As Long, lng科室 As Long
    Dim arrTotal(2) As Currency, strTmp As String, lng记录状态 As Long
    Dim blnSetColor As Boolean
    Dim bln规格分类 As Boolean
    
    With vsfFee
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignGeneral
            .ColKey(i) = Trim(.TextMatrix(0, i))
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            .ColData(i) = "0||0"
        Next
        If .Rows <= 1 Then .Rows = 2
        lng数量 = .ColIndex("数量")
        If lng数量 >= 0 Then .ColAlignment(lng数量) = flexAlignRightCenter: .ColFormat(lng数量) = "#######0.#####"
    
        If mrsList.RecordCount = 0 Then Exit Sub
        
        lng标准金额 = .ColIndex("标准金额")
        If lng标准金额 < 0 Then lng标准金额 = VsfGetColNum(vsfFee, "标准金额")  '当SQL中使用了Group by后,绑定的数据列的Key没有自动加上,用colindex方式取不出来
        
        If Val(tabClass.SelectedItem.Tag) = 0 Then
            lng未结金额 = .ColIndex("未结金额")
            If lng未结金额 < 0 Then lng未结金额 = VsfGetColNum(vsfFee, "未结金额")
        Else
            lng未结金额 = .ColIndex("结帐金额")
            If lng未结金额 < 0 Then lng未结金额 = VsfGetColNum(vsfFee, "结帐金额")
        End If
               
        Select Case mbytList
            Case ListType.C0费用清单, ListType.C5分项目汇总, ListType.C6分类别汇总
                .Subtotal flexSTSum, -1, lng标准金额, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "总计"
                .Subtotal flexSTSum, -1, lng未结金额, "#######" & gstrDec
                .MergeRow(.Rows - 1) = True
            Case ListType.C2分项目明细
                bln规格分类 = chk按规格进行小计.Value = 1
                .Subtotal flexSTSum, 0, lng标准金额, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "总计"
                .Subtotal flexSTSum, 1, lng标准金额, "#######" & gstrDec, &HF5F5F5, vbBlack, True, IIf(bln规格分类, "合计", "小计")
               If bln规格分类 Then
                 .Subtotal flexSTSum, 2, lng标准金额, "#######" & gstrDec, &HF5F5F5, vbBlack, True, "小计"
                End If
                .Subtotal flexSTSum, 0, lng未结金额, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lng未结金额, "#######" & gstrDec
                
                If bln规格分类 Then
                    .Subtotal flexSTSum, 2, lng未结金额, "#######" & gstrDec
                    .Subtotal flexSTSum, 1, lng数量, "#######0.#####"
                    .Subtotal flexSTSum, 2, lng数量, "#######0.#####"
                Else
                    .Subtotal flexSTSum, 1, lng数量, "#######0.#####"
                End If
                
                
                .MergeCol(1) = True
                .MergeRow(.Rows - 1) = True
            Case ListType.C1分科室明细, ListType.C3分类别明细, ListType.C7分月分类汇总, ListType.C9逐日费目汇总
                                
                If mbytList = ListType.C1分科室明细 Or mbytList = ListType.C3分类别明细 Then strTmp = "%s "
                
                .Subtotal flexSTSum, 0, lng标准金额, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "总计"
                .Subtotal flexSTSum, 1, lng标准金额, "#######" & gstrDec, &HF5F5F5, vbBlack, True, strTmp & "小计"
                            
                .Subtotal flexSTSum, 0, lng未结金额, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lng未结金额, "#######" & gstrDec
                If mbytList = ListType.C2分项目明细 Then .Subtotal flexSTSum, 1, lng数量, "#######0.#####"
                .MergeCol(1) = True
                .MergeRow(.Rows - 1) = True
                
            Case ListType.C4分类分项明细, ListType.C8逐日单据汇总
                               
                If mbytList = ListType.C4分类分项明细 Then strTmp = "%s "
                
                .Subtotal flexSTSum, 0, lng标准金额, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "总计"
                .Subtotal flexSTSum, 1, lng标准金额, "#######" & gstrDec, &HF1E8FC, vbBlack, True, strTmp & "合计"
                .Subtotal flexSTSum, 2, lng标准金额, "#######" & gstrDec, &HF5F5F5, vbBlack, True, "小计"
                            
                .Subtotal flexSTSum, 0, lng未结金额, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lng未结金额, "#######" & gstrDec
                .Subtotal flexSTSum, 2, lng未结金额, "#######" & gstrDec
                          
                If mbytList = ListType.C4分类分项明细 Then .Subtotal flexSTSum, 2, lng数量, "#######0.#####"
                
                .MergeCol(1) = True
                .MergeCol(2) = True
                
        End Select
             
        lng超期收回 = .ColIndex("超期收回")
        lng记录状态 = .ColIndex("记录状态") ' '30289
        lng数量 = .ColIndex("数量")
        If lng超期收回 >= 0 Or lng记录状态 >= 0 Then
            For i = 1 To .Rows - 1
                blnSetColor = False
                
                If lng超期收回 >= 0 Then
                    If Val(.TextMatrix(i, lng超期收回)) = 1 Then blnSetColor = True
                End If
                If blnSetColor = False And lng记录状态 >= 0 Then
                    If Val(.TextMatrix(i, lng记录状态)) = 2 Then blnSetColor = True
                    '负数也要用红色显示
                    If blnSetColor = False And InStr(1, "31", Val(.TextMatrix(i, lng记录状态))) > 0 And lng数量 >= 0 Then
                        If Left(Trim(.TextMatrix(i, lng数量)), 1) = "-" Then blnSetColor = True
                    End If
                End If
                If blnSetColor Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC0  '超期收回行,用深红色表示,须.FillStyle = flexFillRepeat
                End If
            
            Next
        End If
        
        .AutoSize 0, .Cols - 1
    End With
End Sub

Private Sub LoadCbo费目()
    Dim str费目 As String
    Dim i As Integer
    Dim strOld As String
    Dim strPreText As String
    strPreText = cboFeeType.Text    '43494
    cboFeeType.Clear '此时做为费目,统计方式为2时做为科室类型
    
    str费目 = ";所有费目"
    
    If Not mrsList Is Nothing Then
        If mrsList.RecordCount > 0 Then mrsList.MoveFirst
        Do While Not mrsList.EOF
            If strOld <> mrsList!费目 Then
                If InStr(1, ";" & str费目 & ";", ";" & mrsList!费目 & ";") = 0 Then
                    str费目 = str费目 & ";" & mrsList!费目
                End If
                strOld = mrsList!费目
            End If
            mrsList.MoveNext
        Loop
    End If
    
    str费目 = Mid(str费目, 2)
    mblnNotClick = True
    For i = 0 To UBound(Split(str费目, ";"))
        cboFeeType.AddItem Split(str费目, ";")(i)
    Next
    zlControl.CboSetIndex cboFeeType.hWnd, 0
    If strPreText <> "" Then zlControl.CboLocate cboFeeType, strPreText
    mblnNotClick = False

End Sub

Private Sub LoadCbo开单科室(ByRef rsTmp As ADODB.Recordset, ByVal blnAddID As Boolean)
    Dim str开单科室 As String
    Dim i As Integer
    Dim strOld As String
            
    If blnAddID Then
        For i = 0 To rsTmp.RecordCount - 1
            cboDept.AddItem rsTmp!开单科室
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!开单部门ID
            
            If mblnClinicOrNurse And cboDept.ListIndex = -1 Then
                If rsTmp!开单部门ID = UserInfo.部门ID Then
                    cboDept.ListIndex = cboDept.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Next
    Else
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If strOld <> rsTmp!开单科室 Then
                If InStr(1, ";" & str开单科室 & ";", ";" & rsTmp!开单科室 & ";") = 0 Then
                    str开单科室 = str开单科室 & ";" & rsTmp!开单科室
                End If
                strOld = rsTmp!开单科室
            End If
            rsTmp.MoveNext
        Loop
        str开单科室 = Mid(str开单科室, 2)
        
        For i = 0 To UBound(Split(str开单科室, ";"))
            cboDept.AddItem Split(str开单科室, ";")(i)
            If mblnClinicOrNurse And cboDept.ListIndex = -1 Then
                If Split(str开单科室, ";")(i) = UserInfo.部门名称 Then
                    cboDept.ListIndex = cboDept.NewIndex
                End If
            End If
        Next
    End If
    cboDept.Tag = "不刷新"
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then zlControl.CboSetIndex cboDept.hWnd, 0
    cboDept.Tag = ""
End Sub

Private Sub chkNotCheckFee_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk按规格进行小计_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk仅显示销帐单据_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk仅显示销帐单据_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdRefresh_Click()
    Call LoadCardData(False, False, True)
    If vsfFee.Enabled And vsfFee.Visible Then vsfFee.SetFocus
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Call SetCondition
    Call Form_Resize
    Call picDetail_Resize
    Call vsfFee_LostFocus
    Call mshDepost_LostFocus
    Call mshInsure_LostFocus
    RaiseEvent Activate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If tabClass.SelectedItem Is Nothing Then Exit Sub
            Call lblMoney_MouseDown(1, 0, 0, 0)
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, strTmp As String
    mstrPrivs = gstrPrivs
    mblnFisrtSetFontSize = True
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsTools.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsTools.VisualTheme = xtpThemeOffice2003
    cbsTools.EnableCustomization False
    Set cbsTools.Icons = zlCommFun.GetPubIcons
    
    fraDeptMode.BackColor = picDetail.BackColor
    fraTypeMode.BackColor = picDetail.BackColor
 
    mstr截止日期 = ""
    mblnContainOutFee = zlDatabase.GetPara("包含门诊费用", glngSys, mlngModul, "1") = "1"
    msngScale = CSng(zlDatabase.GetPara("清单比例", glngSys, mlngModul, 0.75))
    mbytDateType = IIf(zlDatabase.GetPara("费用时间类型", glngSys, mlngModul, "1") = "2", 2, 1)
    lbl发生时间.Caption = IIf(mbytDateType = 1, "发生时间", "登记时间")
    dtpEnd.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    dtpEnd.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    
    dtpBegin.Value = Format(DateAdd("m", -1, dtpEnd.Value), "yyyy-mm-dd 00:00:00")
    dtpBegin.MaxDate = dtpEnd.Value
    
    
    mvs.ReBalance = zlDatabase.GetPara("显示结帐作废", glngSys, mlngModul, "1") = "1"
    mvs.ZeroFee = zlDatabase.GetPara("显示零费用", glngSys, mlngModul, "0") = "1"
    mvs.CheckFee = zlDatabase.GetPara("显示体检费用", glngSys, mlngModul, "0") = "1"
    
    i = IIf(zlDatabase.GetPara("分科模式", glngSys, mlngModul) = "1", 1, 0)
    optDeptMode(i).Value = True
    i = IIf(zlDatabase.GetPara("分类模式", glngSys, mlngModul) = "1", 1, 0)
    optTypeMode(i).Value = True
    
    mblnClinicOrNurse = isCliniOrNurse(UserInfo.部门ID)
    
    mstrUnitIDs = GetUserUnits
    Call InitBaseData
    
    With vsfFee
        .ExplorerBar = flexExSortShowAndMove
        .FillStyle = flexFillRepeat
        .FixedRows = 1
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat  '缺省是flexGridInset
                    
        .Subtotal flexSTClear
        .SubtotalPosition = flexSTBelow
    End With
    
    
    
End Sub

Private Sub Form_Resize()
    Dim tmpW As Long, tmpH As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    tmpW = IIf(mshInsure.Visible, mshInsure.Width + picLR.Width, 0)
    Call cbsTools.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    tabClass.Top = lngTop
    tabClass.Left = lngLeft
    picNum.Top = lngTop + 30
    tabClass.Width = Me.ScaleWidth
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    End If
    
    pic费用信息.Left = tabClass.Left
    pic费用信息.Top = tabClass.Top + tabClass.Height
    pic费用信息.Width = Me.ScaleWidth
    
    lblDepost.Top = pic费用信息.Top + pic费用信息.Height + 30
    lblDepost.Left = pic费用信息.Left + 15
    lblDepost.Width = pic费用信息.Width - 30 - tmpW
    mshDepost.Redraw = flexRDNone
    mshDepost.Top = lblDepost.Top + lblDepost.Height
    mshDepost.Left = pic费用信息.Left
    tmpH = (Me.ScaleHeight - tabClass.Height - pic费用信息.Height - lblDepost.Height - picDetail.Height - fraUD.Height - 30) * (1 - msngScale)
    If tmpH > 0 Then mshDepost.Height = tmpH
    mshDepost.Width = pic费用信息.Width - tmpW
    mshDepost.Redraw = flexRDBuffered
    
    picLR.Left = mshDepost.Left + mshDepost.Width
    picLR.Top = mshDepost.Top
    picLR.Height = mshDepost.Height
    
    lblInsure.Top = lblDepost.Top
    lblInsure.Left = picLR.Left + picLR.Width
    lblInsure.Width = mshInsure.Width - 30
    
    mshInsure.Top = mshDepost.Top
    mshInsure.Left = picLR.Left + picLR.Width
    mshInsure.Height = mshDepost.Height
    
    fraUD.Top = mshDepost.Top + mshDepost.Height
    fraUD.Left = pic费用信息.Left
    fraUD.Width = pic费用信息.Width
    
    picDetail.AutoRedraw = False
    picDetail.Top = fraUD.Top + fraUD.Height
    picDetail.Left = fraUD.Left
    picDetail.Width = fraUD.Width
    picDetail.AutoRedraw = True
    
    vsfFee.Redraw = flexRDNone
    vsfFee.Top = picDetail.Top + picDetail.Height
    vsfFee.Left = pic费用信息.Left
    vsfFee.Width = pic费用信息.Width
    vsfFee.Height = Me.ScaleHeight - lngTop - tabClass.Height - pic费用信息.Height - lblDepost.Height - picDetail.Height - fraUD.Height - mshDepost.Height - 30
    vsfFee.Redraw = flexRDDirect

    zlControl.PicShowFlat pic费用信息, -1, , taCenterAlign
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    zlDatabase.SetPara "清单比例", msngScale, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "费用时间类型", mbytDateType, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "显示结帐作废", IIf(mvs.ReBalance, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "显示零费用", IIf(mvs.ZeroFee, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "显示体检费用", IIf(mvs.CheckFee, 1, 0), glngSys, mlngModul, mblnHavePara
    
    zlDatabase.SetPara "分科模式", IIf(optDeptMode(0).Value, 0, 1), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "分类模式", IIf(optTypeMode(0).Value, 0, 1), glngSys, mlngModul, mblnHavePara
    
    Call zlDatabase.SetPara("费用查询范围", cbo日期.Text & IIf(cbo日期.Text = "自定义范围", "|" & Format(dtpBegin.Value, "yyyy-mm-dd") & "," & Format(dtpEnd.Value, "yyyy-mm-dd"), ""), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("明细仅显销帐单据", IIf(chk仅显示销帐单据.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("按规格分类统计", IIf(chk按规格进行小计.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    '46646
    Call zlDatabase.SetPara("不含体检费用", IIf(chkNotCheckFee.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("包含门诊费用", IIf(mblnContainOutFee, 1, 0), glngSys, mlngModul, mblnHavePara)
 
    '保存过性化设置
    If mbytFontSize <> 9 Then
        zlControl.VSFSetFontSize vsfFee, 9
        zlControl.VSFSetFontSize mshDepost, 9
        zlControl.VSFSetFontSize mshInsure, 9
    End If
    
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False, , mblnHavePara
    Set mrsList = Nothing
    mbytList = ListType.C0费用清单
    mintPreCard = 0: mintPreTime = 0
    Unload frmDailyListAsk
End Sub
 

Private Sub imgColSel_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picDetail.hWnd)
    lngLeft = vRect.Left + imgColSel.Left
    lngTop = vRect.Top + imgColSel.Height + imgColSel.Top
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfFee, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
End Sub
 

Private Sub mshDepost_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False
End Sub

Private Sub mshDepost_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False
End Sub


Private Sub mshInsure_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False
End Sub

Private Sub mshInsure_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False
End Sub

Private Sub pic费用信息_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pic费用信息.ToolTipText = lbl费用信息.Caption
End Sub


Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDepost.Height + Y < 268 Or vsfFee.Height - Y < 1000 Then Exit Sub
        
        fraUD.Top = fraUD.Top + Y
        
        mshDepost.Height = mshDepost.Height + Y
        picLR.Height = picLR.Height + Y
        mshInsure.Height = mshInsure.Height + Y
                
        picDetail.Top = picDetail.Top + Y
        vsfFee.Top = vsfFee.Top + Y
        vsfFee.Height = vsfFee.Height - Y
        
        Refresh
        msngScale = vsfFee.Height / (Me.ScaleHeight _
            - tabClass.Height - pic费用信息.Height - lblDepost.Height - picDetail.Height - fraUD.Height - 45)
    End If
End Sub

Private Sub lblMoney_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim objPopup As CommandBarPopup
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_DetailType, True, True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
'
'
'Private Sub SetActiveList(objAct As Object)
'    If objAct Is mshDepost Then
'        mshDepost.BackColorSel = &H800000
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H808080
'    ElseIf objAct Is mshInsure Then
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H800000
'        vsfFee.BackColorSel = &H808080
'    ElseIf objAct Is vsfFee Then
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H800000
'    Else
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H808080
'    End If
'    'Call mshDepost_EnterCell
'    Call mshInsure_EnterCell
'End Sub

Private Sub SetVsGrindSelColor(ByVal objGrid As Object, Optional blnLostFocus As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的相关颜色
    '入参:objGrid-相关网格控件
    '     blnLostFocus-光标移出
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-28 18:08:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    objGrid.BackColorSel = IIf(blnLostFocus, &H808080, &H800000)       ' &H8000000F
End Sub


Private Function ReadInsureMoney(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mshInsure.Redraw = flexRDNone
    mshInsure.Clear
    On Error GoTo errH
        
    strSQL = "Select 结算方式,To_Char(金额,'9999999990.00') as 结算金额" & _
        " From 保险模拟结算 Where 病人ID=[1] And 主页ID=[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng病人ID, lng主页ID)
    
    mshInsure.Clear
    Set mshInsure.DataSource = rsTmp
    If rsTmp.RecordCount = 0 Then mshInsure.Rows = 2
    
'    Call Grid.BandRec(mshInsure, rsTmp)
'    Call SetGridWidth(mshInsure, Me)        '如果取了,由于没有设置初始列宽,打印会异常
   With mshInsure
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = IIf(i > 0, 7, 1)
        Next
        .Row = 1: .Col = 0
        .Redraw = flexRDBuffered
        '恢复个性化设置
        zl_vsGrid_Para_Restore mlngModul, mshInsure, Me.Name, "mshInsure", False
   End With
    ReadInsureMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    mshInsure.Redraw = flexRDBuffered
    Call SaveErrLog
End Function


Function CurrencyToStr(ByVal Number As Currency) As String
    Dim str1Ary As Variant, str2Ary As Variant
    Dim a As Long, B As Long  '循环基数
    Dim tmp1 As String        '临时转换
    Dim tmp2 As String        '临时转换结果
    Dim Point As Long         '小数点位置
    
    Number = Val(Trim(Number))
    If Number = 0 Then CurrencyToStr = "": Exit Function
    If Number <= -922337203685477# Or Number >= 922337203685477# Then
       Exit Function
    End If
    
    str1Ary = Split("零 壹 贰 叁 肆 伍 陆 柒 捌 玖")
    str2Ary = Split("分 角 元 拾 佰 仟 万 拾 佰 仟 亿 拾 佰 仟 万 拾 佰")
    tmp1 = FormatEx(Number, 2)
    tmp1 = Replace(tmp1, "-", "")  '先去掉“-”号
    Point = InStr(tmp1, ".")       '取得小数点位置
    If Point = 0 Then      '如果有小数点，最大佰万亿
       B = Len(tmp1) + 2   '加2位小数
    Else
       B = Len(Left(tmp1, Point + 1))  '包括点加2位小数
    End If
    ''先将所有数字替换为中文
    For a = 9 To 0 Step -1
        tmp1 = Replace(Replace(tmp1, a, str1Ary(a)), ".", "")
    Next
    For a = 1 To B
        B = B - 1
        If Mid(tmp1, a, 1) <> "" Then
           If B > UBound(str2Ary) Then Exit For
           tmp2 = tmp2 & Mid(tmp1, a, 1) & str2Ary(B)
        End If
    Next
    If tmp2 = "" Then CurrencyToStr = "": Exit Function
    
'    ''〓下面为非正式财务算法，可以去掉〓
'    For a = 1 To Len(tmp2)
'        tmp2 = Replace(tmp2, "零亿", "亿零")
'        tmp2 = Replace(tmp2, "零万", "万零")
'        tmp2 = Replace(tmp2, "零仟", "零")
'        tmp2 = Replace(tmp2, "零佰", "零")
'        tmp2 = Replace(tmp2, "零拾", "零")
'        tmp2 = Replace(tmp2, "零元", "元")
'        tmp2 = Replace(tmp2, "零零", "零")
'        tmp2 = Replace(tmp2, "亿万", "亿")
'    Next
'    ''〓上面为非正式财务算法，可以去掉〓
    
    If Point = 1 Then tmp2 = "零元" + tmp2
    If Number < 0 Then tmp2 = "负" + tmp2
    If Point = 0 Then tmp2 = tmp2 + "整"
    CurrencyToStr = tmp2
End Function

Private Sub vsfFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
End Sub

Private Sub vsfFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "列头信息-" & mbytList, False
End Sub

Private Sub vsfFee_GotFocus()
    Call SetVsGrindSelColor(vsfFee)
End Sub
Private Sub vsfFee_LostFocus()
    Call SetVsGrindSelColor(vsfFee, True)
End Sub

Private Sub vsfFee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If mblnUnBilling Then Call ExecUnBilling
    End If
End Sub



Private Sub ChangeList(blnRefreshData As Boolean)
    Dim strTmp As String
    Dim objControl As CommandBarControl
    '31159
    Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
    If objControl Is Nothing Then Exit Sub
        
    strTmp = objControl.Caption
    If tabClass.SelectedItem.Index = conTab未结 Then
        lblMoney.Caption = " 未结" & Left(strTmp, Len(strTmp) - 4) & "▼"
    Else
        lblMoney.Caption = " 结帐" & Left(strTmp, Len(strTmp) - 4) & "▼"
    End If
    
    If mbytList = ListType.C1分科室明细 Or mbytList = ListType.C4分类分项明细 Then
        fraDeptMode.Visible = mbytList = ListType.C1分科室明细
        fraTypeMode.Visible = mbytList = ListType.C4分类分项明细
        
        If mbytList = ListType.C1分科室明细 Then
            optDeptMode(0).Caption = "开单科室"
            optDeptMode(1).Caption = "执行科室"
        Else
            optTypeMode(0).Caption = "收入项目"
            optTypeMode(1).Caption = "收据费目"
        End If
    Else
        fraTypeMode.Visible = False
        fraDeptMode.Visible = False
    End If
    
    cboFeeType.Visible = (mbytList = ListType.C0费用清单 Or mbytList = ListType.C2分项目明细)
    cboDept.Visible = (mbytList <> ListType.C1分科室明细)
    
    If cboFeeType.ListCount = 0 Then cboFeeType.AddItem "所有费目"
    
    If Visible Then
        Call SetCondition
        If blnRefreshData Then Call LoadCardData(False, False, True)
    End If
    
End Sub

Private Sub SetCondition()
    Dim lngLeft As Long
    Dim sngTop As Single
    
    picDetail.AutoRedraw = True
    sngTop = lblMoney.Top + (lblMoney.Height - chkAdivce.Height) \ 2
    chkAdivce.Left = lblMoney.Left + lblMoney.Width + 150
    chkAdivce.Top = sngTop
    sngTop = lblMoney.Top + (lblMoney.Height - cboFeeType.Height) \ 2
    lngLeft = chkAdivce.Left + chkAdivce.Width + 50
    If cboFeeType.Visible Then
        cboFeeType.Top = sngTop
        cboFeeType.Left = lngLeft
        lngLeft = cboFeeType.Left + cboFeeType.Width + 50
    ElseIf fraTypeMode.Visible Then
        fraTypeMode.Top = cboFeeType.Top + (cboFeeType.Height - fraTypeMode.Height) \ 2
        fraTypeMode.Left = lngLeft
        lngLeft = fraTypeMode.Left + fraTypeMode.Width + 50
        
        fraTypeMode.Width = optTypeMode(1).Width + optTypeMode(0).Width + 100
        optTypeMode(1).Left = optTypeMode(0).Left + optTypeMode(0).Width + 50
        
        
    ElseIf fraDeptMode.Visible Then
        fraDeptMode.Top = cboFeeType.Top + (cboFeeType.Height - fraDeptMode.Height) \ 2
        fraDeptMode.Left = lngLeft
        fraDeptMode.Width = optDeptMode(1).Width + optDeptMode(0).Width + 100
        optDeptMode(1).Left = optDeptMode(0).Left + optDeptMode(0).Width + 50
        
        lngLeft = fraDeptMode.Left + fraDeptMode.Width + 50
    Else
        lngLeft = chkAdivce.Left + chkAdivce.Width + 50
    End If
    
    If cboDept.Visible Then
        cboDept.Left = lngLeft
        lngLeft = cboDept.Left + cboDept.Width + 50
    End If
    
    tabTime.Top = cboFeeType.Top + (cboFeeType.Height - tabTime.Height) \ 2
    If cboBaby.Visible Then
        cboBaby.Left = lngLeft + 50
        tabTime.Left = cboBaby.Left + cboBaby.Width + 50
    Else
        tabTime.Left = lngLeft + 50
    End If
    sngTop = cboFeeType.Top + cboFeeType.Height + 50
    dtpBegin.Height = cbo日期.Height: dtpEnd.Height = cbo日期.Height
    dtpBegin.Top = sngTop: dtpEnd.Top = sngTop: cbo日期.Top = sngTop
    cmdRefresh.Top = sngTop
    lbl发生时间.Top = sngTop + (cboFeeType.Height - lblMoney.Height) \ 2
    lbl至.Top = lbl发生时间.Top: lbl日期范围.Top = lbl发生时间.Top
    chk仅显示销帐单据.Top = sngTop + (cboFeeType.Height - chk仅显示销帐单据.Height) \ 2
    lngLeft = lbl发生时间.Left + lbl发生时间.Width + 50
    cbo日期.Left = lngLeft
    lngLeft = cbo日期.Left + cbo日期.Width + 50
    lbl日期范围.Left = lngLeft: dtpBegin.Left = lngLeft
    lbl至.Left = dtpBegin.Width + dtpBegin.Left + 50
    dtpEnd.Left = lbl至.Left + lbl至.Width + 50
    cmdRefresh.Left = dtpEnd.Left + dtpEnd.Width + 50
    picDetail.AutoRedraw = False
End Sub

Private Sub mshDepost_GotFocus()
    Call SetVsGrindSelColor(mshDepost)
End Sub
Private Sub mshDepost_LostFocus()
    Call SetVsGrindSelColor(mshDepost, True)
End Sub


Private Sub mshInsure_GotFocus()
    Call SetVsGrindSelColor(mshInsure)
End Sub
Private Sub mshInsure_LostFocus()
    Call SetVsGrindSelColor(mshInsure, True)
End Sub

Private Sub picDetail_Resize()
    Dim sngLeft As Single
    
    On Error Resume Next
    Call SetCondition
    tabTime.Width = picDetail.ScaleWidth - tabTime.Left - imgColSel.Width - 450
    With imgColSel
        .Left = picDetail.ScaleWidth - .Width - 200
    End With
    With chk按规格进行小计
        .Left = imgColSel.Left - .Width - 200
        sngLeft = imgColSel.Left - chk仅显示销帐单据.Width - 200
        If .Visible Then sngLeft = .Left
    End With
    With chk仅显示销帐单据
        .Left = sngLeft - .Width - 200
        If .Visible Then sngLeft = .Left
    End With
    '46646
    With chkNotCheckFee
        .Top = chk仅显示销帐单据.Top
        .Left = sngLeft - .Width - 200
    End With
End Sub

Private Sub picLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshInsure.Width - X < 2000 Or mshDepost.Width + X < 2000 Then Exit Sub
        picLR.Left = picLR.Left + X
        lblDepost.Width = lblDepost.Width + X
        mshDepost.Width = mshDepost.Width + X
        lblInsure.Left = lblInsure.Left + X
        lblInsure.Width = lblInsure.Width - X
        mshInsure.Left = mshInsure.Left + X
        mshInsure.Width = mshInsure.Width - X
        
        Refresh
    End If
End Sub



Private Sub vsfFee_DblClick()
    Dim lngColTmp As Long, strNO As String, byt记录性质 As Byte, byt记录状态 As Byte
    If vsfFee.MouseRow = 0 Then Exit Sub
    
    With vsfFee
        If .Row > 0 Then
            lngColTmp = .ColIndex("单据号")
            If lngColTmp <> -1 Then
                strNO = .TextMatrix(.Row, lngColTmp)
                
                lngColTmp = .ColIndex("记录性质")
                If lngColTmp <> -1 Then
                    byt记录性质 = Val(.TextMatrix(.Row, lngColTmp))
                End If
                
                If strNO <> "" And byt记录性质 <> 0 Then
                    lngColTmp = .ColIndex("记录状态")
                    If lngColTmp <> -1 Then
                        byt记录状态 = Val(.TextMatrix(.Row, lngColTmp))
                    End If
                    Call ShowBilling(strNO, byt记录性质, byt记录状态)
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowBilling(ByVal strNO As String, ByVal byt记录性质 As Byte, ByVal byt记录状态 As Byte)
    Dim blnNOMoved As Boolean
    
    If Get费用来源(strNO) = 1 Then
        Call ZLShowChargeWindow(Me, 2, 1, 0, 0, 0, 0, False, 0, "", strNO)
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    gbytBilling = 0 '记帐查阅
    blnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, , 2, Caption)
    
    If BillisBatch(strNO) Then '批量记帐
        frmBillings.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnDelete = byt记录状态 = 2
        frmBillings.mblnNOMoved = blnNOMoved
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnDelete = byt记录状态 = 2
        frmSimpleBilling.mblnNOMoved = blnNOMoved
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        frmCharge.mbytNOType = byt记录性质
        frmCharge.mblnDelete = byt记录状态 = 2
        frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmCharge.mbytInState = 1
        frmCharge.mstrInNO = strNO
        frmCharge.mlngModule = mlngModul
        frmCharge.mblnNOMoved = blnNOMoved
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Function Get费用来源(ByVal strNO As String) As Byte
    '根据单据号判断费用记录来源
    '返回：0-住院,1-门诊
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select 1 From 门诊费用记录 Where 记录性质=2 And NO=[1] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Function
    Get费用来源 = 1
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
    
Private Sub ExecUnBilling()
    Dim strNO As String, strTime As String, blnNOMoved As Boolean, lng病区ID As Long, strUnitIDs As String
    Dim blnBat As Boolean, intTmp As Integer, i As Long, intInsure As Integer, lngDelRow As Long
    Dim str病人IDs As String, strInfo As String, strUser As String
    Dim strInsure As String, arrInsure As Variant, bytType As Byte, blnFlagPrint As Boolean
    Dim byt费用来源 As Byte '0-住院,1-门诊
    
    If Not (mbytList = ListType.C0费用清单 _
            Or mbytList = ListType.C1分科室明细 _
            Or mbytList = ListType.C2分项目明细 _
            Or mbytList = ListType.C3分类别明细 _
            Or mbytList = ListType.C4分类分项明细) Then Exit Sub
        
    If InStr(GetInsidePrivs(Enum_Inside_Program.p住院记帐), "所有病区") = 0 Then
        If InStr("," & mstrUnitIDs & ",", "," & mlng病区ID & ",") = 0 Then
            MsgBox "你没有所有病区的权限，不能对其它病区的病人销帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    With vsfFee
        i = .ColIndex("单据号")
        If i = -1 Then Exit Sub
        strNO = .TextMatrix(.Row, i)
        If strNO = "" Then Exit Sub
        
        strTime = .TextMatrix(.Row, .ColIndex("登记时间"))
        bytType = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        strUser = .TextMatrix(.Row, .ColIndex("操作员"))
        lngDelRow = Val(.TextMatrix(.Row, .ColIndex("序号")))
    End With
    byt费用来源 = Get费用来源(strNO)
    
    '权限判断
    If Not BillOperCheck(IIf(byt费用来源 = 1, 4, 5), strUser, CDate(strTime), "销帐", strNO, , bytType) Then Exit Sub
        
    '是否已转入后备数据表中
    If zlDatabase.NOMoved(IIf(byt费用来源 = 1, "门诊费用记录", "住院费用记录"), strNO, , CStr(bytType), Caption) Then
        If Not ReturnMovedExes(strNO, bytType, Caption) Then Exit Sub
    End If
            
    '项目冲销权限
    If Not CheckDelPriv(strNO, GetInsidePrivs(Enum_Inside_Program.p记帐操作), strTime, bytType, , byt费用来源) Then Exit Sub
        
    '留观病人权限
    strInfo = Check留观病人(strNO, GetInsidePrivs(Enum_Inside_Program.p记帐操作), strTime, bytType, byt费用来源)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '是否已执行
    If byt费用来源 = 0 Then blnBat = BillisBatch(strNO)
    i = BillCanDelete(strNO, bytType, blnBat, strTime, GetInsidePrivs(Enum_Inside_Program.p记帐操作), blnFlagPrint, byt费用来源)
    If i <> 0 Then
        Select Case i
            Case 1 '该单据不存在
                MsgBox "指定单据中的内容不存在,或者你没有相关收费项目的销帐权限！", vbInformation, gstrSysName
            Case 2 '已经全部完全执行
                MsgBox "指定单据中的内容已经全部完全执行！", vbInformation, gstrSysName
            Case 3 '未完全执行部分剩余数量为0
                MsgBox "指定单据中的内容未完全执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '出院病人操作权限判断
    If Not BillCanBeOperate(strNO, GetInsidePrivs(Enum_Inside_Program.p记帐操作), "销帐", _
        strTime, str病人IDs, bytType, byt费用来源) Then Exit Sub
    
    '是否已经结帐
    intTmp = HaveBilling(IIf(byt费用来源 = 1, 1, 2), strNO, False, strTime, bytType)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType, byt费用来源)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                        '医保病人的单据,固定为已结帐的禁止销帐
                        If intTmp = 1 Then
                            MsgBox "该医保记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "该医保记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("该记帐单据包含已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "该记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                                Exit Sub
                            Else
                                MsgBox "该记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName
                            End If
                    End Select
                End If
            Next
        End If
    End If
    
    intInsure = BillExistInsure(strNO, , , bytType, byt费用来源) '判断是否含有医保病人记的帐,记帐表检查其中只要有医保病人
    '医保销帐不允许对负数记录进行销帐
    If intInsure <> 0 Then
        If CheckNONegative(strNO, bytType, byt费用来源) Then
            MsgBox "该单据存在负数记帐记录,不允许进行医保销帐操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    '是否存在重算冲减记录
    If CheckRecalcRecord(strNO, byt费用来源) Then
        MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
            "结帐前请按费别重算费用，否则病人将享受已销帐单据的打折优惠金额！", vbInformation, Caption
    End If
     
    If byt费用来源 = 1 Then
        Call ZLShowChargeWindow(Me, 2, 3, 0, 0, 0, 0, False, 0, "", strNO)
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
        
    gbytBilling = 0
    If blnBat Then '批量记帐
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmBillings.mbytInState = 3
        frmBillings.mstrInNO = strNO
        frmBillings.mlngDelRow = lngDelRow
        frmBillings.mstrTime = strTime
        frmBillings.mstr病人IDs = str病人IDs
        frmBillings.mlngUnitID = 0
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO, bytType) Then '简单记帐
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngUnitID = 0
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        
        frmCharge.mbytUseType = 1
        frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNO
        frmCharge.mlngDelRow = lngDelRow
        frmCharge.mbytNOType = bytType
        frmCharge.mstrTime = strTime
        frmCharge.mlngUnitID = 0
        frmCharge.mlngModule = mlngModul
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If

    If gblnOK Then Call RefreshAllData
End Sub
 
Private Sub ExecPrintDailyDetail()
    frmDailyListAsk.mlngModul = 1141    '仍然以一日清单模块的参数为准
    frmDailyListAsk.mbytInFun = 1
    frmDailyListAsk.mlng病人ID = mlng病人ID
    frmDailyListAsk.Show vbModal, Me
    If frmDailyListAsk.mblnAskOk Then
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "病人ID=" & mlng病人ID, _
            "开始时间=" & Format(frmDailyListAsk.mdatBegin, "YYYY-MM-DD HH:MM:SS"), _
            "结束时间=" & Format(frmDailyListAsk.mdatEnd, "YYYY-MM-DD HH:MM:SS"), _
            "显示退费=" & IIf(mvs.ReBalance, "1", "0"), _
            "显示零费用=" & IIf(mvs.ZeroFee, "1", "0"), _
            "病人病区=" & mlng病区ID, _
            "主页ID=" & frmDailyListAsk.mlngPageID, _
            "费用时间=" & IIf(mbytDateType, "发生时间", "登记时间"), 1
    End If
End Sub

Private Sub optDeptMode_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub

Private Sub optTypeMode_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub


Private Sub tabClass_Click()
    Dim strTmp As String
    Dim objControl As CommandBarControl
    Dim intPreSel As Integer
    
    If tabClass.SelectedItem.Index = mintPreCard Then Exit Sub
    
    mintPreCard = tabClass.SelectedItem.Index
    Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
    If objControl Is Nothing Then Exit Sub
    strTmp = objControl.Caption
    If tabClass.SelectedItem.Index = conTab未结 Then
        mblnPreBalance = mintInsure <> 0   '59073
        lblMoney.Caption = " 未结" & Left(strTmp, Len(strTmp) - 4) & "▼"
    Else
        lblMoney.Caption = " 结帐" & Left(strTmp, Len(strTmp) - 4) & "▼"
    End If
    intPreSel = mintPreTimeIndex
    If Not tabTime.SelectedItem Is Nothing And mintPreTimeIndex = 0 Then
        intPreSel = tabTime.SelectedItem.Index
    End If
    
    If LoadPatiTime Then
        '43494
        '问题号:53136 修改人:刘兴洪,修改时间:2012-12-10 13:26:07
        If tabTime.SelectedItem Is Nothing Or tabClass.Tag = "Loaded" Then
            If tabTime.Tabs.Count >= intPreSel And intPreSel <> 0 Then
                tabTime.Tabs(intPreSel).Selected = True    '调用tabTime_Click
            ElseIf tabTime.Tabs.Count <> 0 Then
                tabTime.Tabs((1)).Selected = True   '调用tabTime_Click
            End If
        Else
            Call tabTime_Click
        End If
    End If
    tabClass.Tag = "Loaded"
End Sub

Private Sub tabTime_Click()
    '问题号:53136 修改人:刘兴洪,修改时间:2012-12-10 11:55:54,含mintPreTimeIndex
    If tabTime.Tag = "1" Then Exit Sub
    If tabTime.SelectedItem.Index = mintPreTime Then Exit Sub
    mintPreTime = tabTime.SelectedItem.Index
    '记录上次最后选择
    mintPreTimeIndex = IIf(tabTime.Tabs.Count > 1, mintPreTime, mintPreTimeIndex)
    
    '显示当前卡片数据
    Call LoadCardData
End Sub

Private Sub vsfFee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim objPopup As CommandBarPopup
        If Not Me.ActiveControl Is vsfFee Then vsfFee.SetFocus
        If vsfFee.MouseRow <> vsfFee.Row Then vsfFee.Row = vsfFee.MouseRow
    
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup, True, False)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
Private Function zlCopyBill(ByVal int场合 As Integer, ByVal frmMain As Object, ByVal lng病区ID As Long, _
    ByVal lng病人ID As Long, bln出院 As Boolean, ByVal bln结清 As Boolean, _
    Optional strUnitIDs As String = "", Optional lng主页ID As Long = 0, _
    Optional lng科室ID As Long = 0, Optional ByVal bln门诊留观病人 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拷贝单据
    '入参:int场合- 0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS);9-费用查询调用
    '编制:刘兴洪
    '日期:2013-02-17 10:45:31
    '问题:54274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String, intIndex As Integer, strNO As String
    Dim lng部门ID As Long, str最后转科时间 As String
    '先看有记帐权限没有
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    If InStr(1, strPrivs, ";住院记帐;") = 0 Then Exit Function
    
    If InStr(GetInsidePrivs(Enum_Inside_Program.p住院记帐), "所有病区") = 0 Then
        If strUnitIDs = "" Then
            '重新获取操作员的所在病区
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng病区ID & ",") = 0 Then
            MsgBox "你没有所有病区的权限，不能对其它病区的病人记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    lng部门ID = 0
    If mbln补费 Then
        If InStr(1, "012", int场合) > 0 Then '    int场合 = 0 - 医生站调用, 1 - 护士站调用, 2 - 医技站调用(PACS / LIS))
            lng部门ID = lng科室ID
        End If
        '补费检查是否超过时限
        If zlCheckPatiFeeRenewValied(lng病人ID, lng主页ID, lng病区ID, lng科室ID, str最后转科时间) = False Then Exit Function
    Else
        lng部门ID = lng科室ID
    End If
    
    With vsfFee
        intIndex = .ColIndex("单据号")
        If intIndex = -1 Then Exit Function
        strNO = .TextMatrix(.Row, intIndex)
        If strNO = "" Then Exit Function
    End With
    '多病人单据,不能拷贝
    If BillisBatch(strNO) Then
        MsgBox "单据『" & strNO & "』是记帐表,不允许复制该记帐单据!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    '出院病人记帐权限
    If bln出院 Then
        If bln结清 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "出院结清强制记帐") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf Not bln结清 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "出院未结强制记帐") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '门诊留观病人走门诊记帐模式
    If bln门诊留观病人 Then
        If Not (gbln门诊留观 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), ";门诊留观记帐;") > 0) Then
            MsgBox "你没有权限对门诊留观病人进行记帐操作！", vbInformation, gstrSysName
            Exit Function
        End If
        zlCopyBill = ZLShowChargeWindow(Me, 2, 11, lng病人ID, lng主页ID, _
            lng部门ID, lng病区ID, False, 0, "", strNO)
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    gblnOK = False
    gbytBilling = 0
    frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
    frmCharge.mbytUseType = 1
    frmCharge.mbytInState = 0
    frmCharge.mblnCopyBill = True
    frmCharge.mstrInNO = strNO
    frmCharge.mlngDeptID = lng部门ID
    frmCharge.mlngUnitID = lng病区ID
    frmCharge.mlngModule = mlngModul
    frmCharge.mlng病人ID = lng病人ID
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If Not gblnOK Then Exit Function
    zlCopyBill = True
End Function

Private Function LoadPages(ByVal intPage As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载第几页
    '入参:intPage-页数
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-03 10:26:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, cllTemp As Collection
    Dim strKey As String
    On Error GoTo errHandle
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    Set cllTemp = mcllBalaceNums(intPage)
    For i = 1 To cllTemp.Count
        If i = 1 Then strKey = "_" & cllTemp(i)(1)
        tabClass.Tabs.Add , "_" & cllTemp(i)(1), cllTemp(i)(2)
        tabClass.Tabs(tabClass.Tabs.Count).Tag = Val(cllTemp(i)(0)) '记录结帐ID,加快速度
        tabClass.Tabs(tabClass.Tabs.Count).ToolTipText = cllTemp(i)(3)
    Next
    If cllTemp.Count > 0 Then tabClass.Tabs(strKey).Selected = True
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    End If
    LoadPages = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFormOperation() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取窗体操作选择，窗体卸载前调用
    '返回:
    '     上次窗体操作条件字符串
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strKey As String

    On Error GoTo errHandle
    
    If mlng病人ID = 0 Then Exit Function

    'zlGetFormOperation字符串格式:
    '病人ID|费用类型(tabClass.SelectedItem.Key)|清单查询类型序号|住院次数(tabTime.SelectedItem.Key)|cboNum.Text|chkAdivce,chkAdivce.Value| _
    chk不包含体检费用,chk不包含体检费用.Value|cbo日期,cbo日期.cbo日期.Text,dtpBegin.Value,dtpEnd.Value| _
    chk仅显示销帐单据,chk仅显示销帐单据.Value|chk按规格进行小计,chk按规格进行小计.Value| _
    cboFeeType,cboFeeType.Text|cboDept,cboDept.Text|cboBaby,cboBaby.Text|optDeptMode,optDeptMode(0).Value| _
    optTypeMode,optTypeMode(0).Value|mshDepost,单据号,结算方式|mshInsure,结算方式|vsfFee,发生日期,单据号,项目编码,记录状态...|cboNum,cboNum.Text

    strKey = mlng病人ID & "|" & tabClass.SelectedItem.Key & "|" & mbytList & "|" & tabTime.SelectedItem.Key & "|" & cboNum.Text
    strKey = strKey & "|chkAdivce," & IIf(chkAdivce.Value = 1, 1, 0) & "|chkNotCheckFee," & IIf(chkNotCheckFee.Value = 1, 1, 0)
    strKey = strKey & "|cbo日期," & cbo日期.Text & "," & Format(dtpBegin.Value, "yyyy-mm-dd hh:MM:ss") & "," & Format(dtpEnd.Value, "yyyy-mm-dd hh:MM:ss")
    strKey = strKey & "|chk仅显示销帐单据," & IIf(chk仅显示销帐单据.Value = 1, 1, 0)
    strKey = strKey & "|chk按规格进行小计," & IIf(chk按规格进行小计.Value = 1, 1, 0)
    strKey = strKey & "|cboFeeType," & cboFeeType.Text
    strKey = strKey & "|cboDept," & cboDept.Text
    strKey = strKey & "|cboBaby," & cboBaby.Text
    strKey = strKey & "|optDeptMode," & IIf(optDeptMode(0).Value, 0, 1)
    strKey = strKey & "|optTypeMode," & IIf(optTypeMode(0).Value, 0, 1)

    With mshDepost
        If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
            strKey = strKey & "|mshDepost," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                     .TextMatrix(.RowSel, .ColIndex("结算方式"))
        End If
    End With
    
    With mshInsure
        If mintInsure <> 0 Then
            If .Rows > 1 And .TextMatrix(1, .ColIndex("结算方式")) <> "" Then
                strKey = strKey & "|mshInsure," & .TextMatrix(.RowSel, .ColIndex("结算方式"))
            End If
        End If
    End With
    
    With vsfFee
        Select Case mbytList
            Case ListType.C0费用清单
                If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel, .ColIndex("记录状态"))
                End If
            Case ListType.C1分科室明细
                If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("开单科室")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel, .ColIndex("记录状态"))
                End If
                
            Case ListType.C2分项目明细
                If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("项目")) <> "" Then   '汇总行
                        If .TextMatrix(.RowSel, .ColIndex("项目")) = "总计" Then
                            strKey = strKey & "|vsfFee,,,,,总计"
                        Else
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("单据号")) & "," & _
                            .TextMatrix(.RowSel - 1, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("记录状态")) & ",汇总"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                        .TextMatrix(.RowSel, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel, .ColIndex("记录状态"))
                    End If
                End If
                
            Case ListType.C3分类别明细
                If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("费目")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel, .ColIndex("记录状态"))
                End If
                
            Case ListType.C4分类分项明细
                If .Rows > 1 And .TextMatrix(1, .ColIndex("单据号")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("收入项目")) & "," & .TextMatrix(.RowSel, .ColIndex("收费项目")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("项目编码")) & "," & .TextMatrix(.RowSel, .ColIndex("记录状态"))
                End If
                
            Case ListType.C5分项目汇总
                If .Rows > 1 And .TextMatrix(1, .ColIndex("收费类别")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("收费类别")) & "," & .TextMatrix(.RowSel, .ColIndex("项目")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("规格"))
                End If
            Case ListType.C6分类别汇总
                If .Rows > 1 And .TextMatrix(1, .ColIndex("费目")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("费目"))
                End If
            Case ListType.C7分月分类汇总
                If .Rows > 1 And .TextMatrix(1, .ColIndex("期间")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("费目")) = "" Then   '汇总行
                        If .TextMatrix(.RowSel, .ColIndex("期间")) <> "总计" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("期间")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("费目")) & ",汇总"
                        Else
                            strKey = strKey & "|vsfFee,,,总计"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("期间")) & "," & .TextMatrix(.RowSel, .ColIndex("费目")) & ","
                    End If
                End If
            Case ListType.C8逐日单据汇总
                If .Rows > 1 And .TextMatrix(1, .ColIndex("发生日期")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("费用项目")) = "" Then   '汇总行
                        If .TextMatrix(.RowSel, .ColIndex("发生日期")) <> "总计" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("单据号")) & "," & _
                            .TextMatrix(.RowSel - 1, .ColIndex("费用项目")) & ",汇总"
                        Else
                            strKey = strKey & "|vsfFee,,,,总计"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("单据号")) & "," & _
                        .TextMatrix(.RowSel, .ColIndex("费用项目")) & ","
                    End If
                End If
            Case ListType.C9逐日费目汇总
                If .Rows > 1 And .TextMatrix(1, .ColIndex("发生日期")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("费用项目")) = "" Then   '汇总行
                        If .TextMatrix(.RowSel, .ColIndex("发生日期")) <> "总计" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("费用项目")) & ",汇总"
                        Else
                            strKey = strKey & "|vsfFee,,,总计"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("发生日期")) & "," & .TextMatrix(.RowSel, .ColIndex("费用项目")) & ","
                    End If
                End If
        End Select
    End With
    strKey = strKey & "|cboNum," & cboNum.Text
    
    zlGetFormOperation = strKey
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlRestoreFormOperation(ByVal strValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复窗体操作选择，窗体刷新前调用
    '入参:
    '     strValue-上次窗体操作条件字符串
    '返回:
    '     True-窗体恢复成功;False-窗体恢复失败
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Integer
    On Error GoTo errHandle
    
    If strValue = "" Then Exit Function
    mstrRestoreFeeCons = strValue
    varData = Split(strValue, "|")
    mbytList = varData(2)
    zlRestoreFormOperation = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlRestoreFeeControls(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复费用查询页签各个控件的操作选择
    '入参:
    '     lng病人ID-病人ID
    '返回:
    '     True-恢复成功;False-恢复失败
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Relng病人ID As Long
    Dim varData As Variant, i As Integer, j As Integer
    On Error GoTo errHandle
    If mstrRestoreFeeCons = "" Then Exit Function
    varData = Split(mstrRestoreFeeCons, "|")
    Relng病人ID = Val(varData(0))
    If Relng病人ID <> lng病人ID Then Exit Function
    
    For i = 1 To tabTime.Tabs.Count
        If tabTime.Tabs(i).Key = varData(3) Then tabTime.Tabs(i).Selected = True
    Next
    mblnNotClick = True
    For i = 5 To UBound(varData)
        Select Case Split(varData(i), ",")(0)
            Case "cbo日期"
                cbo日期.ListIndex = -1
                    For j = 0 To cbo日期.ListCount - 1
                        If cbo日期.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                            cbo日期.ListIndex = j
                            Exit For
                        End If
                        cbo日期.ListIndex = 0
                    Next
                If cbo日期.ListIndex = 5 Then
                    dtpBegin.Value = Format(Split(varData(i), ",")(2), "yyyy-mm-dd hh:MM:ss")
                    dtpEnd.Value = Format(Split(varData(i), ",")(3), "yyyy-mm-dd hh:MM:ss")
                End If
                Call SetDateVisible
            Case "chkAdivce"
                chkAdivce.Value = Val(Split(varData(i), ",")(1))
                
            Case "chkNotCheckFee"
                chkNotCheckFee.Value = Val(Split(varData(i), ",")(1))
                
            Case "chk仅显示销帐单据"
                If chk仅显示销帐单据.Visible And chk仅显示销帐单据.Enabled Then
                    chk仅显示销帐单据.Value = Val(Split(varData(i), ",")(1))
                End If
                
            Case "chk按规格进行小计"
                If chk按规格进行小计.Visible And chk按规格进行小计.Enabled Then
                    chk按规格进行小计.Value = Val(Split(varData(i), ",")(1))
                End If
                
            Case "cboFeeType"
                cboFeeType.ListIndex = -1
                For j = 0 To cboFeeType.ListCount - 1
                    If cboFeeType.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboFeeType.ListIndex = j
                        Exit For
                    End If
                    cboFeeType.ListIndex = 0
                Next
                
            Case "cboDept"
                cboDept.ListIndex = -1
                For j = 0 To cboDept.ListCount - 1
                    If cboDept.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboDept.ListIndex = j
                        Exit For
                    End If
                    cboDept.ListIndex = 0
                Next
                
            Case "cboBaby"
                cboBaby.ListIndex = -1
                For j = 0 To cboBaby.ListCount - 1
                    If cboBaby.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboBaby.ListIndex = j
                        Exit For
                    End If
                    cboBaby.ListIndex = 0
                Next
                
            Case "optDeptMode"
                If optDeptMode(0).Value And optDeptMode(0).Enabled Then
                    optDeptMode(Val(Split(varData(i), ",")(1))).Value = True
                End If
                
            Case "optTypeMode"
                If optTypeMode(0).Value And optTypeMode(0).Enabled Then
                    optTypeMode(Val(Split(varData(i), ",")(1))).Value = True
                End If

        End Select
    Next
    mblnNotClick = False
    zlRestoreFeeControls = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlRestorePosition(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复费用查询页签列表控件的选中行
    '入参:
    '     lng病人ID-病人ID
    '返回:
    '     True-恢复成功;False-恢复失败
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Relng病人ID As Long
    Dim varData As Variant, i As Integer, j As Integer
    On Error GoTo errHandle
    If mstrRestoreFeeCons = "" Then Exit Function
    varData = Split(mstrRestoreFeeCons, "|")
    Relng病人ID = Val(varData(0))
    If Relng病人ID <> lng病人ID Then Exit Function
    
    For i = 5 To UBound(varData)
        Select Case Split(varData(i), ",")(0)
            Case "mshDepost"
                With mshDepost
                    For j = 1 To .Rows - 1
                        If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("单据号")) And _
                           Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("结算方式")) Then
                            .Row = j
                            Exit For
                        End If
                    Next
                End With
            
            Case "mshInsure"
                With mshInsure
                    For j = 1 To .Rows - 1
                        If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("结算方式")) Then
                            .Row = j
                            Exit For
                        End If
                    Next
                End With
                        
            Case "vsfFee"
                With vsfFee
                    Select Case mbytList
                        Case ListType.C0费用清单
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("项目编码")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("记录状态")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                            
                        Case ListType.C1分科室明细
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("开单科室")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("项目编码")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("记录状态")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C2分项目明细
                            If Nvl(Split(varData(i), ",")(5)) = "总计" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                       Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("项目编码")) And _
                                       Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("记录状态")) Then
                                        If Nvl(Split(varData(i), ",")(5)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C3分类别明细
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("费目")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("项目编码")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("记录状态")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C4分类分项明细
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("收入项目")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("收费项目")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("项目编码")) And _
                                   Nvl(Split(varData(i), ",")(6)) = .TextMatrix(j, .ColIndex("记录状态")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C5分项目汇总
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("收费类别")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("项目")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("规格")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C6分类别汇总
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("费目")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C7分月分类汇总
                            If Nvl(Split(varData(i), ",")(3)) = "总计" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("期间")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("费目")) Then
                                        If Nvl(Split(varData(i), ",")(3)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C8逐日单据汇总
                            If Nvl(Split(varData(i), ",")(4)) = "总计" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("单据号")) And _
                                       Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("费用项目")) Then
                                        If Nvl(Split(varData(i), ",")(4)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C9逐日费目汇总
                            If Nvl(Split(varData(i), ",")(3)) = "总计" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("发生日期")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("费用项目")) Then
                                        If Nvl(Split(varData(i), ",")(3)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                    End Select
                End With
        End Select
    Next
    mstrRestoreFeeCons = ""
    zlRestorePosition = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


