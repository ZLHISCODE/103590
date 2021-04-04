VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmAdviceOperate 
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   Icon            =   "frmAdviceOperate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   12120
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox pictmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9960
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic疑问 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   12120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   12120
      Begin VB.TextBox txt疑问 
         Height          =   300
         Left            =   1275
         MaxLength       =   200
         TabIndex        =   1
         Top             =   15
         Width           =   9585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "校对疑问"
         Height          =   180
         Left            =   495
         TabIndex        =   17
         Top             =   75
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   150
         Picture         =   "frmAdviceOperate.frx":058A
         Top             =   45
         Width           =   240
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   714
      BandCount       =   1
      _CBWidth        =   12120
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   345
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "全选"
               Object.ToolTipText     =   "全选(Ctrl+A)"
               Object.Tag             =   "全选"
               ImageKey        =   "全选"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "全清"
               Object.ToolTipText     =   "全清(Ctrl+R)"
               Object.Tag             =   "全清"
               ImageKey        =   "全清"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "执行"
               Key             =   "执行"
               Description     =   "执行"
               Object.ToolTipText     =   "执行(Ctrl+E)"
               Object.Tag             =   "执行"
               ImageKey        =   "执行"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Description     =   "重置"
               Object.ToolTipText     =   "重新设置条件(F12)"
               Object.Tag             =   "重置"
               ImageKey        =   "重置"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "刷新"
               Description     =   "刷新"
               Object.ToolTipText     =   "重新读取数据(F5)"
               Object.Tag             =   "刷新"
               ImageKey        =   "刷新"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出(ALT+X)"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   405
      Width           =   12120
      Begin VB.Frame fraOper 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   825
         TabIndex        =   29
         Top             =   -30
         Width           =   8895
         Begin VB.ComboBox cboTime 
            Height          =   300
            Index           =   0
            ItemData        =   "frmAdviceOperate.frx":0B14
            Left            =   5085
            List            =   "frmAdviceOperate.frx":0B16
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   150
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Index           =   1
            ItemData        =   "frmAdviceOperate.frx":0B18
            Left            =   7830
            List            =   "frmAdviceOperate.frx":0B1A
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton optOper 
            Caption         =   "开始时间"
            Height          =   180
            Index           =   1
            Left            =   2330
            TabIndex        =   31
            Top             =   200
            Width           =   1050
         End
         Begin VB.OptionButton optOper 
            Caption         =   "当前时间"
            Height          =   180
            Index           =   0
            Left            =   1250
            TabIndex        =   30
            Top             =   200
            Width           =   1100
         End
         Begin VB.Label lblS 
            AutoSize        =   -1  'True
            Caption         =   "开始时间晚于开嘱时间"
            Height          =   180
            Left            =   3465
            TabIndex        =   36
            Top             =   195
            Width           =   1800
         End
         Begin VB.Label lblB 
            AutoSize        =   -1  'True
            Caption         =   "开始时间早于开嘱时间"
            Height          =   180
            Left            =   6225
            TabIndex        =   35
            Top             =   210
            Width           =   1800
         End
         Begin VB.Label lblOper 
            Caption         =   "校对时间(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   200
            Width           =   1000
         End
      End
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7320
         TabIndex        =   19
         Top             =   105
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "婴儿医嘱"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   22
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "所有医嘱"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "病人医嘱"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   20
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdAlley 
         Caption         =   "过敏史/病生状态"
         Height          =   350
         Left            =   10545
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Frame fraStop 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   0
         TabIndex        =   18
         Top             =   350
         Visible         =   0   'False
         Width           =   12105
         Begin VB.OptionButton optStop 
            Caption         =   "指定时间"
            Height          =   180
            Index           =   1
            Left            =   2655
            TabIndex        =   28
            Top             =   90
            Width           =   1110
         End
         Begin VB.OptionButton optStop 
            Caption         =   "上次执行时间"
            Height          =   180
            Index           =   0
            Left            =   1140
            TabIndex        =   27
            Top             =   90
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.CheckBox chkRollSend 
            Caption         =   "收回超期的"
            Height          =   195
            Left            =   7275
            TabIndex        =   26
            Top             =   90
            Width           =   1200
         End
         Begin VB.CheckBox chkNoSend 
            Caption         =   "补发未发的"
            Height          =   195
            Left            =   6015
            TabIndex        =   25
            Top             =   90
            Width           =   1230
         End
         Begin VB.ComboBox cbo医生 
            Height          =   300
            Left            =   10550
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   45
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.ComboBox cbo时段 
            Height          =   300
            ItemData        =   "frmAdviceOperate.frx":0B1C
            Left            =   3930
            List            =   "frmAdviceOperate.frx":0B2F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   45
            Width           =   1110
         End
         Begin MSMask.MaskEdBox txt时点 
            Height          =   300
            Left            =   5040
            TabIndex        =   5
            Top             =   45
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblStop 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "终止时间(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   105
            Width           =   990
         End
         Begin VB.Label lbl医生 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "停止医生(&D)"
            Height          =   180
            Left            =   9540
            TabIndex        =   6
            Top             =   105
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名: 住院号: 床号: 科室:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   120
         Width           =   2250
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6735
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6930
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   1560
      TabIndex        =   15
      Top             =   6885
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Align           =   1  'Align Top
      Height          =   1590
      Left            =   0
      TabIndex        =   2
      Top             =   5265
      Visible         =   0   'False
      Width           =   12120
      _cx             =   21378
      _cy             =   2805
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
      ForeColorSel    =   -2147483634
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
      Rows            =   6
      Cols            =   10
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
      Editable        =   2
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
   Begin VB.PictureBox picUD 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   12120
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Align           =   1  'Align Top
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   1230
      Width           =   12120
      _cx             =   21378
      _cy             =   6456
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceOperate.frx":0B51
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   8205
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceOperate.frx":0BEC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16853
            MinWidth        =   25
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":1480
            Text            =   "通过"
            TextSave        =   "通过"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":1A6A
            Text            =   "疑问"
            TextSave        =   "疑问"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":2054
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":268E
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'功能：
'0-医嘱作废:
'    只需选择要作废的医嘱
'1-停止医嘱:
'    需要指定终止时间(缺省为当前,次日生效缺省为次日零点,预定的不变)
'    护士停时需要指定停止医生
'2-确认停止:
'    只需选择需要确认的医嘱
'3-校对医嘱:
'    补录的医嘱可以修改校对时间(非补录的缺省为当前不可改,补录的缺省为开嘱时间+1m)
'4-调整计价项目:
'    增删改每个医嘱的计价项目
'5-暂停医嘱
'    选择需要暂停的医嘱
'6-启用医嘱
'    选择需要启用的医嘱
Private mfrmParent As Form
Private mMainPrivs As String
Private mlng医嘱ID As Long '用于缺省定位
Private mint类型 As Integer '0-医嘱作废,1-停止医嘱,2-确认停止,3-医嘱校对,4-调整计价项目,5-暂停医嘱,6-启用医嘱,7-停嘱审核
Private mbytUseType As Byte '0-医嘱功能调用,1-临床路径项目执行后调用
Private mstrAdviceOfItem As String '返回给临床路径的路径项目对应的医嘱ID的串,用逗号隔
Private mdateStop As Date '临床路径生成时长嘱停止时间(生成时间减1秒)
                          '转科、出院医嘱下达时传入该医嘱的开始执行时间
Private mblnAutoRead As Boolean   '发送前自动校对，此时只读取特殊医嘱来进行校对，包括：持续护理等级,病重/危医嘱,术后医嘱不发送,记录入出量,转科，出院，转院，死亡
                                  '发送时调用确认停止，自动读取传入病人的医嘱

Private mint医嘱处理范围 As Integer    '医嘱处理范围   0-所有医嘱,1-病人医嘱,2-婴儿医嘱
Private mstr缺省校对时间 As String  '第1位：0-当前系统时间,1-开始时间；第2位：开始时间大于开嘱时间时的选择，0-开始时间，1-开嘱时间；第3位：开始时间小于开嘱时间时的选择，0-开始时间，1-开嘱时间。2、3位仅在第1位为1时有效。
Private mstr缺省停止时间 As String '第1位：0-当前系统时间,1-最近一次发送的终止执行时间。第2位：至今未发送的要补发；第3位：超期发送的要收回。
Private mblnOnePati As Boolean     '单病人还是多病人模式
Private mbln发送调用 As Boolean

Private mlng病区ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng婴儿 As Long    '转科医嘱下达弹出医嘱停止界面时才传入
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级

Private mbln护士站 As Boolean
Private mlng病人性质 As Long

Private mint险类 As Integer
Private mblnRefresh As Boolean
Private mblnOK As Boolean
Private mblnReturn As Boolean

Private mrsPrice As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mstrLike As String
Private mint简码 As Integer
Private mstrRollNotify As String '操作后要进行超期收回提醒的病人(病人ID:主页ID,...)

Private mbln医技后续 As Boolean
Private mbln护士签名 As Boolean
Private mbln术后 As Boolean

Private mlng中药房 As Long
Private mlng西药房 As Long
Private mlng成药房 As Long
Private mlng发料部门 As Long
Private mblnHaveAudit As Boolean   '是否具有执业医师资格
Private mlng停嘱审核 As Long       '实习医生停止医嘱需要审核 参数
Private mlng医护科室ID As Long
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long
'PASS
Private mobjPassMap As Object  'PASS 病人信息入口参数 此变量无需赋值，PASS中的病人ID和主页ID包含在医嘱表列中传人
Private mblnPass As Boolean  'PASS权限

Private mclsMipModule As zl9ComLib.clsMipModule
Private mstrPatiClsMsg As String  '清除消息的病人 格式 "病人id1,主页id1;病人id2,主页id2;......"
Private mstrPatiKeepMsg As String  '保留消息的病人 格式 "病人id1,主页id1;病人id2,主页id2;......"
Private mblnAll As Boolean '是否是加载所有医嘱 mstrPatiAll
'重置条件
Private mblnFirst As Boolean
Private mstr病人IDs As String   '病人id,主页id;病人id,主页id;......
Private mint期效 As Integer
Private mint类别 As Integer
Private mblnPauseLast As Boolean
Private mblnFirstLoad As Boolean
Private mbytSize As Byte '字体大小 0-小字体（9号字体) 1-大字体（12 号字体）
Private mstr停嘱原因 As String
Private mbln叮嘱发送执行 As Boolean



Private Enum CtlID
    e所有 = 0
    e病人 = 1
    e婴儿 = 2
    
    e上次执行时间 = 0
    e指定时间 = 1
    
    e当前时间 = 0
    e开始时间 = 1
    
    e早于 = 0
    e晚于 = 1
End Enum

Private Const con_Date = "当前时间=__:__,今天凌晨=00:00,今天早晨=08:00,今天中午=12:00,今天下午=18:00,今天晚上=23:59"

'隐藏列
Private Const COL_ID = 0
Private Const COL_相关ID = 1
Private Const COL_组ID = 2
Private Const COL_序号 = 3
Private Const COL_诊疗类别 = 4
Private Const COL_毒理分类 = 5
Private Const COL_类型 = 6 '1-中药配方,2-检验组合
'Pass警示列
Private Const COL_警示 = 7
'输入列
Private Const COL_选择 = 8 '
Private Const COL_输入 = 9 '
Private Const COL_终止原因 = 10
'可见列
Private Const COL_标志 = 11
Private Const COL_姓名 = 12
Private Const COL_住院号 = 13
Private Const COL_床号 = 14
Private Const COL_婴儿 = 15
Private Const COL_期效 = 16
Private Const COL_开嘱时间 = 17
Private Const COL_开始时间 = 18
Private Const col_医嘱内容 = 19
Private Const COL_皮试 = 20
Private Const COL_总量 = 21
Private Const COL_单量 = 22
Private Const COL_频率 = 23
Private Const COL_用法 = 24
Private Const COL_医生嘱托 = 25
Private Const COL_执行时间 = 26
Private Const COL_终止时间 = 27 '
Private Const COL_执行科室 = 28
Private Const COL_执行性质 = 29
Private Const COL_上次执行 = 30
Private Const COL_开嘱医生 = 31
Private Const COL_校对护士 = 32 '
Private Const COL_校对时间 = 33 '
Private Const COL_停嘱医生 = 34 '
Private Const COL_停嘱时间 = 35 '
'隐藏
Private Const COL_病人ID = 36
Private Const COL_主页ID = 37
Private Const COL_诊疗项目ID = 38
Private Const COL_频率次数 = 39
Private Const COL_频率间隔 = 40
Private Const COL_间隔单位 = 41
Private Const COL_执行标记 = 42
Private Const COL_操作类型 = 43
Private Const COL_试管编码 = 44
Private Const COL_执行科室ID = 45
Private Const COL_病人科室ID = 46
Private Const COL_收费细目ID = 47
Private Const COL_单量单位 = 48
Private Const COL_前提ID = 49
Private Const COL_签名ID = 50
Private Const COL_操作人员 = 51
Private Const COL_开嘱科室ID = 52
Private Const COL_操作说明 = 53
Private Const COL_执行分类 = 54
Private Const COL_标本部位 = 55  '合理用药监测中西成药名
Private Const COL_申请序号 = 56
Private Const COL_审核状态 = 57
Private Const COL_出院科室ID = 58 '病案主页.出院科室ID
Private Const COL_病人性质 = 59


'计价清单的列值
Private Const COLP_医嘱ID = 0 '附加存放变价信息
Private Const COLP_相关ID = 1 '附加存放变价信息
Private Const COLP_诊疗类别 = 2 '附加存放变价信息
Private Const COLP_诊疗项目ID = 3
Private Const COLP_收费细目ID = 4
Private Const COLP_固定 = 5
Private Const COLP_计价医嘱 = 6
Private Const COLP_类别 = 7 '收费类别名称
Private Const COLP_收费项目 = 8
Private Const COLP_单位 = 9
Private Const COLP_数量 = 10
Private Const COLP_单价 = 11
Private Const COLP_执行科室 = 12
Private Const COLP_费用类型 = 13
Private Const COLP_从项 = 14
Private Const COLP_收费方式 = 15
Private Const COLP_收费类别 = 16 '隐藏列
Private Const COLP_执行科室ID = 17
Private Const COLP_跟踪在用 = 18
Private Const COLP_费用性质 = 19

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, ByVal int类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, _
    Optional ByVal lng医嘱ID As Long, Optional ByVal bln护士站 As Boolean, Optional blnRefresh As Boolean, Optional ByVal bytUseType As Byte, Optional ByVal strAdviceOfItem As String, _
    Optional ByVal dateStop As Date, Optional ByVal blnOnePati As Boolean, Optional ByVal strPatis As String, Optional ByVal blnAutoRead As Boolean, Optional ByVal lng婴儿 As Long, _
    Optional ByVal bln术后 As Boolean, Optional ByVal lng医护科室ID As Long, Optional ByVal bln发送调用 As Boolean, Optional ByRef objMip As Object, Optional ByRef strPatisOut As String, _
    Optional ByVal bytSize As Byte, Optional ByVal str停嘱原因 As String) As Boolean
'参数：blnRefresh=是否刷新整个主界面
'      strPatis=发送时,存在特殊医嘱的病人ID串；发送时调确认停止，当前界面选择的病人ID串(病人id,主页id;病人id,主页id;......)
'      blnAutoRead=发送时弹出先校对特殊医嘱，或者发送时调用确认停止
'      lng婴儿=转科医嘱下达弹出医嘱停止界面时才传入
'      bln发送调用=医嘱发送时，无需校对模式调用时，不刷新主界面。
'      strPatisOut传出参数，用于处理护士站消息表格  格式 "病人id1,主页id1;病人id2,主页id2;......"
    Set mfrmParent = frmParent
    mMainPrivs = MainPrivs
    mint类型 = int类型
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng婴儿 = lng婴儿
    mlng病区ID = lng病区ID
    mlng医嘱ID = lng医嘱ID
    mbln护士站 = bln护士站
    mlng医护科室ID = lng医护科室ID
    mbln术后 = bln术后
    mbytSize = bytSize
    mbln叮嘱发送执行 = Val(zlDatabase.GetPara("叮嘱需要发送执行", glngSys)) = 1
    
    If gbln医嘱终止原因 Then
        mstr停嘱原因 = str停嘱原因
    End If
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
     
    mblnOnePati = blnOnePati
    mblnAutoRead = blnAutoRead
    mbln发送调用 = bln发送调用
    If strPatis = "" Then
        mstr病人IDs = mlng病人ID & "," & mlng主页ID
    Else
        mstr病人IDs = strPatis
    End If
    
    mbytUseType = bytUseType
    mstrAdviceOfItem = strAdviceOfItem
    mdateStop = dateStop
    
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
  
    If mblnOK Then blnRefresh = mblnRefresh
    strPatisOut = mstrPatiClsMsg
    
    If mblnOK And (int类型 = 1 And Val(zlDatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达)) = 1 Or int类型 = 2 Or int类型 = 3) Then
        If Val(zlDatabase.GetPara("自动进入医嘱打印", glngSys, p住院医嘱发送)) = 1 Then
            Call frmAdvicePrint.ShowMe(frmParent, lng病人ID, lng主页ID, IIF(int类型 = 2 Or int类型 = 1, "停嘱打印", "连续打印"))
        End If
    End If
End Function

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub cbo时段_Click()
    If cbo时段.ListIndex <> -1 Then
        txt时点.Text = Split(Split(con_Date, ",")(cbo时段.ListIndex), "=")(1)
        If cbo时段.ItemData(cbo时段.ListIndex) = 1 Then
            txt时点.Text = Format(zlDatabase.Currentdate, "HH:mm")
        End If
        
        If Visible Then
            Call SetDefaultTime
        End If
    End If
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99   '合理用药审查
        If mblnPass Then
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar Is Nothing Then Exit Sub
    '护士校对界面菜单级数比医生站少一级
    If mblnPass Then
        Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
    End If
End Sub

Private Sub InitCommandBar()
'功能：初始化工具栏
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    
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
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "药嘱审查", -1, False)
    objMenu.ID = conMenu_EditPopup
    'PASS弹出菜单是在第一次点击鼠标右键时加载的cbsMain_InitCommandsPopup
    
End Sub

Private Sub cmdAlley_Click()
'功能：对病人过敏史/病生状态进行查看
    If mblnPass Then
        Call gobjPass.zlPassCmdAlleyManage(mobjPassMap)
    End If
End Sub

Private Function ResetCond() As Boolean
'功能：重置校对条件
    Dim blnSeek As Boolean
    Me.Refresh
    With frmAdviceOperateCond
        .mMainPrivs = mMainPrivs
        .mint类型 = mint类型
        .mlng病区ID = mlng病区ID
        If mlng婴儿病区ID <> 0 Then
            If mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID Then
                .mlng病区ID = mlng婴儿病区ID
            End If
        End If
        .mlng病人ID = mlng病人ID
        .Show 1, Me
        If .mblnOK Then
            mlng病区ID = .mlng病区ID
            mstr病人IDs = .mstr病人IDs
            mlng医护科室ID = mlng病区ID
            mint期效 = .mint期效
            mint类别 = .mint类别
            mblnPauseLast = .mblnPauseLast
                        
            '只选择了当前病人才定位当前医嘱
            If UBound(Split(mstr病人IDs, ";")) = 0 Then
                If Val(Split(mstr病人IDs, ",")(0)) = mlng病人ID Then blnSeek = True
            End If
            Call RefreshData(IIF(blnSeek, mlng医嘱ID, 0), True)
        End If
        ResetCond = .mblnOK
    End With
End Function

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        If tbr.Buttons("重置").Visible Then
            If Not ResetCond Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全选"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全清"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("执行"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("退出"))
    ElseIf KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("帮助"))
    ElseIf KeyCode = vbKeyF5 Then
        Call tbr_ButtonClick(tbr.Buttons("刷新"))
    ElseIf KeyCode = vbKeyF12 Then
        If tbr.Buttons("重置").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("重置"))
        End If
    ElseIf KeyCode = vbKeyF7 Then '切换输入法
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, arrTmp As Variant, strTmp As String
    
    On Error GoTo errH
    mblnRefresh = False
    mblnReturn = False
    mblnFirstLoad = True
    Call InitAdviceTable
    Call SetAdviceCol '先设置一次列属性,以便正确恢复个性化
    If mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4 Then
        Call InitPriceTable
    End If
    Call zlControl.SetPubFontSize(Me, mbytSize)
    Call RestoreWinState(Me, App.ProductName, mint类型)
    
    strSQL = "Select 婴儿科室ID,婴儿病区ID From 病案主页 Where 病人ID=[1] and 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        mlng婴儿科室ID = Val(rsTmp!婴儿科室ID & "")
        mlng婴儿病区ID = Val(rsTmp!婴儿病区ID & "")
    End If
    
    '设置公共按钮图标
    Set tbr.HotImageList = frmIcons.imgColor
    Set tbr.ImageList = frmIcons.imgGray
    tbr.Buttons("全选").Image = "全选"
    tbr.Buttons("全清").Image = "全清"
    tbr.Buttons("执行").Image = "执行"
    tbr.Buttons("重置").Image = "重置"
    tbr.Buttons("刷新").Image = "刷新"
    tbr.Buttons("帮助").Image = "帮助"
    tbr.Buttons("退出").Image = "退出"
    tbr.ButtonHeight = 500
    
    '缺省时间模式
    If mint类型 = 3 Then
        arrTmp = Array(fraOper, lblOper, optOper(e当前时间), optOper(e开始时间), lblS, cboTime(e当前时间), lblB, cboTime(e开始时间))
        mstr缺省校对时间 = zlDatabase.GetPara("医嘱缺省校对时间", glngSys, p住院医嘱发送, "001", arrTmp, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
        If Len(mstr缺省校对时间) = 1 Then mstr缺省校对时间 = mstr缺省校对时间 & "01"
    ElseIf (mint类型 = 1 Or mint类型 = 7) Then
        mblnHaveAudit = HaveAuditPriv(UserInfo.姓名)
        mlng停嘱审核 = Val(zlDatabase.GetPara("实习医生停止医嘱需要审核", glngSys, p住院医嘱下达))
        arrTmp = Array(fraStop, lblStop, optStop(e当前时间), optStop(e开始时间), txt时点, cbo时段, chkNoSend, chkRollSend)
        mstr缺省停止时间 = zlDatabase.GetPara("医嘱缺省停止时间", glngSys, p住院医嘱发送, "011", arrTmp, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
        If Len(mstr缺省停止时间) = 1 Then mstr缺省停止时间 = mstr缺省停止时间 & "11"
    End If
    
    mblnOK = False
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    Select Case mint简码
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not (mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4) Or Not gbln简码匹配方式切换 Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    mbln医技后续 = Val(zlDatabase.GetPara("医技医嘱后续处理", glngSys, p住院医嘱发送)) <> 0
    mbln护士签名 = Val(zlDatabase.GetPara("校对医嘱电子签名", glngSys, p住院医嘱发送)) <> 0 And gintCA <> 0 And Mid(gstrESign, 2, 1) = "1"
    
    '设置重置可作否,缺省重置条件
    mblnFirst = True
    mblnPauseLast = False
    mint期效 = 0: mint类别 = 0
        
    '0-医嘱作废,1-停止医嘱,2-确认停止,3-医嘱校对,4-调整计价项目,5-暂停医嘱,6-启用医嘱
    If mbln护士站 And Not mblnAutoRead And InStr(",2,3,5,6,", mint类型) > 0 Then
        tbr.Buttons("重置").Enabled = Not mblnOnePati
    Else
        tbr.Buttons("重置").Enabled = False
    End If
    tbr.Buttons("重置").Visible = tbr.Buttons("重置").Enabled 'Enabled用于判断
    
    fraStop.Visible = False
    fraOper.Visible = False
    
    If mint类型 = 0 Then
        Caption = "病人医嘱作废"
        tbr.Buttons("执行").Caption = "作废"
        tbr.Buttons("执行").ToolTipText = "作废选择的医嘱(Ctrl+E)"
    ElseIf (mint类型 = 1 Or mint类型 = 7) Then
        Caption = "病人医嘱停止"
        If mint类型 = 7 Then Caption = "病人停嘱审核"
        tbr.Buttons("执行").Caption = IIF(mint类型 = 7, "审核", "停止")
        tbr.Buttons("执行").ToolTipText = IIF(mint类型 = 7, "审核", "停止") & "选择的医嘱(Ctrl+E)"
        
        fraStop.Visible = True
        fraOper.Visible = False
        
        If mbln护士站 Then
            lbl医生.Visible = True
            cbo医生.Visible = True
        End If
        
        arrTmp = Split(con_Date, ",")
        cbo时段.Clear
        For i = 0 To UBound(arrTmp)
            strTmp = Split(arrTmp(i), "=")(0)
            cbo时段.AddItem strTmp
            If Split(arrTmp(i), "=")(1) = "__:__" Then
                cbo时段.ItemData(cbo时段.NewIndex) = 1
            End If
        Next
        cbo时段.ListIndex = 0
    ElseIf mint类型 = 2 Then
        Caption = "确认医嘱停止"
        tbr.Buttons("执行").Caption = "确认"
        tbr.Buttons("执行").ToolTipText = "确认选择的医嘱(Ctrl+E)"
    
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint类型 = 3 Then
        Caption = "病人医嘱校对"
        tbr.Buttons("执行").Caption = "校对"
        tbr.Buttons("执行").ToolTipText = "确认选择的医嘱(Ctrl+E)"
                
        stbThis.Panels(4).Visible = True
        stbThis.Panels(5).Visible = True
        
        picUD.Visible = True
        vsPrice.Visible = True
        fraStop.Visible = True
                
        '病人过敏史/病生状态可用检测 校对时
        Call zlPASSMap
        If mblnPass Then      'Pass
            Call InitCommandBar
            Call gobjPass.zlPassCmdAlleyEnable(mobjPassMap)
        End If
        For i = e早于 To e晚于
            cboTime(i).AddItem "开始时间"
            cboTime(i).AddItem "开嘱时间"
        Next
        fraStop.Visible = False
        fraOper.Visible = True
    ElseIf mint类型 = 4 Then
        Caption = "调整计价项目"
        tbr.Buttons("执行").Caption = "确认"
        tbr.Buttons("执行").ToolTipText = "确认选择项目的价目(Ctrl+E)"
        
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint类型 = 5 Then
        Caption = "病人医嘱暂停"
        tbr.Buttons("执行").Caption = "暂停"
        tbr.Buttons("执行").ToolTipText = "暂停选择的医嘱(Ctrl+E)"
    ElseIf mint类型 = 6 Then
        Caption = "病人医嘱启用"
        tbr.Buttons("执行").Caption = "启用"
        tbr.Buttons("执行").ToolTipText = "启用选择的医嘱(Ctrl+E)"
    End If
    
    Call SetFilterTime
    
    '读取部门信息
    If mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4 Then
        strSQL = "Select ID,名称 From 部门表 Where 站点='" & gstrNodeNo & "' Or 站点 is Null"
        Set mrsDept = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDept, strSQL, Me.Caption)
    End If
    
    '显示病人信息：一个病人操作的情况(批量医嘱校对可不传病人ID)
    If mlng病人ID = 0 And mint类型 = 3 Then
        lblPati.Caption = ""
        mint险类 = 0
        mlng病人性质 = 0
    Else
        strSQL = _
            " Select B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,B.出院病床," & _
            " B.住院医师,B.出院科室ID,C.名称 as 科室,B.险类,B.病人性质 " & _
            " From 病人信息 A,病案主页 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        lblPati.Caption = "姓名:" & rsTmp!姓名 & "　住院号:" & NVL(rsTmp!住院号) & _
            "　床号:" & NVL(rsTmp!出院病床) & "　科室:" & NVL(rsTmp!科室)
        mint险类 = NVL(rsTmp!险类, 0)
        mlng病人性质 = Val(rsTmp!病人性质 & "")
        
        '可选的停嘱医生:缺省为病人的住院医师或病人科室的第一个医生
        '目前不支持批量停止医嘱,因此肯定是以传入的当前病人为准读取
        If (mint类型 = 1 Or mint类型 = 7) And mbln护士站 Then
            Call Get开嘱医生(rsTmp!出院科室ID, True, NVL(rsTmp!住院医师), 0, cbo医生)
            If cbo医生.ListIndex = -1 And cbo医生.ListCount > 0 Then cbo医生.ListIndex = 0
        End If
    End If
    
    '显示医嘱内容
    If Not tbr.Buttons("重置").Enabled Then
        Call RefreshData(mlng医嘱ID, True)
        If (mblnAutoRead Or mblnOnePati) And (mint类型 = 2 Or mint类型 = 3) Then Call tbr_ButtonClick(tbr.Buttons("全选"))
    End If
    If mint类型 = 3 And mblnAutoRead Then
        stbThis.Panels(2).Text = "以上医嘱须在发送前先校对，因为校对后将自动停止其他长期医嘱。"
    End If
    
    '设置缺省时间
    If (mint类型 = 1 Or mint类型 = 3) And mdateStop = CDate(0) Then
        Call SetDefaultTime
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadMsg()
'功能：检查并处理病人消息，判断当前所选病人是否有可以校对的医嘱。
'说明：
    Dim rsTmp As ADODB.Recordset
    Dim rsMsg As ADODB.Recordset
    Dim strSQL As String
    Dim strPatis As String
    Dim strPati As String
    Dim strPatiClsMsg As String
    Dim strCurDate As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim i As Long, j As Long
    Dim blnTrans As Boolean
    Dim varArr As Variant
    Dim arrSQL As Variant
    Dim lng医嘱ID As Long
    Dim int紧急 As Integer
    Dim strMsgNo As String
    Dim strWhere As String
    
    If Not (mint类型 = 2 Or mint类型 = 3) Then Exit Sub
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    strPatis = mstr病人IDs
    strPatis = Replace(strPatis, ",", ":")
    strPatis = Replace(strPatis, ";", ",")
    If mint类型 = 3 Then
        strMsgNo = "ZLHIS_CIS_001"
        
        If gblnKSSStrict Or gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统 Then
            strWhere = strWhere & " And (Nvl(A.审核状态,0) Not in(1,3,7" & IIF(gbln血库系统 = True, "", ",4,5") & ") or a.医嘱期效=0 and a.审核状态=1 and a.紧急标志=1 and (instr(',5,6,',A.诊疗类别)>0 or A.诊疗类别='E' and B.操作类型='2'))"
        End If
        
        strSQL = "select a.id as 医嘱ID,nvl(a.紧急标志,0) as 紧急,a.病人ID,a.主页ID from 病人医嘱记录 a,诊疗项目目录 b where a.诊疗项目id=b.id(+) and A.医嘱状态=1" & strWhere & _
            " And Exists ( Select 1 From 人员表 M,执业类别 N" & _
            " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
            " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))"
    Else
        strMsgNo = "ZLHIS_CIS_002"
        
        strSQL = "select a.id as 医嘱ID,nvl(a.紧急标志,0) as 紧急,a.病人ID,a.主页ID from 病人医嘱记录 a where A.医嘱状态=8 and Nvl(a.医嘱期效,0)=0"
    End If
    strSQL = strSQL & " And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 And (a.病人ID,a.主页ID) In (Select C1,C2 From Table(f_Num2list2([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatis)
    
    strPatis = mstr病人IDs
    varArr = Split(strPatis, ";")
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    For i = 0 To UBound(varArr)
        strPati = varArr(i)
        lng病人ID = Split(strPati, ",")(0)
        lng主页ID = Split(strPati, ",")(1)
        
        rsTmp.Filter = "病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID
        
        If rsTmp.EOF Then
            '将该病人的消息设为已阅
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strMsgNo & "',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ",To_Date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'))"
            strPatiClsMsg = strPatiClsMsg & ";" & lng病人ID & "," & lng主页ID
        Else
            '有数据跟据消息清单判断，是否产生一条消息
            lng医嘱ID = rsTmp!医嘱ID: int紧急 = 1
            rsTmp.Filter = "病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID & " and 紧急=1"
            If Not rsTmp.EOF Then
                lng医嘱ID = rsTmp!医嘱ID
                int紧急 = 2
            End If
            strSQL = "select 1 From 业务消息清单 A Where a.病人id=[1] And a.就诊id=[2] And a.类型编码 =[3] And a.优先程度=[4] And a.是否已阅=0 And Rownum<2"
            Set rsMsg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, strMsgNo, int紧急)
            If rsMsg.EOF Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng主页ID & ",null," & mlng病区ID & ",2,'有新" & IIF(mint类型 = 3, "下达", "停止") & "医嘱。','0010','" & strMsgNo & "'," & lng医嘱ID & "," & int紧急 & ",0,null," & mlng病区ID & ")"
            End If
        End If
        rsTmp.Filter = 0
    Next
    
    If UBound(arrSQL) <> -1 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    mstrPatiClsMsg = Mid(strPatiClsMsg, 2)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshData(Optional ByVal lng医嘱ID As Long, Optional ByVal blnNotify As Boolean)
'功能：刷新数据
'参数：lng医嘱ID=用于医嘱定位
'      blnNotify=是否提醒特殊医嘱
    Dim blnChange As Boolean, i As Long
    Dim strPatis As String, arrPatis As Variant
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strMsg As String, strTmp As String
    Dim blnSelect As Boolean
    
    '显示医嘱内容
    Call LoadAdvice(strPatis)
    
    '读取计价数据
    If mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4 Then
        Call InitPriceRecordset
        Screen.MousePointer = 11
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            Progress = i / (vsAdvice.Rows - 1) * 100
            blnChange = False
            Call LoadPrice(i, blnChange)
            If blnChange And mint类型 = 4 Then Call SelectRow(i): blnSelect = True
        Next
        Call AppendPriceItem
        Progress = 0: Screen.MousePointer = 0
    End If
    
    If lng医嘱ID <> 0 Then
        i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If i <> -1 Then vsAdvice.Row = i
    End If
    If vsAdvice.Rows = vsAdvice.FixedRows + 1 And blnSelect = False Then
        If Val(vsAdvice.TextMatrix(vsAdvice.Rows - 1, COL_ID)) <> 0 Then
            Call SelectRow(vsAdvice.Rows - 1): blnSelect = True
        End If
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    '特殊医嘱提醒
    If blnNotify And InStr(",3,4,6,", mint类型) > 0 And strPatis <> "" Then
        arrPatis = Split(strPatis, ";")
        For i = 0 To UBound(arrPatis)
            lng病人ID = Split(arrPatis(i), ",")(0)
            lng主页ID = Split(arrPatis(i), ",")(1)
            strTmp = ExistsSpecAdvice(lng病人ID, lng主页ID)
            If strTmp <> "" Then
                strTmp = Replace(Replace(strTmp, "提醒您，", ""), vbCrLf & vbCrLf, vbCrLf)
                strMsg = strMsg & vbCrLf & strTmp
            End If
        Next
        If strMsg <> "" Then MsgBox Mid(strMsg, 3), vbInformation, gstrSysName & " - 提醒您"
    End If
End Sub

Private Sub SelectRow(ByVal lngRow As Long)
'功能：使指定行选中(包括一并给药)
    With vsAdvice
        If mint类型 = 3 Then
            Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
            .Cell(flexcpData, lngRow, COL_选择) = 1
        Else
            .TextMatrix(lngRow, COL_选择) = -1 '直接对TextMatrix时,不要用True
        End If
    End With
    Call vsAdvice_AfterEdit(lngRow, COL_选择)
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    lblPati.Left = 150
    lblPati.Top = 120
    
    If InStr(",1,3,7,", mint类型) > 0 Then
        picPati.Height = cmdAlley.Height + cbo时段.Height + 200
    Else
        picPati.Height = cmdAlley.Height + 50
    End If
    
    vsAdvice.Height = Me.ScaleHeight - cbr.Height - stbThis.Height - picPati.Height _
        - IIF(picUD.Visible, picUD.Height + vsPrice.Height, 0) - IIF(pic疑问.Visible, pic疑问.Height, 0)
    
    lngTmp = optBaby(e所有).Width + optBaby(e病人).Width + optBaby(e婴儿).Width + 40
    fraBaby.Width = lngTmp
    fraBaby.Height = optBaby(e所有).Height
    optBaby(e所有).Left = 10
    Call zlControl.SetPubCtrlPos(False, 0, lblPati, 30, fraBaby, 30, cmdAlley)
    Call zlControl.SetPubCtrlPos(False, 0, optBaby(e所有), 10, optBaby(e病人), 10, optBaby(e婴儿))
 
    If cmdAlley.Visible Then
        cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 200
        lngTmp = cmdAlley.Left - lngTmp
    Else
        lblPati.Width = Me.ScaleWidth - lblPati.Left
        lngTmp = Me.ScaleWidth - lngTmp
    End If
    fraBaby.Left = lngTmp
    
    txt疑问.Width = pic疑问.ScaleWidth - txt疑问.Left - 30
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    fraStop.Left = 0
    fraStop.Height = cbo时段.Height + 80
    fraStop.Top = cmdAlley.Height + 80
    fraStop.Width = Me.ScaleWidth
    
    lblStop.Left = 150
    lblStop.Top = 50
    Call zlControl.SetPubCtrlPos(False, 0, lblStop, 30, optStop(e上次执行时间), 10, optStop(e指定时间), 10, cbo时段, 10, txt时点, 200, chkNoSend, 10, chkRollSend, 10, lbl医生, 10, cbo医生)
        
    cbo医生.Left = Me.ScaleWidth - cbo医生.Width - lblStop.Left
    lbl医生.Left = cbo医生.Left - lbl医生.Width - 100
    
    fraOper.Width = Me.ScaleWidth
    fraOper.Top = cmdAlley.Height + 80
    fraOper.Height = cbo时段.Height + 20
    fraOper.Left = 0
    lblOper.Top = 50
    lblOper.Left = 150
    Call zlControl.SetPubCtrlPos(False, 0, lblOper, 30, optOper(e当前时间), 10, optOper(e开始时间), 100, lblS, 15, cboTime(e早于), 50, lblB, 15, cboTime(e晚于))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSetup As Boolean
    
    Call SaveWinState(Me, App.ProductName, mint类型)
    
    '保存设置参数
    blnSetup = InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0
    If mint类型 = 3 Then
        Call zlDatabase.SetPara("医嘱缺省校对时间", IIF(optOper(e当前时间).value, 0, 1) & IIF(cboTime(e早于).ListIndex = -1, 0, cboTime(e早于).ListIndex) & IIF(cboTime(e晚于).ListIndex = -1, 1, cboTime(e晚于).ListIndex), glngSys, p住院医嘱发送, blnSetup)
    ElseIf mint类型 = 1 Or mint类型 = 7 Then
        If optStop(e上次执行时间).value Then
            Call zlDatabase.SetPara("医嘱缺省停止时间", "1", glngSys, p住院医嘱发送, blnSetup)
        Else
            Call zlDatabase.SetPara("医嘱缺省停止时间", "0" & IIF(chkNoSend.value = 1, "1", "0") & IIF(chkRollSend.value = 1, "1", "0"), glngSys, p住院医嘱发送, blnSetup)
        End If
    End If
    
    Set mrsPrice = Nothing
    Set mrsDept = Nothing
    mMainPrivs = ""
    mlng医嘱ID = 0
    mint类型 = 0
    mlng病区ID = 0
    mlng病人ID = 0
    mlng主页ID = 0
    mbln护士站 = False
    mblnAll = False
    Set mclsMipModule = Nothing
    Set mobjPassMap = Nothing
End Sub

Private Sub SetDefaultTime()
'功能：根据界面设置，设置校对或停止医嘱的缺省时间
    Dim i As Long, vCurDate As Date
    
    vCurDate = zlDatabase.Currentdate
    
    '设置时间值
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 Then
                If mint类型 = 3 Then
                    If optOper(e当前时间).value Then
                        If .TextMatrix(i, COL_标志) = "补录" Then
                            .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_开嘱时间))), "yyyy-MM-dd HH:mm")
                        Else
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                            .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, IIF(cboTime(e早于).ListIndex = 0, COL_开始时间, COL_开嘱时间))
                        Else
                            .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, IIF(cboTime(e晚于).ListIndex = 0, COL_开始时间, COL_开嘱时间))
                        End If
                    End If
                Else
                    If optStop(e指定时间).value Then
                        '当前时间或指定时间
                        '如果传入的停止时间不为空，并且是护理记录，就边改变时间。
                        If mdateStop = CDate(0) Or Not (.TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" Or .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0) Then
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd " & txt时点.Text)
                            
                            If .TextMatrix(i, COL_频率) <> "持续性" Then
                                If chkNoSend.value = 0 Then    '如果不补发
                                    If .TextMatrix(i, COL_上次执行) = "" Then
                                        .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_开始时间)
                                    ElseIf .TextMatrix(i, COL_上次执行) < .TextMatrix(i, COL_输入) Then
                                        .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_上次执行)
                                    End If
                                End If
                                If chkRollSend.value = 0 Then    '如果不收回
                                    If .TextMatrix(i, COL_上次执行) > .TextMatrix(i, COL_输入) Then
                                        .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_上次执行)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '末次时间如果为空，则为开始时间
                        If mdateStop = CDate(0) Or Not (.TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" Or .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0) Then
                            If .TextMatrix(i, COL_诊疗类别) = "H" Or .TextMatrix(i, COL_频率) = "持续性" Then
                                '如果传入时间为空，护理等级特殊处理，默认为当前时间。
                                .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            Else
                                If .TextMatrix(i, COL_上次执行) = "" Then
                                    .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_开始时间)
                                Else
                                    .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_上次执行)
                                End If
                            End If
                        End If
                    End If
                    
                    .TextMatrix(i, COL_输入) = GetValidateStopDate(.TextMatrix(i, COL_输入), i)
                End If
                
                .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
            End If
        Next
    End With
End Sub

Private Sub optStop_Click(Index As Integer)
    If Me.Visible Then
        If (mint类型 = 1 Or mint类型 = 7) Then
            chkNoSend.Visible = optStop(e指定时间).value
            chkRollSend.Visible = optStop(e指定时间).value
        End If
        If optStop(e指定时间).value Then
            If cbo时段.ListIndex <> -1 Then
                If cbo时段.ItemData(cbo时段.ListIndex) = 1 Then
                    txt时点.Text = Format(zlDatabase.Currentdate, "HH:mm")
                End If
            End If
        End If
        Call SetDefaultTime
    End If
End Sub

Private Sub optOper_Click(Index As Integer)
    If Index = e当前时间 Then
        lblS.Visible = False
        lblB.Visible = False
        cboTime(e早于).Visible = False
        cboTime(e晚于).Visible = False
    Else
        lblS.Visible = True
        lblB.Visible = True
        cboTime(e早于).Visible = True
        cboTime(e晚于).Visible = True
    End If
    If Me.Visible Then Call SetDefaultTime
End Sub

Private Sub optBaby_Click(Index As Integer)
    mint医嘱处理范围 = Index
    If Not mblnFirstLoad Then
    Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
End Sub

Private Sub txt时点_GotFocus()
    Call zlControl.TxtSelAll(txt时点)
End Sub

Private Sub txt时点_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        optStop(e指定时间).value = True
        If IsDate("2010-06-22 " & txt时点.Text) Then
            Call SetDefaultTime
            Call zlControl.TxtSelAll(txt时点)
        Else
            txt时点.Text = "__:__"
        End If
    End If
End Sub

Private Function GetValidateStopDate(ByVal strDate As String, ByVal lngRow As Long) As String
'功能：获取有效的停止时间
        
    strDate = Format(strDate, "yyyy-MM-dd HH:mm")
    With vsAdvice
        
        '不应小于开始执行时间
        If strDate < Format(.Cell(flexcpData, lngRow, COL_开始时间), "yyyy-MM-dd HH:mm") Then
            strDate = Format(.Cell(flexcpData, lngRow, COL_开始时间), "yyyy-MM-dd HH:mm")
        End If
    
    End With
    GetValidateStopDate = strDate
End Function

Private Sub txt时点_Validate(Cancel As Boolean)
    If IsDate("2010-06-22 " & txt时点.Text) = False Then
        Cancel = True
    End If
End Sub

Private Sub picUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("简码方式", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, blnAutoRoll As Boolean
    
    Select Case Button.Key
        Case "全选"
            If vsAdvice.ColHidden(COL_选择) Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows + 1 And Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, COL_ID)) = 0 Then Exit Sub
            
            If mint类型 = 3 Then
                For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                    If vsAdvice.Cell(flexcpData, i, COL_选择) = Empty Then '保持疑问的不变
                        Set vsAdvice.Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                        vsAdvice.Cell(flexcpData, i, COL_选择) = 1
                    End If
                Next
            Else
                'flexcpText等同于.TextMatrix(lngRow, COL_选择) = -1(True)
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = "-1"
            End If
        Case "全清"
            If mint类型 = 3 Then
                Set vsAdvice.Cell(flexcpPicture, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = Nothing
                vsAdvice.Cell(flexcpData, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = Empty
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = "0"
            End If
        Case "刷新"
            Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        Case "重置"
            Call ResetCond
        Case "执行"
            Dim bln超期收回 As Boolean
            
            If Not CheckValid(bln超期收回) Then Exit Sub
            If Not CheckSignValid Then Exit Sub
            If ExecuteOperate Then
                If mblnHaveAudit Or (mint类型 <> 1 And mint类型 <> 7) Then
                    '医嘱校对时检查并提醒超期收回(自动)停止的医嘱
                    If mint类型 = 3 And mstrRollNotify <> "" Then
                        Call ShowRollNotify(mstrRollNotify)
                    ElseIf mint类型 = 2 And bln超期收回 Then
                        If PatiCanBilling(mlng病人ID, mlng主页ID, GetInsidePrivs(p住院医嘱发送), p住院医嘱发送) Then
                            blnAutoRoll = Val(zlDatabase.GetPara("停止后自动超期收回", glngSys, p住院医嘱发送)) = 1
                            Me.Hide
                            If frmAdviceRollSend.ShowMe(mfrmParent, mMainPrivs, mlng病区ID, mlng病人ID, mlng主页ID, True, blnAutoRoll, mlng医护科室ID, mlng婴儿病区ID) Then   '单病人模式，不弹出病人选择窗体
                                If blnAutoRoll Then
                                    MsgBox "确认停止后，超期发送的医嘱已自动收回。", vbInformation, gstrSysName
                                End If
                            End If
                        End If
                    End If
                End If
                mblnOK = True: Unload Me
            End If
        Case "帮助"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub txt疑问_Change()
    If Not pic疑问.Visible Then Exit Sub
    
    With vsAdvice
        .TextMatrix(.Row, COL_操作说明) = txt疑问.Text
    End With
End Sub

Private Sub txt疑问_GotFocus()
    zlControl.TxtSelAll txt疑问
End Sub

Private Sub txt疑问_KeyPress(KeyAscii As Integer)
    If txt疑问.MaxLength <> 0 And Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If zlCommFun.ActualLen(txt疑问.Text) > txt疑问.MaxLength Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If vsAdvice.Col = COL_终止原因 Then
        vsAdvice.ComboList = "..."
        vsAdvice.Editable = flexEDKbdMouse
    Else
        vsAdvice.ComboList = ""
    End If

    If NewRow = OldRow Then Exit Sub
    'PASS
    If mblnPass And mint类型 = 3 Then
        Call gobjPass.zlPassSetDrug(mobjPassMap)
    End If
    
    With vsAdvice
        '校对疑问说明
        If mint类型 = 3 Then
            If .Cell(flexcpData, .Row, COL_选择) = 2 Then
                txt疑问.Text = .TextMatrix(.Row, COL_操作说明)
                pic疑问.Visible = True
            Else
                pic疑问.Visible = False
            End If
            Call Form_Resize
        End If
        
        '显示计价项目
        If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
            If (mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4) And Not mrsPrice Is Nothing Then
                Call ShowPrice(NewRow)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_医嘱内容 Then
        vsAdvice.AutoSize Col
    ElseIf Col = COL_皮试 Then
        If vsAdvice.ColWidth(Col) > 1200 Then vsAdvice.ColWidth(Col) = 1200
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Or Col = COL_输入 Or Col = COL_终止原因 Or Col = COL_警示 Then Cancel = True 'Pass
End Sub

Private Sub vsAdvice_DblClick()
    
    If mblnPass And mint类型 = 3 Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    
    With vsAdvice
        If mint类型 = 3 And .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_姓名: lngRight = COL_开始时间
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_频率: lngRight = COL_用法
        End If
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_皮试: lngRight = COL_皮试
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
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
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
'功能：定位到下一输入单元或输入校对标志
    Dim blnGroup As Boolean, i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        With vsAdvice
            If .ColHidden(COL_选择) And .ColHidden(COL_输入) Then
                If .Row + 1 <= .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    .Row = .FixedRows
                End If
            Else
                If .Col = COL_选择 Then
                    If Not .ColHidden(COL_输入) Then
                        .Col = COL_输入
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                    End If
                ElseIf .Col = COL_输入 Then
                    If Not .ColHidden(COL_终止原因) Then
                        .Col = COL_终止原因
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                        .Col = COL_选择
                    End If
                ElseIf .Col = COL_终止原因 Then
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    .Col = COL_选择
                Else
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    If Not .ColHidden(COL_选择) Then .Col = COL_选择
                End If
            End If
            Call .ShowCell(.Row, .Col)
        End With
    ElseIf KeyAscii = 32 Then
        With vsAdvice
            If mint类型 = 3 And .Col = COL_选择 Then
                KeyAscii = 0
                
                If .Cell(flexcpData, .Row, .Col) = Empty Then
                    Set .Cell(flexcpPicture, .Row, .Col) = frmIcons.imgTrueFalse.ListImages("T").Picture
                    .Cell(flexcpData, .Row, .Col) = 1
                ElseIf .Cell(flexcpData, .Row, .Col) = 1 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = frmIcons.imgQuestion.ListImages("Q").Picture
                    .Cell(flexcpData, .Row, .Col) = 2
                    If .TextMatrix(.Row, COL_诊疗类别) = "K" And Val(.TextMatrix(.Row, COL_审核状态)) <> 0 Then
                        MsgBox "该医嘱为已审核的输血医嘱不能设为校对疑问。", vbInformation, gstrSysName
                        Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                        .Cell(flexcpData, .Row, .Col) = Empty
                    End If
                ElseIf .Cell(flexcpData, .Row, .Col) = 2 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                    .Cell(flexcpData, .Row, .Col) = Empty
                End If
                Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                
                If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
                    If .Row - 1 >= .FixedRows Then
                        blnGroup = Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID))
                    End If
                    If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                        blnGroup = Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID))
                    End If
                    If blnGroup Then
                        For i = .Row - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                        For i = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                '一组申请序号的会诊医嘱
                If Val(.TextMatrix(.Row, COL_申请序号)) <> 0 And .TextMatrix(.Row, COL_诊疗类别) = "Z" Then
                    For i = .Row - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(.Row, COL_申请序号)) Then
                            .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                        Else
                            Exit For
                        End If
                    Next
                    For i = .Row + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(.Row, COL_申请序号)) Then
                            .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                        Else
                            Exit For
                        End If
                    Next
                End If
                
            End If
        End With
    Else
        If vsAdvice.Col = COL_终止原因 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAdvice_CellButtonClick(vsAdvice.Row, vsAdvice.Col)
            Else
                vsAdvice.ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End If
End Sub

Private Function AcceptInput(ByVal Row As Long, ByVal Col As Long) As Boolean
    Dim strTmp As String, vPause As Date
    Dim blnDoSame As Boolean
    Dim lng医嘱ID As Long
    Dim strTmpTim As String
    
    AcceptInput = False
    With vsAdvice
        If .EditText <> "" Then .EditText = zlStr.FullDate(.EditText)
        If .EditText = .TextMatrix(Row, Col) Then AcceptInput = True: Exit Function
    
        '检查输入的有效性
        If Not IsDate(.EditText) Then
            MsgBox "请输入一个有效的" & .TextMatrix(0, Col) & " 。", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
    
        If (mint类型 = 1 Or mint类型 = 7) Then '检查终止时间
            '必须大于开始执行时间
            If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                MsgBox "输入的执行终止时间必须大于医嘱的开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
            
            '停医嘱时，检查执行登记情况
            If mint类型 = 1 And Not (.TextMatrix(Row, COL_诊疗类别) = "Z" And InStr(",4,14,9,10,12,", "," & .TextMatrix(Row, COL_操作类型) & ",") > 0) Then
                '获取医嘱id
                lng医嘱ID = IIF(InStr(",5,6,", .TextMatrix(Row, COL_诊疗类别)) > 0, .TextMatrix(Row, COL_相关ID), .TextMatrix(Row, COL_ID))
                '获取时间
                strTmpTim = GetAdviceStopTime(lng医嘱ID)
                '消息提示
                If IsDate(strTmpTim) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(strTmpTim, "yyyy-MM-dd HH:mm") Then
                        strTmp = .EditText 'MsgBox一现,EditText就空了,所以要记录
                        MsgBox "不能停止到执行时间 " & strTmpTim & " 之前，请调整停止时间，如果确实要停止到执行时间之前，请先取消执行登记。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
            
            '不应小于上次执行时间
            If IsDate(.Cell(flexcpData, Row, COL_上次执行)) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                    strTmp = .EditText 'MsgBox一现,EditText就空了,所以要记录
                    If MsgBox("输入的执行终止时间小于医嘱的上次执行时间 " & Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                    End If
                End If
            End If
            
        ElseIf mint类型 = 2 Then  '检查确认停止时间
            If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                MsgBox "确认停止医嘱的时间不能小于医嘱的执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        ElseIf mint类型 = 3 Then  '检查校对时间
            If Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "输入的校对时间不能小于医嘱的开嘱时间 " & Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            Else
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "输入的校对时间不能小于医嘱的开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        ElseIf mint类型 = 5 Then '检查暂停时间
            '应>=开始执行时间,因为该时间点尚未执行
            If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                MsgBox "医嘱的暂停时间应大于等于开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
            '应>上次执行时间,因为该时间点已执行
            If .TextMatrix(Row, COL_上次执行) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的暂停时间应大于上次执行时间 " & Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            '应<执行终止时间,因为该时间点执行有效
            If .TextMatrix(Row, COL_终止时间) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的暂停时间应小于执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
            vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
            If vPause <> CDate(0) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的暂停时间应大于上次暂停后的启用时间 " & Format(vPause, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        ElseIf mint类型 = 6 Then '检查启用时间
            '应>暂停时间
            vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
            If vPause <> CDate(0) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的启用时间应大于上次暂停时间 " & Format(vPause, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            
            '应<=执行终止时间
            If .TextMatrix(Row, COL_终止时间) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的启用时间应小于等于执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        End If
        
            
        .TextMatrix(Row, Col) = IIF(.EditText = "" And strTmp <> "", strTmp, .EditText)
        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
        
        Call vsAdvice_AfterEdit(Row, Col) '一并给药的一并更改:提示后不会自动执行该事件
        
        '设置为相同时间(校对,暂停,启用)
        blnDoSame = InStr(",1,2,3,5,6,", "," & mint类型 & ",") > 0
        If blnDoSame Then
            If Not VsfOnlySelOneRow(Row) Then
                Select Case mint类型
                Case 1
                    strTmp = "停止"
                Case 2
                    strTmp = "确认停止"
                Case 3
                    strTmp = "校对"
                Case 5
                    strTmp = "暂停"
                Case 6
                    strTmp = "启用"
                End Select
                
                If MsgBox("要设置所有已选择的医嘱都在这个时间" & strTmp & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call SetSameTime(Row)
                End If
            End If
        End If
    End With
    AcceptInput = True
End Function

Private Function VsfOnlySelOneRow(lngRow As Long) As Boolean
'功能：判断是否仅有一行可见行(一并给药算一行)
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Call RowIn一并给药(lngRow, lngBegin, lngEnd)
    
    VsfOnlySelOneRow = True
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) And i <> lngRow And (i < lngBegin Or i > lngEnd) Then
                If mint类型 = 3 Then
                    If .Cell(flexcpData, i, COL_选择) <> Empty Then
                        VsfOnlySelOneRow = False
                        Exit Function
                    End If
                Else
                    If Val(.TextMatrix(i, COL_选择)) <> 0 Then
                        VsfOnlySelOneRow = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strInput As String
    Dim strMatch As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI

    If KeyAscii = 13 Then
        If Col = COL_输入 Then
            mblnReturn = True
        End If
        
        With vsAdvice
            If Col = COL_终止原因 And .EditText <> "" Then
                strInput = UCase(.EditText)
                
                If IsNumeric(strInput) Then
                    strMatch = "    A.编码 Like [1]" '数字匹简码
                ElseIf zlCommFun.IsCharAlpha(strInput) Then
                    strMatch = "   a.简码 Like [1]" '字母时只匹配简码
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strMatch = "   a.名称 Like [1]" '名称
                Else
                    strMatch = "  (a.名称 Like [1] or a.简码 Like [1] or A.编码 Like [1])"
                End If
                strSQL = "select a.编码 as id, a.编码,a.名称,a.简码 from 停嘱原因 a where " & strMatch & " order by a.编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "停嘱原因", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%")
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COL_终止原因) = rsTmp!名称 & ""
                    .Cell(flexcpData, Row, COL_终止原因) = rsTmp!名称 & ""
                    .EditText = .TextMatrix(Row, Col)
                    Call SetSame原因(Row)
                Else
                    .TextMatrix(Row, COL_终止原因) = .EditText
                    .Cell(flexcpData, Row, COL_终止原因) = .EditText
                    .EditText = .TextMatrix(Row, Col)
                    Call SetSame原因(Row)
                End If
            End If
        End With
    Else
        If Col = COL_输入 Then
            If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        ElseIf Col = COL_终止原因 Then
            If KeyAscii = 39 Then KeyAscii = 0 '单引号
        End If
    End If
End Sub

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_输入 Then
        vsAdvice.Refresh    '如果有弹出提示，不刷新的话，一并给药通过Drawcell被擦除的单元格会再次显示
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        Else
            If mblnReturn Then
                Call vsAdvice_KeyPress(13) '定位到一下输入单元
            End If
        End If
    End If
    mblnReturn = False
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：一并给药的一起输入
    Dim lngBegin As Long, lngEnd As Long, i As Long
        
    With vsAdvice
        '一并给药的一起选择或输入
        If (Col = COL_选择 Or Col = COL_输入 Or Col = COL_终止原因) And InStr(",5,6,", .TextMatrix(Row, COL_诊疗类别)) > 0 Then
            If RowIn一并给药(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If i <> Row Then
                        .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                        .Cell(flexcpData, i, Col) = .Cell(flexcpData, Row, Col)
                        Set .Cell(flexcpPicture, i, Col) = .Cell(flexcpPicture, Row, Col)
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnReturn = False
    If Col <> COL_选择 And Col <> COL_输入 And Col <> COL_终止原因 Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(Row, COL_ID)) = 0 Then
        Cancel = True
    ElseIf mint类型 = 3 Then
        If Col = COL_输入 And Not (vsAdvice.TextMatrix(Row, COL_标志) = "补录" Or InStr(GetInsidePrivs(p住院医嘱发送), "修改校对时间") > 0) Then
            Cancel = True '校对医嘱时,非补录的校对时间不可更改
        ElseIf Col = COL_选择 Then
            Cancel = True '不能直接编辑
        End If
    End If
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;相关ID;组ID;序号;诊疗类别;毒理分类;中药;,240,4;,300,4;,1530,1;终止原因,1530,1;" & _
        "标志,500,4;姓名,750,1;住院号,750,1;床号,500,1;婴儿,500,1;期效,500,4;开嘱时间;生效时间,1530,1;" & _
        "医嘱内容,3000,1;,375,1;总量,850,1;单量,850,1;频率,1000,1;用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;" & _
        "终止时间,1530,1;执行科室,850,1;执行性质,850,1;上次执行,1530,1;" & _
        "开嘱医生,850,1;校对护士,850,1;校对时间,1530,1;停嘱医生,850,1;停嘱时间,1530,1;" & _
        "病人ID;主页ID;诊疗项目ID;频率次数;频率间隔;间隔单位;执行标记;操作类型;试管编码;" & _
        "执行科室ID;病人科室ID;收费细目ID;单量单位;前提ID;签名ID;操作人员;开嘱科室ID;操作说明;执行分类;标本部位;申请序号;审核状态;出院科室ID;病人性质"
    arrHead = Split(strHead, ";")
    With vsAdvice
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
            End If
        Next
        .ColHidden(COL_警示) = True 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "医嘱ID;相关ID;诊疗类别;诊疗项目ID;收费细目ID;固定;" & _
        "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;单位,500,4;计价数量,850,1;单价,850,7;" & _
        "执行科室,1000,1;费用类型,850,1;从项,450,4;收费方式,1500,1;收费类别;执行科室ID;跟踪在用;费用性质"
    arrHead = Split(strHead, ";")
    With vsPrice
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
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub SetAdviceCol()
'功能：设置一些可见列及编辑属性,应在表格数据装入后调用
    With vsAdvice
        .TextMatrix(0, COL_选择) = "选"
        .Editable = flexEDKbdMouse
        
        .ColHidden(COL_终止原因) = True
        
        '根据情况设置列的可见性
        If mint类型 = 0 Then
            '医嘱作废
            .ColHidden(COL_输入) = True
            .ColHidden(COL_上次执行) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf (mint类型 = 1 Or mint类型 = 7) Then
            '停止医嘱
            .TextMatrix(0, COL_输入) = "终止时间"
            .ColHidden(COL_终止时间) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            If gbln医嘱终止原因 Then
                .ColHidden(COL_终止原因) = False
            End If
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 2 Then
            '确认停止
            .TextMatrix(0, COL_输入) = "确认时间"
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 3 Then
            '医嘱校对
            .TextMatrix(0, COL_输入) = "校对时间"
            .ColHidden(COL_上次执行) = True
            .ColHidden(COL_校对护士) = True
            .ColHidden(COL_校对时间) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .Cell(flexcpPictureAlignment, .FixedRows, COL_选择, .Rows - 1, COL_选择) = 4
            .Cell(flexcpForeColor, .FixedRows, COL_开始时间, .Rows - 1, COL_开始时间) = vbBlue          '蓝色
            
        ElseIf mint类型 = 4 Then
            '调整计价项目
            .ColHidden(COL_输入) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 5 Then
            '暂停医嘱
            .TextMatrix(0, COL_输入) = "暂停时间"
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 6 Then
            '启用医嘱
            .TextMatrix(0, COL_输入) = "启用时间"
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        End If
        
        '设置冻结列
        If Not .ColHidden(COL_输入) Then
            If .TextMatrix(0, COL_输入) = "终止时间" Then
                .FrozenCols = COL_终止原因 + 1 - .FixedCols
            Else
                .FrozenCols = COL_输入 + 1 - .FixedCols
            End If
            .SheetBorder = vbBlack
        ElseIf Not .ColHidden(COL_选择) Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
            .SheetBorder = vbBlack
        End If
        
        '可输入列标识
        .Cell(flexcpBackColor, .FixedRows, COL_选择, .Rows - 1, COL_终止原因) = COLEditBackColor       '浅绿
    End With
End Sub

Private Function GetWhere() As String
'功能：根据窗体功能产生医嘱条件串
'说明：假设"病人医嘱记录"别名为"A"
    Dim strSQL As String
    
    If mint类型 = 0 Then
        '医嘱作废:已校对,但未发送过的临嘱或长嘱。已暂停的长嘱也可以直接作废。
        '临时自由医嘱校对后自动停止，这种也允许作废
        strSQL = " And (A.医嘱状态 Not IN(1,2,4,8,9) And A.上次执行时间 is NULL Or " & IIF(mbln叮嘱发送执行, "A.诊疗项目ID is Null And A.医嘱状态<>4", "A.医嘱期效=1 And A.诊疗项目ID is Null And A.医嘱状态=8") & ")"
    ElseIf (mint类型 = 1 Or mint类型 = 7) Then
        '停止医嘱:长嘱,已暂停的也可以直接停止,含中药配方长嘱
        strSQL = " And A.医嘱状态 Not IN(1,2,4,8,9) And Nvl(A.医嘱期效,0)=0"
    ElseIf mint类型 = 2 Then
        '确认停止:停止状态的长嘱
        strSQL = " And A.医嘱状态=8 And Nvl(A.医嘱期效,0)=0"
    ElseIf mint类型 = 3 Then
        '医嘱校对:对新开的，开嘱医生具有资格的或已审核的医嘱进行校对。
        strSQL = " And A.医嘱状态=1 And Exists(" & _
            "Select M.姓名 From 人员表 M,执业类别 N" & _
            " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
            " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
            " )"
    ElseIf mint类型 = 4 Then
        '调整计价项目
        strSQL = " And A.医嘱状态 Not IN(1,2,4,8,9)"
    ElseIf mint类型 = 5 Then
        '暂停医嘱:长嘱,含中药配方长嘱
        strSQL = " And A.医嘱状态 IN(3,5,7) And Nvl(A.医嘱期效,0)=0"
    ElseIf mint类型 = 6 Then
        '启用医嘱
        strSQL = " And A.医嘱状态=6"
    End If
    GetWhere = strSQL
End Function

Private Function LoadAdvice(strPatis As String) As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
'参数：str病人IDs=用于返回实际有数据的病人串:"病人ID,主页ID,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsPause As New ADODB.Recordset
    Dim str成药 As String, str中药 As String
    Dim strSQL As String, strWhere As String
    Dim bln给药途径 As Boolean, bln输血途径 As Boolean
    Dim lng病人ID As Long, lng主页ID As Long
    Dim i As Long, j As Long, k As Long
    Dim str婴儿 As String, str科室s As String
    Dim vCurDate As Date, strTmp As String, strDepts As String
    Dim int图标数 As Integer
    Dim str会诊医嘱IDs As String
    Dim bln会诊 As Boolean
    
    Screen.MousePointer = 11
    Me.Refresh
    On Error GoTo errH
        
    '----------------------------------------------------------------------
    strPatis = ""
    With vsAdvice
        .Rows = .FixedRows
        .ColHidden(COL_姓名) = True
        .ColHidden(COL_住院号) = True
        .ColHidden(COL_床号) = True
        .ColHidden(COL_婴儿) = True
    End With
    
    '----------------------------------------------------------------------
    strDepts = GetUser科室IDs(True)
    strWhere = GetWhere
    strWhere = strWhere & IIF(Not mbln护士站 Or Not mbln医技后续, " And A.前提ID is NULL", "")
    
    If DeptIsWoman(0, Get科室IDs(mlng病区ID)) Then
        '医嘱处理范围
        If mblnFirstLoad Then
            fraBaby.Visible = True
            mint医嘱处理范围 = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
            optBaby(mint医嘱处理范围).value = True
            mblnFirstLoad = False
        End If
    Else
        mblnFirstLoad = True
        fraBaby.Visible = False
        optBaby(e所有).value = True
    End If
    
    '校对的医嘱范围限制
    If mint类型 = 3 Then
        If InStr(GetInsidePrivs(p住院医嘱发送), "全院医嘱校对") = 0 Then
            If gbln会诊科室下达医嘱处理 Then
                strWhere = strWhere & " And (A.开嘱科室ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X) And nvl(a.会诊医嘱id,0)=0 Or instr(','||[7]||',',','||nvl(a.会诊医嘱id,0)||',')>0)"
                bln会诊 = True
            Else
                strWhere = strWhere & " And A.开嘱科室ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X)"
            End If
        End If
        If mblnAutoRead Then
            '改为不限制，所有未校对的都读取，因为只有这样同时校对其他医嘱，在最后校对特殊医嘱时，才能自动停止这些医嘱（先校对）
        End If
        If gblnKSSStrict Or gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统 Then
            strWhere = strWhere & " And (Nvl(A.审核状态,0) Not in(1,3,7" & IIF(gbln血库系统 = True, "", ",4,5") & ") or a.医嘱期效=0 and a.审核状态=1 and a.紧急标志=1 and (instr(',5,6,',A.诊疗类别)>0 or A.诊疗类别='E' and B.操作类型='2'))"
        End If
    End If
    
    If mint类型 <> 2 Then
        '批量操作时设置的条件
        If mint期效 <> 0 Then
            strWhere = strWhere & " And Nvl(A.医嘱期效,0)=" & mint期效 - 1
        End If
        If mint类别 <> 0 Then
            If mint类别 = 1 Then
                '药品类
                strWhere = strWhere & _
                    " And (A.诊疗类别 IN('5','6','7')" & _
                    " Or (A.诊疗类别='E' And A.相关ID is Not NULL)" & _
                    " Or Exists(Select ID From 病人医嘱记录 S Where 诊疗类别 IN('5','6','7') And S.相关ID=A.ID And 病人ID=[1])" & _
                    " )"
            ElseIf mint类别 = 2 Then
                '其他类
                strWhere = strWhere & _
                    " And Not A.诊疗类别 IN('5','6','7')" & _
                    " And Not(A.诊疗类别='E' And A.相关ID is Not NULL)" & _
                    " And Not Exists(Select ID From 病人医嘱记录 S Where 诊疗类别 IN('5','6','7') And S.相关ID=A.ID And 病人ID=[1])"
            End If
        End If
    End If
    
    '临床路径的医嘱
    If mbytUseType = 1 And (mint类型 = 1 Or mint类型 = 7) Then
        strWhere = strWhere & " And A.ID IN(Select Column_Value From Table(f_Num2list([3])))" & _
            " And Not(Nvl(A.诊疗类别,'ZY')='H' And b.操作类型='1' And b.执行频率=2)" & _
            " And Not(Nvl(A.诊疗类别,'ZY')='Z' And b.操作类型 IN('4','14', '9', '10', '12'))"
        vCurDate = mdateStop
    Else
        If (mint类型 = 1 Or mint类型 = 7) Then
            strWhere = strWhere & "  And Not(Nvl(A.诊疗类别,'ZY')='H' And b.操作类型='1' And b.执行频率=2)"
            If mlng婴儿 <> 0 Then   '婴儿开转科医嘱时,只停止该婴儿的.而母亲的转科医嘱时，婴儿由于没有独立身份也应一并处理
                strWhere = strWhere & "  And A.婴儿 = " & mlng婴儿
            Else
                If mbln术后 Then
                    strWhere = strWhere & "  And NVL(A.婴儿,0) = 0 "
                End If
            End If
            strWhere = strWhere & IIF(mint类型 = 7, " And A.审核标记=2 ", IIF(Not mblnHaveAudit, " And NVL(A.审核标记,0)<>2 ", ""))
        End If
        If (mint类型 = 1 Or mint类型 = 7) And mdateStop <> CDate(0) Then
            vCurDate = mdateStop
        Else
            vCurDate = zlDatabase.Currentdate
        End If
    End If
    '备用医嘱不允许暂停
    If mint类型 = 5 Then
        strWhere = strWhere & " And NVL(a.执行频次,'无')<>'必要时' And NVL(a.执行频次,'无')<>'需要时' "
    End If
    
    mblnAll = (mint医嘱处理范围 = 0 And mint期效 = 0 And mint类别 = 0)
    
    '----------------------------------------------------------------------
    For k = 0 To UBound(Split(mstr病人IDs, ";"))
        lng病人ID = Split(Split(mstr病人IDs, ";")(k), ",")(0)
        lng主页ID = Split(Split(mstr病人IDs, ";")(k), ",")(1)
        If bln会诊 Then str会诊医嘱IDs = Get会诊医嘱IDs(lng病人ID, lng主页ID, strDepts)
        '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法
        '自由录入的医嘱，允许作废，停止，确认停止，校对
        strSQL = _
            "Select /*+ RULE */ A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,A.序号,Nvl(A.诊疗类别,'*') as 诊疗类别,C.毒理分类,NULL as 中药," & _
                " A.审查结果,NULL as 选择,NULL as 输入,null as 原因,Decode(A.紧急标志,1,'紧急',2,'补录','普通') as 标志,A.姓名,P.住院号,P.当前床号 as 床号," & _
                " Decode(Nvl(A.婴儿,0),0,'病人','婴儿'||A.婴儿) as 婴儿,Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
                " To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间,To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,A.医嘱内容,A.皮试结果 as 皮试," & _
                " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'4',A.总给予量||F.计算单位,'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
                " Decode(A.单次用量,NULL,NULL,A.单次用量||Decode(A.诊疗类别,'4',F.计算单位,B.计算单位)) as 单量,A.执行频次 as 频率," & _
                " Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法,A.医生嘱托," & _
                " A.执行时间方案 as 执行时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间," & _
                " Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室," & _
                " Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质, To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & _
                " A.开嘱医生,A.校对护士,To_Char(A.校对时间,'YYYY-MM-DD HH24:MI') as 校对时间," & _
                " A.停嘱医生,To_Char(A.停嘱时间,'YYYY-MM-DD HH24:MI') as 停嘱时间,A.病人ID,A.主页ID,A.诊疗项目ID,A.频率次数,A.频率间隔,A.间隔单位,A.执行标记," & _
                " B.操作类型,B.试管编码,A.执行科室ID,A.病人科室ID,A.收费细目ID,B.计算单位 as 单量单位,A.前提ID,A.新开签名ID as 签名ID,A.开嘱医生,A.开嘱科室ID," & IIF(mint类型 = 7, "S.操作说明,", "Null as 操作说明,") & _
                " b.执行分类,A.标本部位,a.申请序号,a.审核状态,g.出院科室ID,g.病人性质," & IIF(mint类型 = 3, "Decode(a.医嘱状态,1,Decode(a.校对护士,Null,0,1),0) as 疑问更正", "Null as 疑问更正") & ",D.高危药品,d.是否易至跌倒"
        strSQL = strSQL & _
            IIF(mint类型 = 0, ",J.审查结果 as 处方审查结果", ",NULL as 处方审查结果") & _
            " From 病人医嘱记录 A,病人信息 P,病案主页 G,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,收费项目目录 F" & IIF(mint类型 = 7, ",病人医嘱状态 S", "") & _
            IIF(mint类型 = 0, ",处方审查明细 I,处方审查记录 J", "") & _
            " Where A.病人ID=P.病人ID And A.诊疗项目ID=B.ID" & IIF(InStr(",0,1,2,3,", mint类型) > 0, "(+)", "") & _
                " And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+) And p.病人ID=G.病人ID And P.主页ID=G.主页ID " & _
                " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=F.ID(+)" & _
                " And (Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL) Or A.诊疗类别='E' And B.操作类型='8')" & _
                IIF(mint类型 = 7, " And A.ID=S.医嘱ID And S.操作类型=13", "") & " And A.病人ID=[1] And A.主页ID=[2]" & _
                IIF(mint类型 = 0, " And a.ID = i.医嘱ID(+) And I.审方ID = J.ID(+) and (I.最后提交 =1 Or I.审方ID is NULL)", "") & _
                " And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
                Decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & _
                " And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3" & strWhere & " And (G.婴儿科室ID is null or G.婴儿科室ID is not null and (G.婴儿病区ID=[6] or G.婴儿科室ID=[6]) and NVL(A.婴儿,0)<>0 or G.婴儿科室ID is not null and (G.婴儿病区ID<>[6] and G.婴儿科室ID<>[6]) and NVL(A.婴儿,0)=0)" & _
            " Order by Nvl(A.婴儿,0),组ID,A.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, mstrAdviceOfItem, strDepts, mlng病区ID, mlng医护科室ID, str会诊医嘱IDs)
        
        If Not rsTmp.EOF Then
            strPatis = strPatis & ";" & lng病人ID & "," & lng主页ID
            If InStr(str科室s & ",", "," & rsTmp!病人科室id & ",") = 0 Then
                str科室s = str科室s & "," & rsTmp!病人科室id
            End If
            
            '暂停医嘱时读取医嘱的上次启用时间(不一定有)
            '启用医嘱时读取医嘱的暂停时间
            If mint类型 = 5 Or mint类型 = 6 Then
                strSQL = "Select B.医嘱ID,Max(B.操作时间) as 上次时间" & _
                    " From 病人医嘱记录 A,病人医嘱状态 B" & _
                    " Where A.ID=B.医嘱ID And B.操作类型=" & IIF(mint类型 = 5, 7, 6) & _
                    " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
                    " And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And A.病人来源<>3" & strWhere & _
                    " Group by B.医嘱ID"
                Set rsPause = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
            End If
            
            With vsAdvice
                .Redraw = flexRDNone
                Do While Not rsTmp.EOF
                    '添加新行
                    int图标数 = 0
                    strTmp = ""
                    For i = 0 To rsTmp.Fields.Count - 2
                        strTmp = strTmp & vbTab & Replace(NVL(rsTmp.Fields(i).value), vbTab, "")
                    Next
                    .AddItem Mid(strTmp, 2): i = .Rows - 1
                    
                    If mdateStop <> CDate(0) And (mint类型 = 1 Or mint类型 = 7) Then
                        .TextMatrix(i, COL_选择) = "-1"
                    End If
                                        
                    '是否显示婴儿列
                    If InStr(str婴儿 & ",", "," & .TextMatrix(i, COL_婴儿) & ",") = 0 Then
                        If str婴儿 <> "" Then .ColHidden(COL_婴儿) = False
                        str婴儿 = str婴儿 & "," & .TextMatrix(i, COL_婴儿)
                    End If
                    
                    '病人之间的间隔线
                    If .TextMatrix(i, COL_住院号) <> .TextMatrix(i - 1, COL_住院号) And i - 1 >= .FixedRows Then
                        .CellBorderRange i - 1, .FixedCols, i - 1, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                    End If
                    
                    '成药及中药的一些处理
                    bln给药途径 = False: bln输血途径 = False
                    If .TextMatrix(i, COL_诊疗类别) = "E" Then
                        If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                            If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                                bln给药途径 = True
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        '显示成药的给药途径
                                        .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法)
                                        .TextMatrix(j, COL_操作类型) = .TextMatrix(i, COL_操作类型)
                                        .TextMatrix(j, COL_执行分类) = .TextMatrix(i, COL_执行分类)
                                        '显示成药的执行性质
                                        If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                            If Val(.TextMatrix(j, COL_执行标记)) = 2 Then
                                                .TextMatrix(j, COL_执行性质) = "不取药"
                                            Else
                                                .TextMatrix(j, COL_执行性质) = "自备药"
                                            End If
                                        ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                            .TextMatrix(j, COL_执行性质) = "离院带药"
                                        Else
                                            .TextMatrix(j, COL_执行性质) = IIF(Val(.TextMatrix(j, COL_执行标记)) = 1, "自取药", "")
                                        End If
                                        
                                        '滴速
                                        .TextMatrix(j, COL_皮试) = .TextMatrix(i, COL_医生嘱托)
                                    Else
                                        Exit For
                                    End If
                                Next
                            ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                                If .TextMatrix(i - 1, COL_诊疗类别) = "7" Then
                                    .TextMatrix(i, COL_类型) = "1" '中药配方
                                ElseIf .TextMatrix(i - 1, COL_诊疗类别) = "C" Then
                                    .TextMatrix(i, COL_类型) = "2" '检验组合
                                    
                                    '采集方式的管码与一并的第一个检验相同
                                    j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_相关ID)
                                    If j <> -1 Then
                                        .TextMatrix(i, COL_试管编码) = .TextMatrix(j, COL_试管编码)
                                        .TextMatrix(i, COL_开始时间) = .TextMatrix(j, COL_开始时间)
                                        .Cell(flexcpData, i, COL_开始时间) = CStr(.TextMatrix(j, COL_开始时间))
                                    End If
                                End If
                                
                                '显示中药配方或检验组合的执行科室
                                .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)
                                
                                If .TextMatrix(i - 1, COL_诊疗类别) = "7" Then
                                    '显示中药配方执行性质
                                    If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        If Val(.TextMatrix(i - 1, COL_执行标记)) = 2 Then
                                            .TextMatrix(i, COL_执行性质) = "不取药"
                                        Else
                                            .TextMatrix(i, COL_执行性质) = "自备药"
                                        End If
                                    ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(i, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(i, COL_执行性质) = IIF(Val(.TextMatrix(i - 1, COL_执行标记)) = 1, "自取药", "")
                                    End If
                                Else
                                    .TextMatrix(i, COL_执行性质) = ""
                                End If
                                
                                '删除单味中药行,以及检验组合中的检验项目;同时判断检验申请
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        .RemoveItem j: i = .Rows - 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        ElseIf .TextMatrix(i - 1, COL_诊疗类别) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_相关ID)) Then
                            bln输血途径 = True
                            '显示输血途径
                            .TextMatrix(i - 1, COL_用法) = .TextMatrix(i, COL_用法)
                        Else
                            .TextMatrix(i, COL_执行性质) = ""
                        End If
                    End If
                                                                    
                    '处理可见行的的一些标识
                    If Not (bln给药途径 Or bln输血途径) And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                        '处理小数点问题,暂未想到办法
                        If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                            .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                        End If
                        If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                            .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                        End If
                    
                        '时间以MM-DD HH:MI格式显示,以CellData进行判断
                        .Cell(flexcpData, i, COL_开始时间) = .TextMatrix(i, COL_开始时间)
                        .Cell(flexcpData, i, COL_开嘱时间) = .TextMatrix(i, COL_开嘱时间)
                        .Cell(flexcpData, i, COL_上次执行) = .TextMatrix(i, COL_上次执行)
                        .Cell(flexcpData, i, COL_终止时间) = .TextMatrix(i, COL_终止时间)
                        .Cell(flexcpData, i, COL_校对时间) = .TextMatrix(i, COL_校对时间)
                        .Cell(flexcpData, i, COL_停嘱时间) = .TextMatrix(i, COL_停嘱时间)
                        .TextMatrix(i, COL_开始时间) = Format(.TextMatrix(i, COL_开始时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_开嘱时间) = Format(.TextMatrix(i, COL_开嘱时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_上次执行) = Format(.TextMatrix(i, COL_上次执行), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_终止时间) = Format(.TextMatrix(i, COL_终止时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_校对时间) = Format(.TextMatrix(i, COL_校对时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_停嘱时间) = Format(.TextMatrix(i, COL_停嘱时间), "yyyy-MM-dd HH:mm")
                        
                        If (mint类型 = 1 Or mint类型 = 7) Then
                            '停嘱时缺省的医嘱终止时间
                            If mdateStop <> CDate(0) And (.TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" Or .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0) Then
                                .TextMatrix(i, COL_输入) = Format(vCurDate - 1 / 24 / 60, "yyyy-MM-dd HH:mm")
                            ElseIf optStop(e上次执行时间).value Then
                                '末次时间如果为空，则为开始时间
                                If .TextMatrix(i, COL_频率) <> "持续性" Then
                                    If .TextMatrix(i, COL_上次执行) = "" Then
                                        .TextMatrix(i, COL_输入) = CStr(.Cell(flexcpData, i, COL_开始时间))
                                    Else
                                        .TextMatrix(i, COL_输入) = CStr(.Cell(flexcpData, i, COL_上次执行))
                                    End If
                                Else
                                    '当前时间或指定时间
                                    .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                '当前时间或指定时间
                                .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                                                
                                If .TextMatrix(i, COL_频率) <> "持续性" Then
                                    If chkNoSend.value = 0 Then      '如果不补发
                                        If .TextMatrix(i, COL_上次执行) = "" Then
                                            .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_开始时间)
                                        ElseIf .TextMatrix(i, COL_上次执行) < .TextMatrix(i, COL_输入) Then
                                            .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_上次执行)
                                        End If
                                    End If
                                    If chkRollSend.value = 0 Then   '如果不收回
                                        If .TextMatrix(i, COL_上次执行) > .TextMatrix(i, COL_输入) Then
                                            .TextMatrix(i, COL_输入) = .TextMatrix(i, COL_上次执行)
                                        End If
                                    End If
                                End If
                            End If
                            If mint类型 = 7 Then
                                '审核的时候默认时间为申请停止的时间（病人医嘱状态.操作说明）
                                If rsTmp!操作说明 & "" <> "" Then
                                    strTmp = rsTmp!操作说明 & "<T>"
                                    .TextMatrix(i, COL_输入) = Format(Split(strTmp, "<T>")(0), "yyyy-MM-dd HH:mm")
                                    .TextMatrix(i, COL_终止原因) = Split(strTmp, "<T>")(1)
                                    .Cell(flexcpData, i, COL_终止原因) = Split(strTmp, "<T>")(1)
                                End If
                            Else
                                '停嘱原因
                                .TextMatrix(i, COL_终止原因) = mstr停嘱原因
                                .Cell(flexcpData, i, COL_终止原因) = mstr停嘱原因
                            End If
                            
                            .TextMatrix(i, COL_输入) = GetValidateStopDate(.TextMatrix(i, COL_输入), i)

                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 2 Then
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            '应>=终止时间
                            If .TextMatrix(i, COL_输入) < .Cell(flexcpData, i, COL_终止时间) Then
                                .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_终止时间))), "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 3 Then
                            '校对时的缺省校对时间
                             If optOper(e当前时间).value Then
                                If .TextMatrix(i, COL_标志) = "补录" Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_开嘱时间))), "yyyy-MM-dd HH:mm")
                                Else
                                    .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, IIF(cboTime(e早于).ListIndex = 0, COL_开始时间, COL_开嘱时间))
                                Else
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, IIF(cboTime(e晚于).ListIndex = 0, COL_开始时间, COL_开嘱时间))
                                End If
                            End If
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                                                                                
                            '疑问修改后
                            If Val("" & rsTmp!疑问更正) = 1 Then
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgFlag.ListImages("M").Picture
                                int图标数 = 1
                            End If
                            
                            'Pass:根据审查结果显示警示灯
                            '只有护士校对的时候传入了映射表格对象并显示警示列,其他（停止、确认停止...）未传人映射对象不用设置警示值,
                            If mblnPass Then
                                If .TextMatrix(i, COL_警示) <> "" Then
                                    Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_警示))) '用于单药警告
                                    .TextMatrix(i, COL_警示) = ""
                                End If
                            End If
                            
                        ElseIf mint类型 = 5 Then
                            If mblnPauseLast Then
                                If .TextMatrix(i, COL_上次执行) <> "" Then
                                    '缺省在上次执行时间之后暂停
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_上次执行))), "yyyy-MM-dd HH:mm")
                                Else
                                    '如无上次执行时间则以开始时间为准
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_开始时间)
                                End If
                            Else
                                '暂停医嘱时间:暂停段中,医嘱暂停点无效,启用点有效。
                                .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            
                            '应>=开始执行时间,因为该时间点尚未执行
                            If .TextMatrix(i, COL_输入) < .Cell(flexcpData, i, COL_开始时间) Then
                                .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_开始时间)
                            End If
                            '应>上次执行时间,因为该时间点已执行
                            If .TextMatrix(i, COL_上次执行) <> "" Then
                                If .TextMatrix(i, COL_输入) <= .Cell(flexcpData, i, COL_上次执行) Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_上次执行))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            '应<执行终止时间,因为该时间点执行有效
                            If .TextMatrix(i, COL_终止时间) <> "" Then
                                If .TextMatrix(i, COL_输入) >= .Cell(flexcpData, i, COL_终止时间) Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", -1, CDate(.Cell(flexcpData, i, COL_终止时间))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                            rsPause.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_输入) <= Format(rsPause!上次时间, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, rsPause!上次时间), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 6 Then
                            '启用医嘱时间
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            
                            '应>暂停时间
                            rsPause.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_输入) <= Format(rsPause!上次时间, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, rsPause!上次时间), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            '应<=执行终止时间
                            If .TextMatrix(i, COL_终止时间) <> "" Then
                                If .TextMatrix(i, COL_输入) > .Cell(flexcpData, i, COL_终止时间) Then
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_终止时间)
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        End If
                        
                        '行高
                        If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                        
                        '毒麻精药品标识
                        If .TextMatrix(i, COL_毒理分类) <> "" Then
                            If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                                .Cell(flexcpFontBold, i, col_医嘱内容) = True
                            End If
                        End If
                        
                        '皮试结果标识
                        If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1" And .TextMatrix(i, COL_皮试) <> "" Then
                            j = GetSkinTestResult(Val(.TextMatrix(i, COL_诊疗项目ID)), .TextMatrix(i, COL_皮试))
                            .Cell(flexcpForeColor, i, COL_皮试) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_皮试))
                        End If
                        
                        '电子签名标识
                        If Val(.TextMatrix(i, COL_签名ID)) <> 0 Then
                            Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgSign.ListImages("签名").Picture
                            int图标数 = 1
                        End If
                        
                        If Val(rsTmp!高危药品 & "") > 0 Then
                            If .Cell(flexcpPicture, i, col_医嘱内容) Is Nothing Then
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
                                int图标数 = 1
                            Else
                                If .Cell(flexcpPicture, i, col_医嘱内容) <> frmIcons.imgQuestion.ListImages("高危药品").Picture Then
                                    pictmp.Cls
                                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("高危药品").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                    Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                    int图标数 = 2
                                End If
                            End If
                        End If
         
                        If Val(rsTmp!处方审查结果 & "") = 2 Then '处方审核未通过，加图标
                            .TextMatrix(i, COL_选择) = 1
                            If int图标数 = 0 Then
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                            ElseIf int图标数 = 1 Then
                                pictmp.Cls
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgFlag.ListImages("审核未通过").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                int图标数 = 2
                            ElseIf int图标数 = 2 Then
                                pictmp.Cls
                                pictmp.Width = 720
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, 480, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgFlag.ListImages("审核未通过").Picture, 480, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                pictmp.Width = 480
                                int图标数 = 3
                            End If
                        End If
                        
                        '易跌倒图标
                        If Val(rsTmp!是否易至跌倒 & "") > 0 Then
                            If int图标数 = 0 Then
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgQuestion.ListImages("易跌倒").Picture
                            ElseIf int图标数 = 1 Then
                                pictmp.Cls
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                Set vsAdvice.Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                int图标数 = 2
                            ElseIf int图标数 = 2 Then
                                pictmp.Cls
                                pictmp.Width = 720
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, 480, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 480, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                pictmp.Width = 480
                                int图标数 = 3
                            ElseIf int图标数 = 3 Then
                                pictmp.Cls
                                pictmp.Width = 960
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, 720, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 720, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                                pictmp.Width = 480
                                int图标数 = 4
                            End If
                        End If
                        
                    End If
                    
                    If bln给药途径 Or bln输血途径 Then .RemoveItem i
                    
                    Progress = rsTmp.AbsolutePosition / rsTmp.RecordCount * 100
                    
                    rsTmp.MoveNext
                Loop
            End With
        End If
    Next
        
    '----------------------------------------------------------------------
    '病人信息显示
    If strPatis <> "" Then
        strPatis = Mid(strPatis, 2)
    End If
    If UBound(Split(strPatis, ";")) = 0 Then
        '只有一个病人的数据的情况
        lng病人ID = Split(strPatis, ",")(0)
        lng主页ID = Split(strPatis, ",")(1)
        If lng病人ID <> mlng病人ID Or fraBaby.Visible Then '不是当前病人重新取来显示
            strSQL = _
                " Select B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,B.出院病床," & _
                " B.住院医师,B.出院科室ID,C.名称 as 科室" & _
                " From 病人信息 A,病案主页 B,部门表 C" & _
                " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                " And B.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
            lblPati.Caption = "姓名:" & rsTmp!姓名 & "　住院号:" & NVL(rsTmp!住院号) & _
                "　床号:" & NVL(rsTmp!出院病床) & "　科室:" & NVL(rsTmp!科室)
        End If
    ElseIf UBound(Split(strPatis, ";")) > 0 Then
        '有多个病人数据的情况
        vsAdvice.ColHidden(COL_姓名) = False
        vsAdvice.ColHidden(COL_住院号) = False
        vsAdvice.ColHidden(COL_床号) = False
                
        strSQL = "Select 名称 From 部门表 Where ID IN(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(str科室s, 2))
        str科室s = ""
        Do While Not rsTmp.EOF
            str科室s = str科室s & "," & rsTmp!名称
            rsTmp.MoveNext
        Loop
        lblPati.Caption = "(" & Mid(str科室s, 2) & ")共有 " & UBound(Split(strPatis, ";")) + 1 & " 个病人的医嘱"
    ElseIf UBound(Split(strPatis, ";")) = -1 Then
        '没有任何病人数据的情况
        lblPati.Caption = "当前条件下没有任何相关信息。"
    End If
    
    '----------------------------------------------------------------------
    If vsAdvice.Rows = vsAdvice.FixedRows Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        vsPrice.Rows = vsPrice.FixedRows
        vsPrice.Rows = vsPrice.FixedRows + 1
    Else
        '电子签名图标对齐
        vsAdvice.Cell(flexcpPictureAlignment, vsAdvice.FixedRows, col_医嘱内容, vsAdvice.Rows - 1, col_医嘱内容) = 0
        '自动调整行高
        vsAdvice.AutoSize col_医嘱内容
    End If
    Call SetAdviceCol
    vsAdvice.Row = vsAdvice.FixedRows
    If Not vsAdvice.ColHidden(COL_选择) Then
        vsAdvice.Col = COL_选择
    Else
        vsAdvice.Col = col_医嘱内容
    End If
    vsAdvice.Redraw = flexRDDirect
    
    Progress = 0: Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    strPatis = ""
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Function CheckValid(Optional ByRef bln超期收回 As Boolean) As Boolean
'功能：确认前检查合法性
'参数：bln超期收回=确认停止时传回是否存在需要超期收回的医嘱
    Dim str超期 As String, str超长 As String
    Dim str特殊 As String, strTmp As String
    Dim curDate As Date, i As Long, k As Long
    Dim strPatis As String, strMsg As String
    Dim rsDrug As ADODB.Recordset, strUnRoll As String, lng药品ID As Long, blnDo As Boolean
    Dim lng医嘱ID As Long
    Dim strTmpTim As String
    Dim str执行登记 As String
    Dim lng相关ID As Long
    Dim str终止原因 As String
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    mstrRollNotify = ""
    curDate = zlDatabase.Currentdate
    strUnRoll = zlDatabase.GetPara("发药后不收回", glngSys, p住院医嘱发送)
    
    With vsAdvice
        '是否有可以操作的记录
        If .Rows = .FixedRows + 1 And Val(.TextMatrix(.FixedRows, COL_ID)) = 0 Then
            If mint类型 = 0 Then
                '医嘱作废
                strTmp = "当前没有可以作废的医嘱。"
            ElseIf mint类型 = 1 Then
                '停止医嘱
                strTmp = "当前没有可以停止的医嘱。"
            ElseIf mint类型 = 2 Then
                '确认停止
                strTmp = "当前没有被停止的医嘱。"
            ElseIf mint类型 = 3 Then
                '医嘱校对
                strTmp = "当前没有新开的医嘱。"
            ElseIf mint类型 = 4 Then
                '调整计价项目
                strTmp = "当前没有通过校对的有效医嘱。"
            ElseIf mint类型 = 5 Then
                '暂停医嘱
                strTmp = "当前没有可以暂停的医嘱。"
            ElseIf mint类型 = 6 Then
                '启用医嘱
                strTmp = "当前没有暂停后需要启用的医嘱。"
            ElseIf mint类型 = 7 Then
                '停嘱审核
                strTmp = "当前没有可以审核的停嘱。"
            End If
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            Exit Function
        End If
        
        On Error GoTo errH
        '是否有选择
        str超期 = "": str超长 = "": str特殊 = ""
        If Not .ColHidden(COL_选择) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And (Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty) Then
                    k = k + 1
                    If InStr(strPatis & ",", "," & .TextMatrix(i, COL_病人ID)) = 0 Then
                        strPatis = strPatis & "," & .TextMatrix(i, COL_病人ID)
                    End If
                    
                    If (mint类型 = 1 Or mint类型 = 7) Then
                        '停医嘱时，检查执行登记情况
                        If mint类型 = 1 And Not (.TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",4,14,9,10,12,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0) Then
                            lng医嘱ID = IIF(InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0, .TextMatrix(i, COL_相关ID), .TextMatrix(i, COL_ID))
                            If lng医嘱ID <> lng相关ID Then strTmpTim = GetAdviceStopTime(lng医嘱ID)
                            lng相关ID = lng医嘱ID
                            If IsDate(strTmpTim) Then
                                If .TextMatrix(i, COL_执行时间) = "" _
                                    And (Val(.TextMatrix(i, COL_频率次数)) = 0 Or Val(.TextMatrix(i, COL_频率间隔)) = 0 Or .TextMatrix(i, COL_间隔单位) = "") Then
                                    '"持续性"长嘱,按天算
                                    If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd") < Format(CDate(strTmpTim), "yyyy-MM-dd") Then
                                        str执行登记 = str执行登记 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容) & " 执行时间：" & Format(CDate(strTmpTim), "yyyy-MM-dd")
                                    End If
                                Else
                                    If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") < Format(strTmpTim, "yyyy-MM-dd HH:mm") Then
                                        str执行登记 = str执行登记 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容) & " 执行时间：" & Format(CDate(strTmpTim), "yyyy-MM-dd  HH:mm")
                                    End If
                                End If
                            End If
                        End If
                         
                        '收集超期发送的医嘱(排开备用医嘱）
                        If IsDate(.Cell(flexcpData, i, COL_上次执行)) And .TextMatrix(i, COL_频率) <> "必要时" And .TextMatrix(i, COL_频率) <> "需要时" Then
                            If .TextMatrix(i, COL_执行时间) = "" _
                                And (Val(.TextMatrix(i, COL_频率次数)) = 0 Or Val(.TextMatrix(i, COL_频率间隔)) = 0 Or .TextMatrix(i, COL_间隔单位) = "") Then
                                '"持续性"长嘱,按天算
                                If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_上次执行)), "yyyy-MM-dd") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容)
                                End If
                            Else
                                If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                                    '检查无需收回的药品
                                    blnDo = True
                                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 And strUnRoll <> "" Then
                                        If .TextMatrix(i, COL_收费细目ID) <> "" Then
                                            lng药品ID = Val(.TextMatrix(i, COL_收费细目ID))
                                        Else
                                            lng药品ID = GetLastSendMediCineID(Val(.TextMatrix(i, COL_ID)), CDate(.Cell(flexcpData, i, COL_上次执行)), Val(.TextMatrix(i, COL_病人性质)))
                                        End If
                                        If lng药品ID <> 0 Then
                                            gstrSQL = "Select 发药类型 From 药品规格 Where 药品ID = [1] And 发药类型 is Not Null"
                                            Set rsDrug = zlDatabase.OpenSQLRecord(gstrSQL, "超期收回检查", lng药品ID)
                                            If rsDrug.RecordCount > 0 Then
                                                If InStr("," & strUnRoll & ",", "," & rsDrug!发药类型 & ",") > 0 Then
                                                    If CheckMedicineSended(Val(.TextMatrix(i, COL_ID)), CDate(.Cell(flexcpData, i, COL_上次执行))) Then
                                                        blnDo = False
                                                    End If
                                                End If
                                            End If
                                        Else '无需收回：医嘱未记费（如自备药）或相关费用被删除了（如划价单被删除）
                                            blnDo = False
                                        End If
                                    End If
                                    If blnDo Then
                                        str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容)
                                    End If
                                End If
                            End If
                        End If
                        
                        '收集超长停止的医嘱
                        If CDate(.TextMatrix(i, COL_输入)) - curDate > 7 Then
                            str超长 = str超长 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容) & "，停止时间：" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm")
                        End If
                        
                        '未填写终止原因
                        
                        If gbln医嘱终止原因 And InStr(gstr可不填停嘱原因科室, "," & Val(.TextMatrix(i, COL_病人科室ID)) & ",") = 0 Then
                            If .TextMatrix(i, COL_终止原因) = "" Then
                                .Row = i: .ShowCell .Row, COL_终止原因
                                MsgBox "该医嘱未录入终止原因。", vbInformation, gstrSysName
                                Exit Function
                            Else
                                If zlCommFun.ActualLen(.TextMatrix(i, COL_终止原因)) > txt疑问.MaxLength Then
                                    .Row = i: .ShowCell .Row, COL_终止原因
                                    MsgBox "该医嘱终止原因内容太长，最多允许 " & txt疑问.MaxLength / 2 & " 个汉字或 " & txt疑问.MaxLength & " 个字符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    ElseIf mint类型 = 2 Then
                        '收集超期发送的医嘱
                        If IsDate(.Cell(flexcpData, i, COL_上次执行)) And .TextMatrix(i, COL_频率) <> "必要时" And .TextMatrix(i, COL_频率) <> "需要时" Then
                            If .TextMatrix(i, COL_执行时间) = "" _
                                And (Val(.TextMatrix(i, COL_频率次数)) = 0 Or Val(.TextMatrix(i, COL_频率间隔)) = 0 Or .TextMatrix(i, COL_间隔单位) = "") Then
                                '"持续性"长嘱,按天算
                                If Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_上次执行)), "yyyy-MM-dd") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容)
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, col_医嘱内容)
                                End If
                            End If
                        End If
                    ElseIf mint类型 = 3 Then
                        '收集特殊性医嘱,通过校对的才判断
                        '3-转科;4-术后;5-出院;6-转院,11-死亡,14-术前
                        If .Cell(flexcpData, i, COL_选择) = 1 And _
                            .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,4,5,6,11,14,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                            
                            If InStr(str特殊 & ",", "," & .TextMatrix(i, COL_病人ID) & ":" & .TextMatrix(i, COL_主页ID) & ",") = 0 Then
                                str特殊 = str特殊 & "," & .TextMatrix(i, COL_病人ID) & ":" & .TextMatrix(i, COL_主页ID)
                            End If
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, COL_姓名) & _
                                IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容)
                            
                            '转科医嘱检查
                            If Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                                If CheckCanSendAdvice(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)), Val(.TextMatrix(i, COL_ID)), Val(.Cell(flexcpData, i, COL_婴儿))) Then
                                    Call MsgBox("发现转科医嘱：" & vbCrLf & .TextMatrix(i, COL_姓名) & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & "必须将可以发送的长期医嘱处理后才能校对。", vbInformation, gstrSysName)
                                    Exit Function
                                End If
                            End If
                        ElseIf .Cell(flexcpData, i, COL_选择) = 2 Then
                            If zlCommFun.ActualLen(.TextMatrix(i, COL_操作说明)) > txt疑问.MaxLength Then
                                .Row = i: .ShowCell .Row, COL_选择
                                MsgBox "该医嘱校对疑问的说明内容太长，最多允许 " & txt疑问.MaxLength / 2 & " 个汉字或 " & txt疑问.MaxLength & " 个字符。", vbInformation, gstrSysName
                                If txt疑问.Visible Then txt疑问.SetFocus
                                Exit Function
                            End If
                        End If
                        
                        '护理等级医嘱的判断
                        If .TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" Then
                            If Check护理等级变动交叉(.TextMatrix(i, COL_病人ID), .TextMatrix(i, COL_主页ID), .Cell(flexcpData, i, COL_开始时间)) Then
                                Exit Function
                            End If
                        End If
                    ElseIf mint类型 = 0 Then
                        If mbln叮嘱发送执行 Then
                            If .TextMatrix(i, COL_诊疗项目ID) = "" Then
                                strSQL = "select nvl(max(decode(执行结果,1,1,0)),0) as 执行状态 from 病人医嘱执行 where 医嘱id=[1]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "医嘱执行检查", .TextMatrix(i, COL_ID))
                                
                                If Not rsTmp.EOF Then
                                    If InStr(",1,3,", NVL(rsTmp!执行状态, 0)) > 0 Then
                                        MsgBox "自由录入医嘱[" & .TextMatrix(i, col_医嘱内容) & "]已经执行,不能作废！", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If k = 0 Then
                MsgBox "没有选择任何医嘱，请选择需要" & tbr.Buttons("执行").Caption & "的医嘱。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
                
        '医生
        If (mint类型 = 1 Or mint类型 = 7) And mbln护士站 Then
            If cbo医生.ListIndex = -1 Then
                MsgBox "请选择停止医嘱的医生。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    strTmp = ""
    strPatis = IIF(UBound(Split(Mid(strPatis, 2), ",")) > 0, "你选择了多个病人的医嘱，请仔细进行检查以避免出现差错。" & vbCrLf & vbCrLf, "")
    If mint类型 = 0 Then
        '医嘱作废
        strTmp = "确实要作废已经选择的医嘱吗？"
    ElseIf (mint类型 = 1 Or mint类型 = 7) Then
        If str执行登记 <> "" Then
            MsgBox "下列医嘱被填写了执行登记：" & vbCrLf & str执行登记 & _
                vbCrLf & vbCrLf & "请修改停止时间或取消执行登记。", vbInformation, gstrSysName
            Exit Function
        End If
        '停止医嘱
        If str超期 <> "" Then '检查是否有需要退回超前摆药的情况
            If MsgBox("下列医嘱被超期发送：" & vbCrLf & str超期 & _
                vbCrLf & vbCrLf & "在停止确认后可以使用""超期发送收回""进行处理。" & _
                vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str超长 <> "" Then
            If MsgBox("下列医嘱的停止时间超过当前时间太久：" & vbCrLf & str超长 & _
                vbCrLf & vbCrLf & "如果停止时间不正确，将会对医嘱的发送和计费产生影响。" & _
                vbCrLf & "确实要在指定的时间停止这些医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str超期 = "" And str超长 = "" Then
            strTmp = "确实要" & IIF(mint类型 = 7, "审核", "停止") & "已经选择的医嘱吗？"
        End If
    ElseIf mint类型 = 2 Then
        '确认停止
        If str超期 <> "" And InStr(GetInsidePrivs(p住院医嘱发送), ";超期发送收回;") > 0 Then
            bln超期收回 = True
        Else
            strTmp = "确认已经选择的医嘱停止吗？"
        End If
    ElseIf mint类型 = 3 Then
        '医嘱校对
        If strMsg <> "" Then
            mstrRollNotify = Mid(str特殊, 2)
            
            '如果启用了电子签名，检查存在"已停止但未确认停止"的医嘱，提示护士先进行确认停止
            '因为特殊医嘱校对时会将"已停止但未确认停止"的医嘱的"执行终止时间"调整为特殊医嘱的开始执行时间，医嘱停止的签名源文包含了"执行终止时间"，这会导致签名验证无法通过
            If Mid(gstrESign, 2, 1) = "1" Then  '住院医生站启用了电子签名才检查
                strTmp = ""
                '在判断时，排除未启用签名科室下达的医嘱
                If CheckStopedUnAffirm(mstrRollNotify, strTmp) Then
                    MsgBox "要校对的医嘱中包含下列特殊医嘱：" & vbCrLf & strMsg & _
                        vbCrLf & vbCrLf & "校对后会将未确认停止的医嘱重新停止，为了不影响签名验证，请先对以下病人进行确认停止操作：" & strTmp, vbInformation, gstrSysName
                    Exit Function
                End If
                strTmp = ""
            End If
            
            
            If MsgBox(strPatis & "要校对的医嘱中包含下列特殊医嘱：" & vbCrLf & strMsg & _
                vbCrLf & vbCrLf & "这些医嘱校对后会停止其它长期医嘱，确实要校对当前选择的医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            strTmp = strPatis & "确实要对已经选择的医嘱进行校对处理吗？"
        End If
    ElseIf mint类型 = 5 Then
        '暂停医嘱
        strTmp = strPatis & "确实要暂停已经选择的医嘱吗？"
    ElseIf mint类型 = 6 Then
        '启用医嘱
        strTmp = strPatis & "确实要启用已经选择的医嘱吗？"
    End If
    If strTmp <> "" Then
        If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    CheckValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckSignValid() As Boolean
'功能：1.检查未签名的医嘱不能进行校对
'      2.一次签名的医嘱必须一起通过校对
    Dim col医嘱ID As New Collection, str医嘱ID As String
    Dim col签名ID As New Collection, str签名ID As String
    Dim str住院 As String, str医技 As String
    Dim lng签名id As Long, strTmp As String
    Dim int状态 As Integer, i As Long, j As Long
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    If mint类型 <> 3 Then CheckSignValid = True: Exit Function
    
    With vsAdvice
        '获取护士人员列表：只是护士，不是医生
        If Mid(gstrESign, 2, 1) = "1" Or Mid(gstrESign, 3, 1) = "1" Then
            strNurse = ""
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                    If .Cell(flexcpData, i, COL_选择) = 1 And Val(.TextMatrix(i, COL_签名ID)) = 0 Then
                        If InStr(strNurse & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") = 0 Then
                            strNurse = strNurse & "," & Val(.TextMatrix(i, COL_ID))
                        End If
                    End If
                End If
            Next
            If strNurse <> "" Then
                strSQL = "Select /*+ Rule*/" & vbNewLine & _
                    " a.姓名,b.医嘱ID" & vbNewLine & _
                    "From 人员表 A," & vbNewLine & _
                    "     (Select Distinct 操作人员,医嘱ID" & vbNewLine & _
                    "       From 病人医嘱状态" & vbNewLine & _
                    "       Where 医嘱id In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And 操作类型 = 1) B" & vbNewLine & _
                    "Where a.姓名 = b.操作人员 And Exists (Select 1 From 人员性质说明 X Where x.人员id = a.Id And x.人员性质 = '护士') And Not Exists" & vbNewLine & _
                    " (Select 1 From 人员性质说明 Y Where y.人员id = a.Id And y.人员性质 = '医生')" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strNurse, 2))
                On Error GoTo 0
                
                strNurse = ""
                Do While Not rsTmp.EOF
                    strNurse = strNurse & "," & rsTmp!医嘱ID
                    rsTmp.MoveNext
                Loop
                strNurse = strNurse & ","
            End If
        End If
        
        For i = .FixedRows To .Rows - 1
            'flexcpData:0-不处理,1-校对,2-疑问
            If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                '1.收集未签名的医嘱内容
                If .Cell(flexcpData, i, COL_选择) = 1 And Val(.TextMatrix(i, COL_签名ID)) = 0 Then
                    '设置为使用签名的场合
                    If InStr(strNurse, "," & Val(.TextMatrix(i, COL_ID)) & ",") = 0 Then '护士录入的医嘱不进行签名检查
                        If Val(.TextMatrix(i, COL_前提ID)) = 0 Then
                            If CheckSign(1, Val(.TextMatrix(i, COL_开嘱科室ID)), , , , , gobjESign, .TextMatrix(i, COL_开嘱医生)) Then
                                If UBound(Split(str住院, vbCrLf)) < 10 Then
                                    str住院 = str住院 & vbCrLf & "●" & .TextMatrix(i, col_医嘱内容)
                                ElseIf InStr(str住院, "… …") = 0 Then
                                    str住院 = str住院 & vbCrLf & "… …"
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_前提ID)) <> 0 Then
                            If CheckSign(3, Val(.TextMatrix(i, COL_开嘱科室ID)), , , , , gobjESign, .TextMatrix(i, COL_开嘱医生)) Then
                                If UBound(Split(str医技, vbCrLf)) < 10 Then
                                    str医技 = str医技 & vbCrLf & "●" & .TextMatrix(i, col_医嘱内容)
                                ElseIf InStr(str医技, "… …") = 0 Then
                                    str医技 = str医技 & vbCrLf & "… …"
                                End If
                            End If
                        End If
                    End If
                End If
                
                '2.收集已签名医嘱的校对状态
                lng签名id = Val(.TextMatrix(i, COL_签名ID))
                If lng签名id <> 0 Then
                    j = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID))) '组ID
                    int状态 = .Cell(flexcpData, i, COL_选择)
                    If int状态 = 2 Then int状态 = 0 '这里疑问等同于不校对
                    If InStr(str签名ID & ",", "," & lng签名id & ",") > 0 Then
                        '收集各个签名在界面上的校对状态
                        strTmp = Split(col签名ID("_" & lng签名id), "=")(1)
                        If InStr(strTmp, int状态) = 0 Then
                            col签名ID.Remove "_" & lng签名id
                            col签名ID.Add lng签名id & "=" & strTmp & int状态, "_" & lng签名id
                        End If
                        
                        '收集各个签名已读到界面的医嘱(组ID)
                        strTmp = col医嘱ID("_" & lng签名id)
                        If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                            col医嘱ID.Remove "_" & lng签名id
                            col医嘱ID.Add strTmp & "," & j, "_" & lng签名id
                        End If
                    Else
                        str签名ID = str签名ID & "," & lng签名id
                        col签名ID.Add lng签名id & "=" & int状态, "_" & lng签名id
                        col医嘱ID.Add j, "_" & lng签名id
                    End If
                End If
            End If
        Next
        
        '检查已签名医嘱校对情况
        strTmp = "": str医嘱ID = Mid(str医嘱ID, 2)
        For i = 1 To col签名ID.Count
            lng签名id = Split(col签名ID(i), "=")(0)
            str签名ID = Split(col签名ID(i), "=")(1)
            
            '本次一起签名的未读入界面的未校对医嘱
            str医嘱ID = col医嘱ID("_" & lng签名id)
            str医嘱ID = ExistOtherSignAdvice(lng签名id, str医嘱ID)
            If str医嘱ID <> "" Then
                If InStr(str签名ID, "0") = 0 Then
                    str签名ID = str签名ID & "0"
                    strTmp = strTmp & str医嘱ID
                End If
            End If
            
            If Not (str签名ID = "1" Or str签名ID = "0") Then
                '这次签名的内容不是"都要通过校对或都不通过校对(包括疑问)"的情况
                j = .FindRow(CStr(lng签名id), , COL_签名ID)
                Do While j <> -1
                    If Val(.TextMatrix(j, COL_ID)) <> 0 And Not .RowHidden(j) Then
                        If InStr(",0,2,", .Cell(flexcpData, j, COL_选择)) > 0 Then
                            strTmp = strTmp & vbCrLf & .TextMatrix(j, COL_姓名) & "：" & IIF(Len(.TextMatrix(j, col_医嘱内容)) > 40, Left(.TextMatrix(j, col_医嘱内容), 40) & "...", .TextMatrix(j, col_医嘱内容))
                        End If
                    End If
                    j = .FindRow(CStr(lng签名id), j + 1, COL_签名ID)
                Loop
                Exit For '暂只提示第一组
            End If
        Next
    End With
    
    '1.没有签名的医嘱不允许校对：对住院医嘱和医技医嘱分别进行检查
    If str住院 <> "" Then
        MsgBox "以下医嘱医生还没有签名，不能进行校对：" & vbCrLf & str住院, vbInformation, gstrSysName
        Exit Function
    End If
    If str医技 <> "" Then
        MsgBox "以下医嘱医生还没有签名，不能进行校对：" & vbCrLf & str医技, vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.一起签名的医嘱必须一起通过校对
    If strTmp <> "" Then
        MsgBox "以下医嘱与其他本次要通过校对的医嘱一起签名，但当前处理为不校对或校对疑问：" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "一起签名的医嘱必须一起通过校对，请调整相关医嘱的校对状态。", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckSignValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExistOtherSignAdvice(ByVal lng签名id As Long, ByVal str医嘱ID As String) As String
'功能：检查是否存在某次新开医嘱签名中本次没有读取到界面上的医嘱(因为要一起通过校对,如果有,这些医嘱也是没校对的)
'返回：未读取到界面的未校对医嘱内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.姓名,B.医嘱内容 From 病人医嘱状态 A,病人医嘱记录 B" & _
        " Where A.医嘱ID=B.ID And A.操作类型=1 And B.医嘱状态 IN(1,2)" & _
        " And (B.相关ID is Null Or B.诊疗类别 IN('5','6'))" & _
        " And Not Exists(Select 1 From 病人医嘱记录 S Where 诊疗类别 IN('5','6') And S.相关ID=B.ID)" & _
        " And Instr([2],','||Nvl(B.相关ID,B.ID)||',')=0 And A.签名ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng签名id, "," & str医嘱ID & ",")
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & vbCrLf & NVL(rsTmp!姓名) & "：" & IIF(Len(NVL(rsTmp!医嘱内容)) > 40, Left(NVL(rsTmp!医嘱内容), 40) & "...", NVL(rsTmp!医嘱内容))
        rsTmp.MoveNext
    Loop
    ExistOtherSignAdvice = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng医嘱ID As Long, int费用性质 As Integer, ByVal lng项目id As Long, ByVal lngCol As Long)
'功能：定位到并显示指定医嘱的指定计价行
'参数：lngRow=医嘱行号,lng医嘱ID=计价医嘱ID
'      lng项目ID=计价项目ID,lngCol=计价表格显示列
    Dim k As Long
    
    With vsAdvice
        .Row = lngRow: .Col = col_医嘱内容 '进入行自动ShowPrice,mrsPrice发生变化
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_医嘱ID)) = lng医嘱ID _
                And Val(vsPrice.TextMatrix(k, COLP_费用性质)) = int费用性质 _
                And Val(vsPrice.TextMatrix(k, COLP_收费细目ID)) = lng项目id Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function ExecuteOperate() As Boolean
    Dim arrSQL As Variant, lng相关ID As Long, blnTrans As Boolean
    Dim blnExe As Boolean, i As Long, j As Long
    Dim lng医嘱ID As Long, lng执行科室ID As Long
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名id As Long, lng证书ID As Long
    Dim strSource As String, strSign As String
    Dim strOper As String, strTimeStamp As String, strTimeStampCode As String
    Dim colSomeTime As New Collection
    Dim rsAdviceTmp As ADODB.Recordset
    Dim strAdvicesTmp As String
    Dim lngPatiID As Long
    Dim lngPageId As Long
    Dim lngLastRow As Long    '上一次勾选的行数
    Dim strRevokeIDs As String, arrRevokeID() As String
    Dim str配液提示 As String, str配液禁止 As String, str配液打包 As String
    Dim strSQL As String, rsTmp As Recordset
    Dim lng护理等级医嘱id As Long '除开本次作废的护理等级医嘱外的最近的自动停止的护理等级医嘱id
    Dim blnPrintBeforeRedo As Boolean '是否存在已经打印过的医嘱打印时间是在最后重整操作之前
    Dim rsMsgRow As ADODB.Recordset
    Dim lngTmp As Long
    Dim strTmp As String
    Dim int紧急 As Integer '本次停止的医嘱是否是紧急医嘱
    Dim str作废输血提示 As String
    Dim str给药IDs As String, varTmp As Variant
    
    Dim strMsg As String
    Dim strPrintDel As String
    Dim arrPrintDel As Variant
    Dim lngLastPatiID As Long
    Dim lngLastPageID As Long
    Dim lngLastPatiDeptID As Long
    Dim rs输血 As ADODB.Recordset
    Dim bln输血 As Boolean
    Dim strErr As String
    
    Screen.MousePointer = 11
    
    mstrPatiKeepMsg = ""
    
    Call InitRecordSet(rsAdviceTmp, rsMsgRow, rs输血)
    
    '产生SQL
    arrSQL = Array()
    With vsAdvice
        If mint类型 = 3 Then
            If InitObjRecipeAudit(p住院医嘱下达) Then
                '处方审查系统产生待审数据
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty Then
                        If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Then
                            If lngLastPatiID <> Val(.TextMatrix(i, COL_病人ID)) Then
                                If Mid(str给药IDs, 2) <> "" Then
                                    Call gobjRecipeAudit.BuildData(Mid(str给药IDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                                    str给药IDs = ""
                                End If
                            End If
                            lngLastPatiID = Val(.TextMatrix(i, COL_病人ID))
                            lngLastPageID = Val(.TextMatrix(i, COL_主页ID))
                            lngLastPatiDeptID = Val(.TextMatrix(i, COL_病人科室ID))
                            If InStr("," & str给药IDs & ",", "," & .TextMatrix(i, COL_相关ID) & ",") = 0 Then str给药IDs = str给药IDs & "," & .TextMatrix(i, COL_相关ID)
                        End If
                    End If
                Next
                If Mid(str给药IDs, 2) <> "" Then
                    Call gobjRecipeAudit.BuildData(Mid(str给药IDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                End If
            End If
            '调用中联合理用药审方结果判断
            Call Check处方审查
        End If
        If mint类型 <> 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty Then
                    '一组医嘱只校对一次,除一并给药外,其它医嘱只有一个显示行
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> lng相关ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    If blnExe Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        '(组ID)使用相关ID为NULL的医嘱的ID(给药途径,中药用法,检查项目,主要手术,及独立医嘱)
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            lng医嘱ID = Val(.TextMatrix(i, COL_相关ID))
                        Else
                            lng医嘱ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If mint类型 = 0 Then      '医嘱作废
                            '医生作废医嘱电子签名
                            If Val(.TextMatrix(i, COL_签名ID)) <> 0 And CheckSign(IIF(mbln护士站, 2, 1), mlng医护科室ID, , , , , gobjESign) Then
                                str医嘱ID = str医嘱ID & "," & lng医嘱ID
                            End If
                            '作废护理等级
                            If .TextMatrix(i, COL_诊疗类别) = "H" Then
                                lng护理等级医嘱id = Get病人护理等级医嘱id(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)), Val(.TextMatrix(i, COL_婴儿)), lng医嘱ID)
                            End If
                            
                            '92129:正在配血和完成配血的情况，医嘱不能作废
                            If .TextMatrix(i, COL_诊疗类别) = "K" And gbln血库系统 Then
                                If InStr(1, ",2,5,6,", "," & Val(.TextMatrix(i, COL_审核状态)) & ",") <> 0 Then
                                    On Error GoTo errH
                                    strSQL = "Select Nvl(执行分类,0) as 执行分类 from 病人医嘱记录 A, 诊疗项目目录 B  where A.相关ID  = [1] and A.诊疗项目ID = B.ID"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询诊疗项目的执行分类", lng医嘱ID)
                                    If rsTmp.RecordCount > 0 Then
                                        If Val(rsTmp!执行分类) = 0 Then
                                            str作废输血提示 = str作废输血提示 & vbCrLf & .TextMatrix(i, col_医嘱内容)
                                        End If
                                    End If
                                    On Error GoTo 0
                                End If
                                Call rs输血.AddNew(Array("医嘱ID", "类型"), Array(lng医嘱ID, 4)): bln输血 = True
                            End If
                            
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng护理等级医嘱id & ")"
                            strRevokeIDs = strRevokeIDs & "," & lng医嘱ID
                            rsMsgRow.Filter = "病人id=" & Val(.TextMatrix(i, COL_病人ID)) & " And 主页id=" & Val(.TextMatrix(i, COL_主页ID)) & " And 操作类型=2"
                            If rsMsgRow.EOF Then
                                rsMsgRow.AddNew
                                rsMsgRow!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                rsMsgRow!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                rsMsgRow!行号 = i
                                rsMsgRow!操作类型 = 2
                                rsMsgRow.Update
                            End If
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人危急值医嘱_Update(3,null," & lng医嘱ID & ")"   '删除危急值对应关系
                        ElseIf (mint类型 = 1 Or mint类型 = 7) Then '停止医嘱
                            If mblnHaveAudit Then
                                '医生停止医嘱电子签名
                                If Val(.TextMatrix(i, COL_签名ID)) <> 0 And CheckSign(IIF(mbln护士站, 2, 1), mlng医护科室ID, , , , , gobjESign) Then
                                    str医嘱ID = str医嘱ID & "," & lng医嘱ID
                                    '记录停止医嘱的执行终止时间：由于是在执行过程之前取签名源文,这时还未写入数据库
                                    colSomeTime.Add Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
                                End If
                            End If
                            '检查输液配液记录
                            '1、已配药和已发送的记录，如果未打包、且未销帐的，则不允许停止。
                            '2、已配药和已发送的记录，如果已打包、且未销帐的，允许停止，但是需要提示。
                            '3、如果是已摆药、但未配药的记录，允许停止但需要提示。
                            '4、如果是已经销帐的记录，允许停止，不提示。
                            If gstr输液配置中心 <> "" And (.TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6") Then
                                strSQL = "Select Max(Decode(Instr(',4,5,6,7,8,', ',' || B.操作类型 || ','), 0, Null, To_Char(A.执行时间, 'yyyy-MM-dd HH24:MI'))) As 允许执行时间," & _
                                    " Max(Decode(A.操作状态, 1, Null, To_Char(A.执行时间, 'yyyy-MM-dd HH24:MI'))) As 提示执行时间, Min(A.是否打包) As 打包" & _
                                    " From 输液配药记录 A,输液配药状态 B Where A.医嘱id = [1] and A.ID=B.配药ID And A.执行时间 > [2] and A.操作状态<>10"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, CDate(Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm")))
                                If rsTmp.RecordCount > 0 Then
                                    If rsTmp!允许执行时间 & "" <> "" Then
                                        If Val(rsTmp!打包 & "") = 1 Then
                                            str配液打包 = str配液打包 & vbCrLf & .TextMatrix(i, col_医嘱内容) & "（执行时间:" & rsTmp!允许执行时间 & "）"
                                        Else
                                            '已发送和配药的，禁止停止
                                            str配液禁止 = str配液禁止 & vbCrLf & .TextMatrix(i, col_医嘱内容) & "（执行时间:" & rsTmp!允许执行时间 & "）"
                                        End If
                                    End If
                                    If rsTmp!提示执行时间 & "" <> "" Then
                                        str配液提示 = str配液提示 & vbCrLf & .TextMatrix(i, col_医嘱内容) & "（执行时间:" & rsTmp!提示执行时间 & "）"
                                    End If
                                End If
                            End If
                            
                            strSQL = "ZL_病人医嘱记录_停止(" & lng医嘱ID & ",To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                 "'" & IIF(mbln护士站, zlCommFun.GetNeedName(cbo医生.Text), UserInfo.姓名) & "',0," & IIF(mblnHaveAudit, 1, 0) & "," & mlng停嘱审核 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'"
                            
                            If gbln医嘱终止原因 Then
                                strSQL = strSQL & ",'" & .TextMatrix(i, COL_终止原因) & "')"
                            Else
                                strSQL = strSQL & ")"
                            End If
                            
                            arrSQL(UBound(arrSQL)) = strSQL
                            
                            '跟主界面显示相关的医嘱停止
                            If .TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" _
                                Or .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 Then
                                mblnRefresh = True
                            End If
                            rsMsgRow.Filter = "病人id=" & Val(.TextMatrix(i, COL_病人ID)) & " And 主页id=" & Val(.TextMatrix(i, COL_主页ID)) & " And 操作类型=" & mint类型
                            If rsMsgRow.EOF Then
                                rsMsgRow.AddNew
                                rsMsgRow!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                rsMsgRow!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                rsMsgRow!行号 = i
                                rsMsgRow!操作类型 = mint类型
                                rsMsgRow.Update
                            End If
                            
                            If int紧急 = 0 And mint类型 = 1 And .TextMatrix(i, COL_标志) = "紧急" Then int紧急 = 1
                            
                        ElseIf mint类型 = 2 Then  '确认停止
                            '确认停止医嘱电子签名
                            If mbln护士签名 And Val(.TextMatrix(i, COL_签名ID)) <> 0 And CheckSign(2, mlng医护科室ID, , , , , gobjESign) Then
                                If InStr(mstr病人IDs, ";") > 0 Then
                                    If rsAdviceTmp.State = adStateClosed Then rsAdviceTmp.Open
                                    If Val(.TextMatrix(i, COL_病人ID)) <> 0 And .TextMatrix(i, COL_病人ID) & "|" & .TextMatrix(i, COL_主页ID) <> .TextMatrix(lngLastRow, COL_病人ID) & "|" & .TextMatrix(lngLastRow, COL_主页ID) Then
                                        rsAdviceTmp.AddNew
                                        rsAdviceTmp!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                        rsAdviceTmp!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                    End If
                                    rsAdviceTmp!医嘱ids = rsAdviceTmp!医嘱ids & "," & lng医嘱ID
                                    rsAdviceTmp.Update
                                    lngLastRow = i
                                End If
                                
                                str医嘱ID = str医嘱ID & "," & lng医嘱ID
                                '记录确认停止医嘱时间：由于是在执行过程之前取签名源文,这时还未写入数据库
                                colSomeTime.Add Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_确认停止(" & lng医嘱ID & "," & _
                            "To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.姓名 & "')"
                        ElseIf mint类型 = 3 Then  '医嘱校对
                            '护士校对医嘱电子签名，校对疑问不签名
                            If mbln护士签名 And CheckSign(2, mlng医护科室ID, , , , , gobjESign) And Val(.TextMatrix(i, COL_签名ID)) <> 0 And .Cell(flexcpData, i, COL_选择) = 1 Then
                                If InStr(mstr病人IDs, ";") > 0 Then
                                    If rsAdviceTmp.State = adStateClosed Then rsAdviceTmp.Open
                                    If Val(.TextMatrix(i, COL_病人ID)) <> 0 And .TextMatrix(i, COL_病人ID) & "|" & .TextMatrix(i, COL_主页ID) <> .TextMatrix(lngLastRow, COL_病人ID) & "|" & .TextMatrix(lngLastRow, COL_主页ID) Then
                                        rsAdviceTmp.AddNew
                                        rsAdviceTmp!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                        rsAdviceTmp!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                    End If
                                    rsAdviceTmp!医嘱ids = rsAdviceTmp!医嘱ids & "," & lng医嘱ID
                                    rsAdviceTmp.Update
                                    lngLastRow = i
                                End If
                                
                                str医嘱ID = str医嘱ID & "," & lng医嘱ID
                                '记录校对医嘱时间：由于是在执行过程之前取签名源文,这时还未写入数据库
                                colSomeTime.Add Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_校对(" & lng医嘱ID & "," & _
                                IIF(.Cell(flexcpData, i, COL_选择) = 1, 3, 2) & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                "'" & IIF(.Cell(flexcpData, i, COL_选择) = 2, .TextMatrix(i, COL_操作说明), "") & "'," & _
                                "NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                            
                            '跟主界面显示相关的医嘱校对
                            If .TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" _
                                Or .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 Then
                            End If
                            If Not mbln发送调用 Then
                                mblnRefresh = True  '主界面调用时刷新主界面的清单
                            End If
                            
                            '当时是由于2011-1-13大医一院新增发送特殊医嘱时自动弹出校对这样改的，改得不对，调整为在发送窗体中设置mblnRefresh的值
                            '住院患者病情变更，校对后生效，要触发消息，正常校对通过
                            If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 And Val(.Cell(flexcpData, i, COL_选择)) = 1 Then
                                rsMsgRow.AddNew
                                rsMsgRow!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                rsMsgRow!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                rsMsgRow!行号 = i
                                rsMsgRow!操作类型 = 3
                                rsMsgRow!当前病情 = Get病人当前病情(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)))
                                rsMsgRow.Update
                            End If
                            
                            If Val(.Cell(flexcpData, i, COL_选择)) = 2 Then
                                '校对疑问
                                rsMsgRow.Filter = "病人id=" & Val(.TextMatrix(i, COL_病人ID)) & " And 主页id=" & Val(.TextMatrix(i, COL_主页ID)) & " And 操作类型=4"
                                If rsMsgRow.EOF Then
                                    rsMsgRow.AddNew
                                    rsMsgRow!病人ID = Val(.TextMatrix(i, COL_病人ID))
                                    rsMsgRow!主页ID = Val(.TextMatrix(i, COL_主页ID))
                                    rsMsgRow!行号 = i
                                    rsMsgRow!操作类型 = 4
                                    rsMsgRow.Update
                                End If
                            End If
                            If .TextMatrix(i, COL_诊疗类别) = "K" And gbln血库系统 Then
                                Call rs输血.AddNew(Array("医嘱ID", "类型"), Array(lng医嘱ID, 3)): bln输血 = True
                            End If
                        ElseIf mint类型 = 5 Then  '暂停医嘱
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_暂停(" & lng医嘱ID & ",To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.姓名 & "')"
                        ElseIf mint类型 = 6 Then  '启用医嘱
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_启用(" & lng医嘱ID & ",To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.姓名 & "')"
                        End If
                    End If
                Else
                    '界面上未勾选的医嘱行
                    If mint类型 = 3 Or mint类型 = 2 Then '医嘱校对和确认停止
                        If InStr(";" & mstrPatiKeepMsg & ";", ";" & Val(.TextMatrix(i, COL_病人ID)) & "," & Val(.TextMatrix(i, COL_主页ID)) & ";") = 0 Then
                            mstrPatiKeepMsg = mstrPatiKeepMsg & ";" & Val(.TextMatrix(i, COL_病人ID)) & "," & Val(.TextMatrix(i, COL_主页ID))
                        End If
                    End If
                End If
                lng相关ID = Val(.TextMatrix(i, COL_相关ID))
            Next
            mstrPatiKeepMsg = Mid(mstrPatiKeepMsg, 2)
            If mint类型 = 0 Then strRevokeIDs = Mid(strRevokeIDs, 2)
        End If
        
        If str配液禁止 <> "" Then
            If MsgBox("以下医嘱停止时间之后存在已经配药或发送的记录，是否继续停止医嘱?" & str配液禁止, vbQuestion + vbYesNo + vbDefaultButton2, "医嘱停止") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        ElseIf str配液打包 <> "" Then
            If MsgBox("以下医嘱停止时间之后存在已经配药或发送，但已经打包的记录，是否继续停止医嘱？" & str配液打包, vbQuestion + vbYesNo + vbDefaultButton2, "医嘱停止") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        ElseIf str配液提示 <> "" Then
            If MsgBox("以下医嘱停止时间之后存在已经摆药、但未配药的记录，是否继续停止医嘱？" & str配液提示, vbQuestion + vbYesNo + vbDefaultButton2, "医嘱停止") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        
        If str作废输血提示 <> "" Then
            MsgBox "本次作废的输血医嘱已经完成配血，不能直接作废医嘱，若要作废请与输血科联系。" & str作废输血提示, vbInformation, "医嘱作废"
            Screen.MousePointer = 0
            Exit Function
        End If
        
        '医嘱打印检查
        If strRevokeIDs <> "" Then
            strPrintDel = Get病人打印记录DelSQL(2, mlng病人ID, mlng主页ID, , , , strRevokeIDs, fraBaby.Visible, strMsg)
            If strMsg <> "" Then
                MsgBox "您作废的医嘱中包含已经打印的医嘱，请重打。", vbInformation, gstrSysName
                strPrintDel = ""
            End If
        End If
        
        If mblnHaveAudit Or (mint类型 <> 1 And mint类型 <> 7) Then
            '医嘱计价部分
            lng相关ID = 0
            If mint类型 = 2 Or mint类型 = 3 Or mint类型 = 4 Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) = 1 Then
                        '一并给药的只需处理一次
                        blnExe = False
                        If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) <> lng相关ID Then blnExe = True
                        Else
                            blnExe = True
                        End If
                        
                        If blnExe Then
                            '删除对应的计价
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_相关ID)) & ")"
                            Else
                                arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                            End If
                            
                            '生成新的计价
                            '本来用一次性循环快些,但为了判断是否要保存及输入合法性,必须用Filter
                            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) <> 0 Then
                                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                    " Or 医嘱ID=" & Val(vsAdvice.TextMatrix(i, COL_相关ID))
                            Else
                                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                    " Or 相关ID=" & vsAdvice.TextMatrix(i, COL_ID)
                            End If
                            For j = 1 To mrsPrice.RecordCount
                                '之中存在收费细目ID为空的无用记录(最初用于确定可选的计价医嘱)
                                '药品、卫材医嘱的计价固定对应不保存；非跟踪在用的时价卫材的变价需要输入，因此要保存到计价表中
                                If Not IsNull(mrsPrice!收费细目ID) And (InStr(",4,5,6,7,", mrsPrice!诊疗类别) = 0 _
                                    Or mrsPrice!诊疗类别 = "4" And NVL(mrsPrice!在用, 0) = 0 And NVL(mrsPrice!变价, 0) = 1) Then
                                    If NVL(mrsPrice!数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                        '普通项目的变价单价要求输入，包括非跟踪在用的时价卫材医嘱
                                        If NVL(mrsPrice!单价, 0) = 0 And NVL(mrsPrice!变价, 0) = 1 _
                                            And Not (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                            Call SeekPriceRow(i, mrsPrice!医嘱ID, mrsPrice!费用性质, mrsPrice!收费细目ID, COLP_单价)
                                            Screen.MousePointer = 0
                                            MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, gstrSysName
                                            vsPrice.SetFocus: Exit Function
                                        End If
                                        
                                        '计价执行科室:只保存非药品及卫材医嘱的，药品和卫材计价的执行科室
                                        If InStr(",4,5,6,7,", mrsPrice!诊疗类别) = 0 _
                                            And (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                            lng执行科室ID = NVL(mrsPrice!执行科室ID, 0)
                                            
                                            '卫材必须设置执行科室
                                            If lng执行科室ID = 0 And mrsPrice!收费类别 = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!医嘱ID, mrsPrice!费用性质, mrsPrice!收费细目ID, COLP_执行科室)
                                                Screen.MousePointer = 0
                                                MsgBox "卫材""" & vsPrice.TextMatrix(vsPrice.Row, COLP_收费项目) & """没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                                    "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                                vsPrice.SetFocus: Exit Function
                                            End If
                                        Else
                                            lng执行科室ID = 0
                                        End If
                                        
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Insert(" & mrsPrice!医嘱ID & "," & _
                                            mrsPrice!收费细目ID & "," & mrsPrice!数量 & "," & NVL(mrsPrice!单价, 0) & "," & _
                                            NVL(mrsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & "," & NVL(mrsPrice!费用性质, 0) & "," & NVL(mrsPrice!收费方式, 0) & ")"
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    lng相关ID = Val(.TextMatrix(i, COL_相关ID))
                Next
            End If
        
            '作废或停止时的电子签名
            If (mint类型 = 0 Or (mint类型 = 1 Or mint类型 = 7) Or mint类型 = 3 Or mint类型 = 2) And str医嘱ID <> "" Then
                strOper = Decode(mint类型, 0, "作废", 1, "停止", 3, "校对", 2, "确认停止")
            
                If (mint类型 = 3 Or mint类型 = 2) And rsAdviceTmp.State = adStateOpen Then
                    j = rsAdviceTmp.RecordCount
                    rsAdviceTmp.MoveFirst
                Else
                    j = 1
                End If
                For i = 1 To j
                    If (mint类型 = 3 Or mint类型 = 2) And rsAdviceTmp.State = adStateOpen Then
                        If rsAdviceTmp.EOF Then Exit For
                        strAdvicesTmp = rsAdviceTmp!医嘱ids & ""
                        lngPatiID = Val(rsAdviceTmp!病人ID & "")
                        lngPageId = Val(rsAdviceTmp!主页ID & "")
                    Else
                        strAdvicesTmp = str医嘱ID
                        lngPatiID = mlng病人ID
                        lngPageId = mlng主页ID
                    End If
                    
                    '护士不能作废、停止医生已签名的医嘱
                    If mbln护士站 And mint类型 <> 3 And mint类型 <> 2 Then
                        MsgBox "你要" & strOper & "的医嘱中包含医生已签名的医嘱，只能由医生来" & strOper & "并签名。", vbInformation, gstrSysName
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    '医生停止,作废时必须要签名
                    If gobjESign Is Nothing Then
                        If gintCA = 0 Then
                            MsgBox strOper & "已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能" & strOper & "。", vbInformation, gstrSysName
                        Else
                            MsgBox strOper & "已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能" & strOper & "。", vbInformation, gstrSysName
                        End If
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    '获取签名医嘱源文
                    strAdvicesTmp = Mid(strAdvicesTmp, 2) '组ID,返回为明细ID
                    intRule = ReadAdviceSignSource(Decode(mint类型, 0, 4, 1, 8, 3, 3), lngPatiID, lngPageId, strAdvicesTmp, 0, False, strSource, , colSomeTime)
                    If intRule = 0 Then Screen.MousePointer = 0: Exit Function
                    If strSource = "" Then
                        Screen.MousePointer = 0
                        MsgBox "不能读取需要" & strOper & "的已签名医嘱源文内容。", vbInformation, gstrSysName
                        Exit Function
                    End If
                     
                    strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
                    If strSign <> "" Then
                        If strTimeStamp <> "" Then
                            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            strTimeStamp = "NULL"
                        End If
                        lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_医嘱签名记录_Insert(" & lng签名id & "," & Decode(mint类型, 0, 4, 1, 8, 3, 3, 2, 9) & "," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strAdvicesTmp & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
                    Else
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    If (mint类型 = 3 Or mint类型 = 2) And rsAdviceTmp.State = adStateOpen Then
                        rsAdviceTmp.MoveNext
                    End If
                Next
                If rsAdviceTmp.State = adStateOpen Then rsAdviceTmp.Close
            End If
        End If
    End With
    varTmp = Split(strPrintDel, "|")
    
    If mint类型 = 0 Then
        Call CreatePlugInOK(p住院医嘱下达)
        If Not gobjPlugIn Is Nothing Then '调用作废前外挂接口
            On Error Resume Next
            arrRevokeID = Split(strRevokeIDs, ",")
            For i = 0 To UBound(arrRevokeID)
                If Val(arrRevokeID(i)) <> 0 Then
                    strMsg = ""
                    blnExe = gobjPlugIn.AdviceRevokedBefore(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(arrRevokeID(i)), -1, strMsg)
                    Call zlPlugInErrH(err, "AdviceRevokedBefore")
                    If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
                        If Not blnExe Then
                            MsgBox strMsg, vbInformation, gstrSysName
                            Screen.MousePointer = 0
                            Exit Function
                        End If
                    End If
                End If
            Next
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
        End If
    End If

    '执行SQL
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(varTmp)
        zlDatabase.ExecuteProcedure CStr(varTmp(i)), Me.Caption
    Next
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    If bln输血 And gbln血库系统 Then
        If InitObjBlood(True) = True Then
            rs输血.MoveFirst
            For i = 1 To rs输血.RecordCount
                If gobjPublicBlood.AdviceOperation(p住院医嘱下达, Val(rs输血!医嘱ID & ""), Val(rs输血!类型 & ""), False, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Screen.MousePointer = 0
                    MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
                rs输血.MoveNext
            Next
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    ExecuteOperate = True
    With vsAdvice
        If Not rsMsgRow.EOF Then
            rsMsgRow.Filter = "操作类型=1" '停止
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!行号)
                    '如果是实习生在停止医嘱则产生ZLHIS_CIS_027－新停审核提醒
                    If Not mblnHaveAudit And mlng停嘱审核 = 1 Then
                        Call ZLHIS_CIS_027(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_姓名), .TextMatrix(j, COL_住院号), , _
                            IIF(mlng病人性质 = 1, 1, 2), Val(.TextMatrix(j, COL_主页ID)), mlng病区ID, "", Val(.TextMatrix(j, COL_病人科室ID)), , , .TextMatrix(j, COL_床号), _
                            Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_诊疗类别), .TextMatrix(j, COL_操作类型), UserInfo.姓名, .TextMatrix(j, COL_停嘱时间), int紧急)
                    Else
                        Call ZLHIS_CIS_002(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_姓名), .TextMatrix(j, COL_住院号), , _
                            IIF(mlng病人性质 = 1, 1, 2), Val(.TextMatrix(j, COL_主页ID)), mlng病区ID, "", Val(.TextMatrix(j, COL_病人科室ID)), , , .TextMatrix(j, COL_床号), _
                            Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_诊疗类别), .TextMatrix(j, COL_操作类型), UserInfo.姓名, .TextMatrix(j, COL_停嘱时间), int紧急)
                    End If
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "操作类型=7" '停止审核
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!行号)
                    Call ZLHIS_CIS_002(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_姓名), .TextMatrix(j, COL_住院号), , _
                        IIF(mlng病人性质 = 1, 1, 2), Val(.TextMatrix(j, COL_主页ID)), mlng病区ID, "", Val(.TextMatrix(j, COL_病人科室ID)), , , .TextMatrix(j, COL_床号), _
                        Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_诊疗类别), .TextMatrix(j, COL_操作类型), UserInfo.姓名, .TextMatrix(j, COL_停嘱时间), int紧急)
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "操作类型=2" '作废
            If Not rsMsgRow.EOF Then
                strTimeStamp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!行号)
                    Call ZLHIS_CIS_003(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_姓名), .TextMatrix(j, COL_住院号), , _
                        IIF(mlng病人性质 = 1, 1, 2), Val(.TextMatrix(j, COL_主页ID)), mlng病区ID, "", Val(.TextMatrix(j, COL_病人科室ID)), "", , .TextMatrix(j, COL_床号), _
                        Val(.TextMatrix(j, COL_ID)), IIF(.TextMatrix(j, COL_期效) = "长嘱", 0, 1), .TextMatrix(j, COL_诊疗类别), Val(.TextMatrix(j, COL_操作类型)), Val(.TextMatrix(j, COL_执行分类)), _
                        Val(.TextMatrix(j, COL_执行科室ID)), UserInfo.姓名, strTimeStamp)
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "操作类型=3" '校对
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!行号)
                    lngTmp = 0: strTmp = ""
                    Call GetPatChange(Val(.TextMatrix(j, COL_ID)), 13, lngTmp, strTmp)
                    Set rsTmp = zlDatabase.OpenSQLRecord("select 名称 from 部门表 where id=[1]", Me.Caption, Val(.TextMatrix(j, COL_病人科室ID)))
                    Call ZLHIS_PATIENT_005(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_主页ID), .TextMatrix(j, COL_姓名), "", .TextMatrix(j, COL_住院号), _
                        0, , Val(.TextMatrix(j, COL_病人科室ID)), IIF(rsTmp.EOF, "", rsTmp!名称 & ""), rsMsgRow!当前病情 & "", lngTmp, .TextMatrix(j, COL_开嘱时间), strTmp, .TextMatrix(i, COL_开嘱医生), Val(.TextMatrix(j, COL_ID)))
                    rsMsgRow.MoveNext
                Next
            End If
            
            rsMsgRow.Filter = "操作类型=4" '校对疑问消息
            If Not rsMsgRow.EOF Then
                strTimeStamp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!行号)
                    Call ZLHIS_CIS_035(mclsMipModule, Val(.TextMatrix(j, COL_病人ID)), .TextMatrix(j, COL_姓名), .TextMatrix(j, COL_住院号), , _
                        IIF(mlng病人性质 = 1, 1, 2), Val(.TextMatrix(j, COL_主页ID)), mlng病区ID, "", Val(.TextMatrix(j, COL_病人科室ID)), "", , .TextMatrix(j, COL_床号), _
                        Val(.TextMatrix(j, COL_ID)), IIF(.TextMatrix(j, COL_期效) = "长嘱", 0, 1), .TextMatrix(j, COL_诊疗类别), Val(.TextMatrix(j, COL_操作类型)), Val(.TextMatrix(j, COL_执行分类)), _
                        Val(.TextMatrix(j, COL_执行科室ID)), UserInfo.姓名)
                    rsMsgRow.MoveNext
                Next
            End If
            
        End If
    End With
    Call ReadMsg
    Screen.MousePointer = 0
    If mint类型 = 0 Then
        '前一次的医嘱是否启用如果启用了，要检查是否被打印过
        If lng护理等级医嘱id <> 0 Then
            strSQL = "Select 婴儿,期效,页号 From 病人医嘱打印 Where 打印标记 In (1, 2) And 医嘱id = [1] And Not Exists" & _
                " (Select 1 From 病人医嘱记录 Where ID = [1] And 医嘱状态 In (8, 9)) And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng护理等级医嘱id)
            If Not rsTmp.EOF Then
                If MsgBox("本次作废了护理等级医嘱会自动启用前一次自动停止的护理等级医嘱，启用的护理等级医嘱已经进行了医嘱单停止时间或确认停止时间的套打，" & _
                    "如果要清除打印，将从第" & rsTmp!页号 & "开始清除，是否要清除后重打？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    strSQL = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & Val(rsTmp!婴儿 & "") & "," & Val(rsTmp!期效 & "") & "," & Val(rsTmp!页号 & "") & ")"
                    On Error GoTo errH
                    gcnOracle.BeginTrans: blnTrans = True
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    gcnOracle.CommitTrans: blnTrans = False
                    On Error GoTo 0
                End If
            End If
        End If
        
        Call CreatePlugInOK(p住院医嘱下达)
        '调用作废后外挂接口
        On Error Resume Next
        If Not gobjPlugIn Is Nothing Then
            arrRevokeID = Split(strRevokeIDs, ",")
            For i = 0 To UBound(arrRevokeID)
                If Val(arrRevokeID(i)) <> 0 Then
                    Call gobjPlugIn.AdviceRevoked(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(arrRevokeID(i)))
                    Call zlPlugInErrH(err, "AdviceRevoked")
                End If
            Next
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitPriceRecordset()
'说明：编辑时,当计价医嘱及收费项目都输入后,才加入记录集
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "诊疗类别", adVarChar, 1
    mrsPrice.Fields.Append "诊疗项目ID", adBigInt
    
    mrsPrice.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "执行标记", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "费用性质", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "收费方式", adInteger, , adFldIsNullable
    
    mrsPrice.Fields.Append "收费类别", adVarChar, 1, adFldIsNullable
    mrsPrice.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "单价", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "在用", adInteger '卫材是否跟踪在用
    mrsPrice.Fields.Append "变价", adInteger
    mrsPrice.Fields.Append "从项", adInteger
    mrsPrice.Fields.Append "固定", adInteger '现有的收费关系中是否固定对照
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub ShowDefaultRow()
'功能：对于可以计价的医嘱,缺省增加一行并设置缺省计价医嘱
'说明：ComboList="#医嘱ID1;计价医嘱1|#医嘱ID2;计价医嘱2|..."
'      仅在第一次显示计价表和回车新增行时调用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrCombo As Variant, lngRow As Long
    Dim lng医嘱ID As Long, int费用性质 As String, str计价医嘱 As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If .ColData(COLP_计价医嘱) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_计价医嘱), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_医嘱ID)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_收费细目ID)) <> 0 Then
                '第一次显示时缺省增加一行
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '不是第一次显示时缺省计价医嘱与上一行相同
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_固定)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_医嘱ID)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                int费用性质 = 0
                lng医嘱ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str计价医嘱 = Replace(arrCombo(i), "#" & lng医嘱ID & ";", "")
                
                If lng医嘱ID < 0 Then
                    int费用性质 = Val(Left(Abs(lng医嘱ID), 1))
                    lng医嘱ID = Val(Mid(Abs(lng医嘱ID), 2))
                End If
                If blnHave Then
                    If lng医嘱ID = Val(.TextMatrix(lngRow - 1, COLP_医嘱ID)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            '模拟选中这个计价医嘱
            strSQL = "Select 相关ID,诊疗类别,诊疗项目ID From 病人医嘱记录 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COLP_医嘱ID) = lng医嘱ID
                .TextMatrix(lngRow, COLP_费用性质) = int费用性质
                .TextMatrix(lngRow, COLP_计价医嘱) = str计价医嘱
                .TextMatrix(lngRow, COLP_相关ID) = NVL(rsTmp!相关ID)
                .TextMatrix(lngRow, COLP_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(lngRow, COLP_诊疗类别) = rsTmp!诊疗类别
                .Cell(flexcpData, lngRow, COLP_计价医嘱) = .TextMatrix(lngRow, COLP_计价医嘱)
                
                '只有一个计价医嘱时不必停留
                If UBound(arrCombo) = 0 Then
                    .Col = COLP_收费项目
                Else
                    .Col = COLP_计价医嘱
                End If
            End If
        End If
        Call .ShowCell(.Row, .Col)
        If blnFirst Then .TopRow = .Row '第一次显示时,ShowCell居然不起作用
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset, strSQL As String, i As Long
    Dim lng医嘱ID As Long, int费用性质 As Integer
    Dim lng原嘱ID As Long, int原费用性质 As Integer
    Dim lng收费细目ID As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_计价医嘱 Then
            '如果绑定了ComboData,TextMatrix取值就为ComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lng医嘱ID = .ComboData
                If lng医嘱ID < 0 Then
                    int费用性质 = Val(Left(Abs(lng医嘱ID), 1))
                    lng医嘱ID = Val(Mid(Abs(lng医嘱ID), 2))
                End If
                lng原嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                int原费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                
                '检查该计价医嘱是否已有相同收费细目
                If lng收费细目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """已经设置了收费项目""" & .TextMatrix(Row, COLP_收费项目) & """。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                '原来的医嘱如果有从项至少要保留一个(主项是固定不可动的)
                If lng原嘱ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int原费用性质 & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                '表格内容：mrsPrice中可能已删除,所以要从数据库读
                strSQL = "Select 相关ID,诊疗类别,诊疗项目ID,标本部位,检查方法,执行标记 From 病人医嘱记录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If rsTmp.EOF Then
                    MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """可能已经被其它人删除,请退出重新进入。", vbInformation, gstrSysName
                    Exit Sub
                End If
                .TextMatrix(Row, COLP_医嘱ID) = lng医嘱ID
                .TextMatrix(Row, COLP_费用性质) = int费用性质
                .TextMatrix(Row, COLP_相关ID) = NVL(rsTmp!相关ID)
                .TextMatrix(Row, COLP_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(Row, COLP_诊疗类别) = rsTmp!诊疗类别
                
                '记录集内容
                If lng收费细目ID <> 0 Then
                    '新选择的医嘱是否有从项决定修改后的项目是否从项
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 从项=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_从项) = IIF(blnHaveSub, "√", "")
                    
                    If lng原嘱ID = 0 Then
                        mrsPrice.AddNew '加入
                    Else '更新
                        mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                    End If
                    mrsPrice!医嘱ID = lng医嘱ID
                    mrsPrice!相关ID = rsTmp!相关ID
                    mrsPrice!诊疗项目ID = rsTmp!诊疗项目ID
                    mrsPrice!诊疗类别 = rsTmp!诊疗类别
                    
                    mrsPrice!标本部位 = rsTmp!标本部位
                    mrsPrice!检查方法 = rsTmp!检查方法
                    mrsPrice!执行标记 = NVL(rsTmp!执行标记, 0)
                    mrsPrice!费用性质 = int费用性质
                    mrsPrice!收费方式 = 0
                    
                    If lng原嘱ID = 0 Then
                        mrsPrice!收费细目ID = lng收费细目ID
                        mrsPrice!数量 = Val(.TextMatrix(Row, COLP_数量))
                        mrsPrice!单价 = Val(.TextMatrix(Row, COLP_单价))
                        mrsPrice!在用 = Val(.TextMatrix(Row, COLP_跟踪在用))
                        mrsPrice!变价 = Val(.Cell(flexcpData, Row, 0))
                        mrsPrice!固定 = 0
                    End If
                    mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            End If
        ElseIf Col = COLP_收费项目 Or Col = COLP_执行科室 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        ElseIf Col = COLP_数量 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        ElseIf Col = COLP_单价 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    On Error GoTo errH
    With vsAdvice
        If Col = COL_终止原因 Then
            strSQL = "select a.编码 as id, a.编码,a.名称,a.简码 from 停嘱原因 a order by a.编码"
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "停嘱原因", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COL_终止原因) = rsTmp!名称 & ""
                .Cell(flexcpData, Row, COL_终止原因) = rsTmp!名称 & ""
                Call SetSame原因(Row)
            Else
                If Not blnCancel Then
                    MsgBox "没有可用的停嘱原因，请先到字典管理中设置！", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str项目IDs As String, blnCancel As Boolean
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim int费用性质 As Integer, vPoint As PointAPI
    Dim strStock As String
    Dim strSQL2 As String
    
    With vsPrice
        If Col = COLP_收费项目 Then
            '不能选择已有的项目
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_医嘱ID)) = Val(.TextMatrix(Row, COLP_医嘱ID)) _
                    And Val(.TextMatrix(Row, COLP_医嘱ID)) <> 0 And i <> Row Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                End If
            Next
            str项目IDs = Mid(str项目IDs, 2)
            
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            
            
            '药品卫材库存
            Call GetDefaultDeptPar(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_出院科室ID)))
            If mlng西药房 <> 0 Or mlng成药房 <> 0 Or mlng中药房 <> 0 Or mlng发料部门 <> 0 Then
                strStock = _
                    "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                    " From 药品库存 A,收费项目目录 B" & _
                    " Where A.性质 = 1 And (Nvl(A.批次,0)=0 Or A.效期 Is Null Or A.效期>Trunc(Sysdate))" & _
                        " And A.库房ID=Decode(B.类别,'5',[3],'6',[4],'7',[5],'4',[6],Null)" & _
                        " And A.药品ID=B.ID And B.类别 IN('4','5','6','7')" & _
                        " And (b.执行科室 <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = b.Id And (w.病人来源=2 or (w.病人来源 is Null And w.开单科室id = [7]))))" & _
                    " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
            Else
                strStock = "Select Null as 药品ID,Null as 库存 From Dual"
            End If
            
            strSQL = _
                " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
                " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 价格,NULL as 库存,NULL as 费用类型,NULL as 医保大类," & _
                " NULL as 说明,-NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,-ID as ID,Nvl(-上级ID,To_Number('999999999'||类型)) as 上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 价格,NULL as 库存,NULL as 费用类型,NULL as 医保大类," & _
                " NULL as 说明,-NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,ID,上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 价格,NULL as 库存,NULL as 费用类型,NULL as 医保大类," & _
                " NULL as 说明,-NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL2 = _
                " Select A.末级,A.ID,A.上级ID,A.编码,A.名称,A.单位,A.规格,A.产地,A.类别," & _
                " Decode(Nvl(A.是否变价,0),1,Decode(Instr('567',A.类别ID),0,Sum(Nvl(A.原价,0))||'-'||Sum(Nvl(A.现价,0))||'/'||单位,'时价')," & _
                "   Decode(Instr('567',A.类别ID),0,Sum(A.现价)||'/'||A.单位,LTrim(To_Char(Sum(A.现价)*A.住院包装,'999990.0000'))||'/'||A.住院单位)) as 价格," & _
                " Decode(Instr('4567',A.类别ID),0,NULL,1," & _
                "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存,'999990.0000'))||A.单位)," & _
                "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存/Nvl(A.住院包装,1),'999990.0000'))||A.住院单位)) as 库存," & _
                " A.费用类型,A.医保大类,A.说明,Sum(A.原价) as 原价ID,Sum(A.现价) as 现价ID,Sum(A.缺省价格) as 缺省价格ID,A.是否变价 as 是否变价ID,A.类别ID,A.跟踪在用ID" & _
                " From (" & _
                " Select Distinct 1 as 末级,A.ID,Decode(Instr('567',A.类别),0,A.分类ID,-E.分类ID) as 上级ID,A.编码,A.名称," & _
                " A.计算单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价," & _
                " -NULL as 跟踪在用ID,D.住院单位,D.住院包装" & _
                " From 收费项目目录 A,收费价目 B,收费项目类别 C,药品规格 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.ID=B.收费细目ID [选择替换的过条件1]  And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "8", "9", "10") & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.类别 Not IN('4','J','1') And A.类别=C.编码 And A.ID=D.药品ID(+) And D.药名ID=E.ID(+)" & _
                " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[7]) = [7]))))" & _
                " And (a.类别 Not in ('5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[7])=[7]))"
            If DeptExist("发料部门", 2) Then
                strSQL2 = strSQL2 & " Union ALL" & _
                    " Select Distinct 1 as 末级,A.ID,-E.分类ID as 上级ID,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " B.原价,B.现价,B.缺省价格,A.是否变价,D.跟踪在用 as 跟踪在用ID,NULL as 住院单位,NULL as 住院包装" & _
                    " From 收费项目目录 A,收费价目 B,收费项目类别 C,材料特性 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                    " Where A.ID=B.收费细目ID [选择替换的过条件2]  And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "8", "9", "10") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And A.ID Not IN(" & str项目IDs & ")", "") & _
                    " And A.类别='4' And A.类别=C.编码 And A.ID=D.材料ID And D.诊疗ID=E.ID And D.核算材料=0" & _
                    " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                    " And Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[7])=[7])"
            End If
            strSQL2 = strSQL2 & " ) A,(" & strStock & ") S Where A.ID=S.药品ID(+)" & _
            " Group by A.末级,A.ID,A.上级ID,A.编码,A.名称,A.单位,A.规格,A.产地,A.类别,A.费用类型,A.医保大类,A.说明,A.是否变价,A.类别ID,A.跟踪在用ID,A.住院单位,A.住院包装,S.库存"
            '[选择替换的过条件1],[选择替换的过条件2],这两个串在选器中处理的
            '要确保 "占位参数" 在最后一位，该参数在选择器中拼接，要解决4000长度的问题
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "收费项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, _
                "," & str项目IDs & ",", mint险类, mlng西药房, mlng成药房, mlng中药房, mlng发料部门, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "占位参数")
            If Not rsTmp Is Nothing Then
                '医保对码检查
                If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID))) Then
                    .SetFocus: Exit Sub
                End If
            
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有可用的收费项目，请先到收费项目管理中设置！", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_执行科室 Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_收费类别) = "4" Then
                '跟踪在用的卫材
                strSQL = _
                    " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                    " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And A.收费细目ID=[1]" & _
                    " Order by B.服务对象,C.编码"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                '药品
                '药品从系统指定的储备药房中找
                If Not Check上班安排(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                    Decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!名称
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!执行科室ID = rsTmp!ID
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查输入(选择)计价项目是否医保对码
'返回：如果未对码，并且提示选择不继续，则返回真。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int险类 As Integer
    
    If gint医保对码 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select 险类 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then int险类 = NVL(rsTmp!险类, 0)
    If int险类 <> 0 Then
        If Not ItemExistInsure(lng病人ID, rsInput!ID, int险类) Then
            If gint医保对码 = 1 Then
                If MsgBox("项目""" & rsInput!名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gint医保对码 = 2 Then
                MsgBox "项目""" & rsInput!名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lng医嘱ID As Long, int费用性质 As Integer, ByVal lng原项目ID As Long)
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng行号 As Long, dbl单价 As Double
    Dim blnHaveSub As Boolean, dbl总量 As Double
    Dim rsTmp As ADODB.Recordset
    
    With vsPrice
        '表格内容
        .TextMatrix(lngRow, COLP_收费类别) = rsInput!类别ID
        .TextMatrix(lngRow, COLP_收费细目ID) = rsInput!ID
        .TextMatrix(lngRow, COLP_类别) = rsInput!类别
        .TextMatrix(lngRow, COLP_收费项目) = rsInput!名称
        If Not IsNull(rsInput!产地) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & "(" & rsInput!产地 & ")"
        End If
        If Not IsNull(rsInput!规格) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & " " & rsInput!规格
        End If
        
        '如果是加入药品计价(非药嘱),按零售单位处理
        .TextMatrix(lngRow, COLP_数量) = 1 '缺省计价数量为1
        .TextMatrix(lngRow, COLP_单位) = NVL(rsInput!单位)
                
        '单价计算处理:药嘱计价不可能在这里处理,非药嘱药品计价按售价处理
        .Cell(flexcpData, lngRow, 0) = 0
        .Cell(flexcpData, lngRow, 1) = 0
        .Cell(flexcpData, lngRow, 2) = 0
        
        '执行科室
        lng行号 = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If lng行号 = -1 Then
            Set rsTmp = Sys.RowValue("病人医嘱记录", lng医嘱ID)
            lng病人ID = rsTmp!病人ID
            lng主页ID = NVL(rsTmp!主页ID, 0)
            lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
            lng病人科室ID = NVL(rsTmp!病人科室id, 0)
            dbl总量 = NVL(rsTmp!总给予量, 0)
            If dbl总量 = 0 Then dbl总量 = 1
        Else
            lng病人ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人ID))
            lng主页ID = Val(vsAdvice.TextMatrix(lng行号, COL_主页ID))
            lng执行科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))
            lng病人科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID))
            dbl总量 = Val(vsAdvice.TextMatrix(lng行号, COL_总量))
            If dbl总量 = 0 Then dbl总量 = 1
        End If
            
        '非药嘱和跟踪在用的卫材专门求执行科室
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Or rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 Then
            lng执行科室ID = Get收费执行科室ID(lng病人ID, lng主页ID, rsInput!类别ID, rsInput!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
            '记录卫材是否跟踪在用
            If rsInput!类别ID = "4" Then
                .TextMatrix(lngRow, COLP_跟踪在用) = NVL(rsInput!跟踪在用ID, 0)
            End If
        End If
        If lng执行科室ID <> 0 Then
            mrsDept.Filter = "ID=" & lng执行科室ID
            If Not mrsDept.EOF Then
                .TextMatrix(lngRow, COLP_执行科室) = mrsDept!名称
            End If
        End If
        .TextMatrix(lngRow, COLP_执行科室ID) = lng执行科室ID
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng病人ID, lng主页ID, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        '单价
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                dbl单价 = NVL(rsInput!现价ID, 0)
            Else '未确定计价医嘱时,药品无法计算价格
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, dbl总量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级) '按缺省计价数量为1个零售单位计算
            End If
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
        ElseIf rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 And NVL(rsInput!是否变价ID, 0) = 1 Then
            '跟踪在用的时价卫材和药品一样计算
            dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, dbl总量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
        Else
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!现价ID, 0), gstrDecPrice)
            Else
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!缺省价格ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, 0) = 1
                .Cell(flexcpData, lngRow, 1) = NVL(rsInput!原价ID, 0)
                .Cell(flexcpData, lngRow, 2) = NVL(rsInput!现价ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_费用类型) = NVL(rsInput!费用类型)
        .TextMatrix(lngRow, COLP_固定) = "0"
        .TextMatrix(lngRow, COLP_收费方式) = "正常收取"
        
        '用于输入恢复
        .Cell(flexcpData, lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目)
        .Cell(flexcpData, lngRow, COLP_数量) = .TextMatrix(lngRow, COLP_数量)
        .Cell(flexcpData, lngRow, COLP_单价) = .TextMatrix(lngRow, COLP_单价)
        .Cell(flexcpData, lngRow, COLP_执行科室) = .TextMatrix(lngRow, COLP_执行科室)
        
        '记录集内容
        If lng医嘱ID <> 0 Then
            If lng原项目ID = 0 Then
                '当前医嘱是否有从项决定新增的项目是否从项
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 从项=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_从项) = IIF(blnHaveSub, "√", "")

                mrsPrice.AddNew '加入
            Else '更新
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
            End If
            If lng原项目ID = 0 Then
                If rsTmp Is Nothing Then
                    Set rsTmp = Sys.RowValue("病人医嘱记录", lng医嘱ID)
                End If
                mrsPrice!医嘱ID = lng医嘱ID
                mrsPrice!相关ID = IIF(Val(.TextMatrix(lngRow, COLP_相关ID)) = 0, Null, Val(.TextMatrix(lngRow, COLP_相关ID)))
                mrsPrice!诊疗类别 = .TextMatrix(lngRow, COLP_诊疗类别)
                mrsPrice!诊疗项目ID = Val(.TextMatrix(lngRow, COLP_诊疗项目ID))
                mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                
                mrsPrice!标本部位 = rsTmp!标本部位
                mrsPrice!检查方法 = rsTmp!检查方法
                mrsPrice!执行标记 = NVL(rsTmp!执行标记, 0)
                mrsPrice!费用性质 = int费用性质
            End If
            mrsPrice!收费方式 = 0
            mrsPrice!收费类别 = rsInput!类别ID
            mrsPrice!收费细目ID = rsInput!ID
            If lng执行科室ID <> 0 Then
                mrsPrice!执行科室ID = lng执行科室ID
            Else
                mrsPrice!执行科室ID = Null
            End If
            mrsPrice!在用 = NVL(rsInput!跟踪在用ID, 0)
            mrsPrice!变价 = NVL(rsInput!是否变价ID, 0)
            mrsPrice!数量 = 1
            mrsPrice!单价 = Val(.TextMatrix(lngRow, COLP_单价))
            mrsPrice!固定 = 0
            mrsPrice.Update
            Call SelectRow(vsAdvice.Row)
        End If
    End With
End Sub

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_计价医嘱 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_固定)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_医嘱ID)) <> 0 And Val(.TextMatrix(.Row, COLP_收费细目ID)) <> 0 Then
                    '医嘱如果有从项至少要保留一个(主项是固定不可动的)
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(.Row, COLP_医嘱ID)) & " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_计价医嘱) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("确定要删除当前计价行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(.Row, COLP_医嘱ID)) & " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 收费细目ID=" & Val(.TextMatrix(.Row, COLP_收费细目ID))
                    mrsPrice.Delete
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_计价医嘱
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_收费项目 Or .Col = COLP_执行科室) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str项目IDs As String, int费用性质 As Integer
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI, strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COLP_计价医嘱 Then
                '下拉时回车
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '不然EnterNextCell函数要退出
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_数量 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费数量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_单价 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费单价输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '检查变价输入范围
                strTmp = CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_收费项目 And .EditText <> "" Then
                '不能选择已有的项目
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COLP_医嘱ID)) = Val(.TextMatrix(Row, COLP_医嘱ID)) _
                        And Val(.TextMatrix(Row, COLP_医嘱ID)) <> 0 And i <> Row Then
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                    End If
                Next
                str项目IDs = Mid(str项目IDs, 2)
                
                '不同的输入匹配方式
                strInput = UCase(.EditText)
                strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
                ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strMatch = " And C.名称 Like [2] And C.码类=[3]"
                End If
                
                '药品卫材库存
                Call GetDefaultDeptPar(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_出院科室ID)))
                If mlng西药房 <> 0 Or mlng成药房 <> 0 Or mlng中药房 <> 0 Or mlng发料部门 <> 0 Then
                    strStock = _
                        "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                        " From 药品库存 A,收费项目目录 B" & _
                        " Where A.性质 = 1 And (Nvl(A.批次,0)=0 Or A.效期 Is Null Or A.效期>Trunc(Sysdate))" & _
                            " And A.库房ID=Decode(B.类别,'5',[6],'6',[7],'7',[8],'4',[9],Null)" & _
                            " And A.药品ID=B.ID And B.类别 IN('4','5','6','7')" & _
                        " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
                Else
                    strStock = "Select Null as 药品ID,Null as 库存 From Dual"
                End If
                If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                
                strSQL = ""
                If Not DeptExist("发料部门", 2) Then strSQL = " And A.类别<>'4'"
                strSQL = "Select * From (" & _
                    " Select A.末级,A.ID,A.类别,A.编码,A.名称,Decode(Instr('567',A.类别ID),0,A.单位,C.住院单位) as 单位,A.规格,A.产地," & _
                    " Decode(Nvl(A.是否变价,0),1,Decode(Instr('567',A.类别ID),0,Sum(Nvl(A.原价,0))||'-'||Sum(Nvl(A.现价,0))||'/'||A.单位,'时价')," & _
                    "   Decode(Instr('567',A.类别ID),0,Sum(A.现价)||'/'||A.单位,LTrim(To_Char(Sum(A.现价)*C.住院包装,'999990.0000'))||'/'||C.住院单位)) as 价格," & _
                    " Decode(Instr('4567',A.类别ID),0,NULL,1," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存,'999990.0000'))||A.单位)," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存/Nvl(C.住院包装,1),'999990.0000'))||C.住院单位)) as 库存,A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " Sum(A.原价) as 原价ID,Sum(A.现价) as 现价ID,Sum(A.缺省价格) as 缺省价格ID,A.是否变价 as 是否变价ID,A.类别ID,B.跟踪在用 as 跟踪在用ID,B.核算材料" & _
                    " From (" & _
                    " Select Distinct 1 as 末级,A.ID,a.执行科室,A.类别 as 类别ID,D.名称 as 类别,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,A.费用类型,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价" & _
                    " From 收费项目目录 A,收费价目 B,收费项目别名 C,收费项目类别 D" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "11", "12", "13") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.收费细目ID And A.类别=D.编码 And A.类别 Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,材料特性 B,药品规格 C,保险支付项目 M,保险支付大类 N,(" & strStock & ") S" & _
                    " Where A.ID=B.材料ID(+) And A.ID=C.药品ID(+) And A.ID=S.药品ID(+)" & _
                    " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[5]" & _
                    " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[10]) = [10]))))" & _
                    " And (a.类别id not in ('4','5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[10])=[10]))" & _
                    " Group by A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地,A.费用类型,C.住院单位,C.住院包装,S.库存,N.名称,A.说明,A.是否变价,A.类别ID,B.跟踪在用,B.核算材料" & _
                    " ) Where Nvl(核算材料,0) = 0 Order by 类别,编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "收费项目", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint简码 + 1, "," & str项目IDs & ",", mint险类, mlng西药房, mlng成药房, mlng中药房, mlng发料部门, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                If Not rsTmp Is Nothing Then
                    '医保对码检查
                    If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID))) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                        .SetFocus: Exit Sub
                    End If
                    
                    lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的收费项目！", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            ElseIf Col = COLP_执行科室 And .EditText <> "" Then '执行科室
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_收费类别) = "4" Then
                    '跟踪在用的卫材
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1] And (C.编码 Like [3] Or C.名称 Like [4] Or C.简码 Like [4])" & _
                        " Order by B.服务对象,C.编码"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                    '药品从系统指定的储备药房中找
                    If Not Check上班安排(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                        Decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!名称
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    
                    '更新记录集
                    lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                        mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                        mrsPrice!执行科室ID = rsTmp!ID
                        mrsPrice.Update
                        Call SelectRow(vsAdvice.Row)
                    End If
                    
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_数量 Or Col = COLP_单价 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_计价医嘱 Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_收费项目 Or NewCol = COLP_执行科室 Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
    
    If NewRow <> OldRow Then
        '显示药品及跟踪卫材的库存
        With vsPrice
            stbThis.Panels(2).Text = ""
            If Val(.TextMatrix(NewRow, COLP_收费细目ID)) <> 0 Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_收费类别)) > 0 _
                    Or .TextMatrix(NewRow, COLP_收费类别) = "4" And Val(.TextMatrix(NewRow, COLP_跟踪在用)) = 1 Then
                    '这里计价只输入和显示相对数量：药嘱药品按住院单位，非药嘱药品按售价单位
                    If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                        If GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID))) > 0 Then
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "有库存"
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "无库存"
                        End If
                    Else
                        If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_诊疗类别)) > 0 Then
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "可用库存:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID))), 5) & .TextMatrix(NewRow, COLP_单位)
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "可用库存:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID)), 0), 5) & .TextMatrix(NewRow, COLP_单位)
                        End If
                    End If
                End If
            End If
        End With
        
        '显示医保大类
        stbThis.Panels(3).Text = Get医保大类(NewRow)
    End If
End Sub

Private Function Get医保大类(ByVal lngRow As Long) As String
'功能：获取指定行的费用类型
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str大类 As String
    
    With vsPrice
        If Val(.TextMatrix(lngRow, COLP_收费细目ID)) <> 0 Then
            strSQL = "Select N.名称 From 保险支付项目 M,保险支付大类 N Where M.收费细目ID=[1] And M.大类ID=N.ID And M.险类=[2]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COLP_收费细目ID)), mint险类)
            If Not rsTmp.EOF Then str大类 = NVL(rsTmp!名称)
        End If
    End With
    Get医保大类 = IIF(str大类 <> "", "医保大类:" & str大类, "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Cancel = True: Exit Sub
    
    If Not CellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = COLP_数量 Or Col = COLP_单价 Or Col = COLP_执行科室 Then
        If vsPrice.TextMatrix(Row, COLP_收费项目) = "" Then
            Cancel = True '必须先确定收费项目
        End If
    End If
    
    If Col = COLP_数量 Or Col = COLP_单价 Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断价表中单元格是否可以编辑
    CellEditable = vsPrice.Editable
    With vsPrice
        If lngCol = COLP_执行科室 Then
            '非药品及卫材医嘱的，药品和卫材计价的执行科室可以修改
            If Not ((.TextMatrix(lngRow, COLP_收费类别) = "4" And Val(.TextMatrix(lngRow, COLP_跟踪在用)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_收费类别)) > 0) And InStr(",4,5,6,7,", .TextMatrix(lngRow, COLP_诊疗类别)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_收费项目) = "" Or .TextMatrix(lngRow, COLP_诊疗类别) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_固定)) <> 0 Then
            '固定对照行仅可以修改变价
            If Not (.Cell(flexcpData, lngRow, 0) = 1 And lngCol = COLP_单价) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_单价 Then
                If .Cell(flexcpData, lngRow, 0) <> 1 Then CellEditable = False
            ElseIf lngCol <> COLP_计价医嘱 And lngCol <> COLP_数量 And lngCol <> COLP_收费项目 Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：定位到价表中下一个可以输入的单元格
    Dim i As Long, j As Long
    
    With vsPrice
        '当前单元格如果未输入完整,则退出
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_单价 And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '从下一单元开始循环搜索
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_计价医嘱) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '当前表格内没有找到下一个可编辑单元,如果有需计价医嘱,则增加一新行
            If CStr(.ColData(COLP_计价医嘱)) <> "" Then
                '当前行未输入完整,则定位到不完整单元
                If .TextMatrix(lngRow, COLP_计价医嘱) = "" Then
                    .Col = COLP_计价医嘱
                ElseIf .TextMatrix(lngRow, COLP_数量) = "" Then
                    .Col = COLP_数量
                ElseIf .TextMatrix(lngRow, COLP_收费项目) = "" Then
                    .Col = COLP_收费项目
                ElseIf .Cell(flexcpData, lngRow, 0) = 1 And Val(.TextMatrix(lngRow, COLP_单价)) = 0 Then
                    .Col = COLP_单价
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_计价医嘱
                    
                    '缺省选择计价医嘱(如果可能)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '不可编辑时随意定一个
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Function LoadPrice(ByVal lngRow As Long, Optional blnChange As Boolean) As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费 关系进行更新
'返回：blnChange=是否根据当前的诊疗收费 关系对现有的计价内容进行了调整
    Dim rsMan As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim rsAdd As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim dblPrice As Double, strSubItem As String
    Dim lng执行科室ID As Long
    Dim lng材料ID As Long, blnLoad As Boolean
    Dim lng医嘱ID As Long, lng相关ID As Long
    
    On Error GoTo errH
    
    With vsAdvice
        '已经读取过了,不再重复读取
        If .TextMatrix(lngRow, COL_ID) = "" Then LoadPrice = True: Exit Function
        If .RowData(lngRow) = 1 Then LoadPrice = True: Exit Function
        
        
        lng医嘱ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(.TextMatrix(lngRow, COL_病人ID)), Val(.TextMatrix(lngRow, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        
                            
        '药品、卫材的计价(这里仅用于显示；数量为相对数量,药品固定为1；实价药品单价显示时计算)
        '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
        If .TextMatrix(lngRow, COL_诊疗类别) = "4" Then
            '卫材计价
            strSQL = _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,Null as 标本部位,Null as 检查方法,0 as 执行标记," & _
                " C.类别 as 收费类别,A.收费细目ID,1 as 数量,Decode(Nvl(C.是否变价,0),1,Nvl(X.单价,D.缺省价格),D.现价) as 单价," & _
                " 0 as 从项,A.执行科室ID,B.跟踪在用,C.是否变价,C.撤档时间,0 as 费用性质,0 as 收费方式" & _
                " From 病人医嘱记录 A,材料特性 B,收费项目目录 C,收费价目 D,病人医嘱计价 X" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=X.医嘱ID(+)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "3", "4", "5") & _
                " And A.收费细目ID=B.材料ID And A.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(2,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算1个住院包装的单价
            strSQL = _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,Null as 标本部位,Null as 检查方法,0 as 执行标记," & _
                " C.类别 as 收费类别,C.ID as 收费细目ID,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价," & _
                " 0 as 从项,A.执行科室ID,0 as 跟踪在用,C.是否变价,C.撤档时间,0 as 费用性质,0 as 收费方式" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "3", "4", "5") & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(2,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        ElseIf .TextMatrix(lngRow, COL_类型) = "1" Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,Null as 标本部位,Null as 检查方法,0 as 执行标记," & _
                " C.类别 as 收费类别,C.ID as 收费细目ID,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价," & _
                " 0 as 从项,A.执行科室ID,0 as 跟踪在用,C.是否变价,C.撤档时间,0 as 费用性质,0 as 收费方式" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN(2,3)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "3", "4", "5") & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价：除药品外的计价,包含相关医嘱计价
        blnLoad = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '给药途径:一并给药的只读取一次来共用
            If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                If .TextMatrix(lngRow - 1, COL_相关ID) = .TextMatrix(lngRow, COL_相关ID) Then
                    blnLoad = False
                End If
            End If
        End If
        If blnLoad Then
            '成药的给药途径；中药配方的煎法，用法；检查及部位；手术及附加手术,麻醉项目
            '不计价,手工计价；叮嘱,院外执行；的医嘱不读取
            '用Union方式可以利用索引
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记," & _
                "   B.类别 as 收费类别,A.收费细目ID,A.数量,A.单价,Nvl(A.从项,0) as 从项,A.执行科室ID," & _
                "   C.跟踪在用,B.是否变价,B.撤档时间,Nvl(A.费用性质,0) as 费用性质,Nvl(A.收费方式,0) as 收费方式" & _
                " From (" & _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,A.标本部位,Decode(A.诊疗类别,'E',Decode(Z.操作类型,'4',Null,A.检查方法),A.检查方法) as 检查方法,A.执行标记," & _
                "   B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,B.费用性质,B.收费方式" & _
                " From 病人医嘱记录 A,病人医嘱计价 B,诊疗项目目录 Z" & _
                " Where Z.id=a.诊疗项目ID And A.诊疗类别 Not IN('4','5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.ID=[1]" & _
                " Union ALL" & _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记," & _
                "   B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,B.费用性质,B.收费方式" & _
                " From 病人医嘱记录 A,病人医嘱计价 B" & _
                " Where A.诊疗类别 Not IN('4','5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.ID=[2]" & _
                " Union ALL" & _
                " Select A.ID,A.相关ID,A.序号,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记," & _
                "   B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,B.费用性质,B.收费方式" & _
                " From 病人医嘱记录 A,病人医嘱计价 B" & _
                " Where A.诊疗类别 Not IN('4','5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.相关ID=[1]" & _
                " ) A,收费项目目录 B,材料特性 C" & _
                " Where A.收费细目ID=B.ID(+) And A.收费细目ID=C.材料ID(+)" & _
                " Order by 序号,费用性质,从项"
        End If
        Set rsMan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        
        '诊疗收费 关系中收费数量及固有对照是否变化
        '不同的科室可能有不同的收费对照，如果查出来多行，则针对执行科室的优先
        strSQL = "Select * From (Select C.诊疗项目ID,C.收费项目ID,C.收费数量,C.固有对照,C.从属项目," & _
            " Nvl(C.检查部位,'None') as 检查部位,Nvl(C.检查方法,'None') as 检查方法," & _
            " Nvl(C.费用性质,0) as 费用性质,Nvl(C.收费方式,0) as 收费方式,C.适用科室id" & _
            " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
            " From 病人医嘱记录 A,病人医嘱计价 B,诊疗收费关系 C" & _
            " Where A.ID=B.医嘱ID And A.诊疗项目ID+0=C.诊疗项目ID And B.收费细目ID+0=C.收费项目ID" & _
            " And (C.适用科室ID is Null or C.适用科室ID = Nvl(A.执行科室id,[3]) And C.病人来源 = 2)" & _
            " And (A.相关ID is Null And A.执行标记 IN(1,2) And C.费用性质=1" & _
            "       Or A.标本部位=C.检查部位 And A.检查方法=C.检查方法 And Nvl(C.费用性质,0)=0" & _
            "       Or (A.检查方法 is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(C.费用性质,0)=0 And C.检查部位 is Null And C.检查方法 is Null)" & _
            " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0)<>5 And (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1])" & _
            " ) Where Nvl(适用科室id, 0) = Top"
        Set rsCur = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng相关ID, mlng病区ID)
        
        '加入药品及现有的计价
        For i = 1 To rsMan.RecordCount
            strSubItem = ""
            
            mrsPrice.AddNew '暂未输入计价关系的也要加入用于确定可计价医嘱(该记录无用)
            mrsPrice!医嘱ID = rsMan!ID
            mrsPrice!相关ID = rsMan!相关ID
            mrsPrice!诊疗类别 = rsMan!诊疗类别
            mrsPrice!诊疗项目ID = rsMan!诊疗项目ID
            mrsPrice!固定 = IIF(InStr(",4,5,6,7,", rsMan!诊疗类别) > 0, 1, 0)
            
            '检查项目的扩展
            mrsPrice!标本部位 = rsMan!标本部位
            mrsPrice!检查方法 = rsMan!检查方法
            mrsPrice!执行标记 = NVL(rsMan!执行标记, 0)
            mrsPrice!费用性质 = NVL(rsMan!费用性质, 0)
            mrsPrice!收费方式 = NVL(rsMan!收费方式, 0)
            
            '在设置医嘱计价时,对于原已设置,但现已撤档的项目,当作未设置(以便重新增加)
            If Not IsNull(rsMan!收费细目ID) _
                And Format(NVL(rsMan!撤档时间, "3000-01-01"), "yyyy-MM-dd") = "3000-01-01" Then
                mrsPrice!收费类别 = rsMan!收费类别
                mrsPrice!收费细目ID = rsMan!收费细目ID
                mrsPrice!执行科室ID = rsMan!执行科室ID
                mrsPrice!在用 = NVL(rsMan!跟踪在用, 0)
                mrsPrice!变价 = NVL(rsMan!是否变价, 0)
                mrsPrice!数量 = rsMan!数量
                
                '药品(仅用于显示)：如果为时价，显示时计算；否则就是取的最新价格
                '卫材：如果为跟踪在用时价，显示时计算；否则取定价或以前定的(如果有)
                '非药品：如果为变价,则取以前定的(如果有)；否则下面取最新价格
                mrsPrice!单价 = rsMan!单价
                mrsPrice!从项 = NVL(rsMan!从项, 0)
                        
                '诊疗收费 关系中收费数量及固有对照是否变化
                If InStr(",4,5,6,7,", rsMan!诊疗类别) = 0 Then '包含非药品、卫材医嘱的药品计价
                    If rsMan!诊疗类别 = "D" Then
                        rsCur.Filter = "诊疗项目ID=" & rsMan!诊疗项目ID & " And 收费项目ID=" & rsMan!收费细目ID & _
                            " And 检查部位='" & NVL(rsMan!标本部位, "None") & "' And 检查方法='" & NVL(rsMan!检查方法, "None") & "'" & _
                            " And 费用性质=" & NVL(rsMan!费用性质, 0)
                    Else
                        rsCur.Filter = "诊疗项目ID=" & rsMan!诊疗项目ID & " And 收费项目ID=" & rsMan!收费细目ID & " And 检查部位='None' And 检查方法='None' And 费用性质=" & NVL(rsMan!费用性质, 0)
                    End If
                    If Not rsCur.EOF Then
                        If NVL(rsCur!固有对照, 0) <> 0 And NVL(rsMan!数量, 0) <> NVL(rsCur!收费数量, 0) Then
                            mrsPrice!数量 = rsCur!收费数量 '变成了固有对照才取新设置的数量
                            blnChange = True
                        End If
                        mrsPrice!从项 = NVL(rsCur!从属项目, 0)
                        mrsPrice!固定 = NVL(rsCur!固有对照, 0)
                    End If
                    '价格取最新的(非变价)
                    dblPrice = CalcPrice(rsMan!收费细目ID, , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    If dblPrice <> 0 Then mrsPrice!单价 = Format(dblPrice, gstrDecPrice)
                End If
            End If
            mrsPrice.Update
            
            '存在从属项目的要计价医嘱
            If mrsPrice!从项 = 1 Then
                If InStr(strSubItem & ";", ";" & mrsPrice!医嘱ID & "," & mrsPrice!费用性质 & ";") = 0 Then
                    strSubItem = strSubItem & ";" & mrsPrice!医嘱ID & "," & mrsPrice!费用性质
                End If
            End If
            
            '诊疗收费 关系中新增了的对照(在未校对之前,病人医嘱计价没有内容,这时也是相对增加的)
            If InStr(",1,2,", NVL(rsMan!医嘱状态, 0)) > 0 And InStr(",4,5,6,7,", rsMan!诊疗类别) = 0 Then '包含非药品、卫材医嘱的药品计价
                lng医嘱ID = rsMan!ID
                blnLoad = False: rsMan.MoveNext
                If rsMan.EOF Then
                    blnLoad = True
                ElseIf rsMan!ID <> lng医嘱ID Then
                    blnLoad = True
                End If
                rsMan.MovePrevious
                If blnLoad Then
                    lng材料ID = 0 '检验试管费用,只收取试管对应的卫材费用
                    If .TextMatrix(lngRow, COL_试管编码) <> "" Then
                        lng材料ID = GetTubeMaterial(.TextMatrix(lngRow, COL_试管编码))
                    End If
                    strSQL = _
                        "Select 诊疗项目ID,收费类别,收费项目ID,收费数量,固有对照,从属项目," & _
                        "   病人科室ID,执行科室ID,跟踪在用,是否变价,标本部位,检查方法,执行标记,费用性质,收费方式,Sum(单价) as 单价 From (" & _
                        " Select c.诊疗项目ID,f.类别 as 收费类别,c.收费项目ID,c.收费数量,c.固有对照,Nvl(c.从属项目,0) as 从属项目," & _
                        " B.病人科室ID,B.执行科室ID,E.跟踪在用,f.是否变价,Decode(Nvl(f.是否变价,0),1,D.缺省价格,D.现价) as 单价," & _
                        " B.标本部位,B.检查方法,B.执行标记,Nvl(c.费用性质,0) as 费用性质,Nvl(c.收费方式,0) as 收费方式,c.适用科室id" & _
                        " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                        " From 诊疗收费关系 C,病人医嘱记录 B,收费项目目录 F,收费价目 D,材料特性 E" & _
                        " Where c.诊疗项目ID+0=B.诊疗项目ID And B.ID=[1] And f.ID=E.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "F", "D", "4", "5", "6") & _
                        " And (B.相关ID is Null And B.执行标记 IN(1,2) And c.费用性质=1" & _
                        "       Or B.标本部位=c.检查部位 And B.检查方法=c.检查方法 And Nvl(c.费用性质,0)=0" & _
                        "       Or (B.检查方法 is Null or b.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=b.诊疗项目ID And Z.操作类型='4')) And Nvl(c.费用性质,0)=0 And c.检查部位 is Null And c.检查方法 is Null)" & _
                        " And c.收费项目ID Not IN(Select 收费细目ID From 病人医嘱计价 Where 医嘱ID=[1])" & _
                        " And c.收费项目ID=f.ID And c.收费项目ID=D.收费细目ID And f.服务对象 IN(2,3)" & _
                        " And (c.收费方式=1 And f.类别='4' And c.收费项目ID=[2] Or Not(c.收费方式=1 And f.类别='4' And [2]<>0))" & _
                        " And (f.撤档时间 is NULL Or f.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And (f.站点='" & gstrNodeNo & "' Or f.站点 is Null) And Sysdate Between D.执行日期 and Nvl(D.终止日期,Sysdate)" & _
                        " And (C.适用科室ID is Null or C.适用科室ID = Nvl(B.执行科室id,[3]) And c.病人来源 = 2)" & _
                        " ) Where Nvl(适用科室id, 0) = Top" & _
                        " Group by 诊疗项目ID,收费类别,收费项目ID,收费数量,固有对照,从属项目," & _
                        "   病人科室ID,执行科室ID,跟踪在用,是否变价,标本部位,检查方法,执行标记,费用性质,收费方式" & _
                        " Order by 费用性质,从属项目"
                    Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!ID), lng材料ID, mlng病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    If Not rsAdd.EOF Then
                        For j = 1 To rsAdd.RecordCount
                            '非药嘱和跟踪在用的卫材专门求执行科室
                            lng执行科室ID = NVL(rsAdd!执行科室ID, 0)
                            If InStr(",5,6,7,", rsAdd!收费类别) > 0 Or rsAdd!收费类别 = "4" And NVL(rsAdd!跟踪在用, 0) = 1 Then
                                lng执行科室ID = Get收费执行科室ID(Val(.TextMatrix(lngRow, COL_病人ID)), Val(.TextMatrix(lngRow, COL_主页ID)), rsAdd!收费类别, rsAdd!收费项目ID, 4, NVL(rsAdd!病人科室id, 0), 0, 2, lng执行科室ID, , , 2)
                            End If
                            
                            mrsPrice.AddNew
                            mrsPrice!医嘱ID = rsMan!ID
                            mrsPrice!相关ID = rsMan!相关ID
                            mrsPrice!诊疗类别 = rsMan!诊疗类别
                            mrsPrice!诊疗项目ID = rsMan!诊疗项目ID
                            mrsPrice!收费类别 = rsAdd!收费类别
                            mrsPrice!收费细目ID = rsAdd!收费项目ID
                            If lng执行科室ID <> 0 Then
                                mrsPrice!执行科室ID = lng执行科室ID
                            Else
                                mrsPrice!执行科室ID = Null
                            End If
                            mrsPrice!在用 = NVL(rsAdd!跟踪在用, 0)
                            mrsPrice!变价 = NVL(rsAdd!是否变价, 0)
                            mrsPrice!数量 = rsAdd!收费数量
                            mrsPrice!单价 = rsAdd!单价
                            mrsPrice!从项 = NVL(rsAdd!从属项目, 0)
                            mrsPrice!固定 = NVL(rsAdd!固有对照, 0)
                            
                            '检查项目的扩展
                            mrsPrice!标本部位 = rsAdd!标本部位
                            mrsPrice!检查方法 = rsAdd!检查方法
                            mrsPrice!执行标记 = NVL(rsAdd!执行标记, 0)
                            mrsPrice!费用性质 = NVL(rsAdd!费用性质, 0)
                            mrsPrice!收费方式 = NVL(rsAdd!收费方式, 0)
                            
                            mrsPrice.Update
                            
                            '存在从属项目的要计价医嘱
                            If mrsPrice!从项 = 1 Then
                                If InStr(strSubItem & ";", ";" & mrsPrice!医嘱ID & "," & mrsPrice!费用性质 & ";") = 0 Then
                                    strSubItem = strSubItem & ";" & mrsPrice!医嘱ID & "," & mrsPrice!费用性质
                                End If
                            End If
                            If NVL(mrsPrice!数量, 0) <> 0 Then blnChange = True '有变化
                            
                            rsAdd.MoveNext
                        Next
                    End If
                End If
            End If
            
            '对存在从项的计价进行处理，保证只有一个主项
            If strSubItem <> "" Then
                If AdjustSubPrice(Mid(strSubItem, 2)) Then blnChange = True
            End If
            rsMan.MoveNext
        Next
                
        .RowData(lngRow) = 1
    End With
    
    LoadPrice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdjustSubPrice(ByVal strSubItem As String) As Boolean
'功能：对存在从项的计价进行处理，保证只有一个主项
'参数：strSubitem=包含从项计价项目的医嘱项目，"医嘱ID,费用性质;..."
'返回：计价内容是否有变化
    Dim rsTmp As ADODB.Recordset
    Dim arrAdvice As Variant, blnChange As Boolean
    Dim intCount As Integer, i As Integer
    Dim strSQL As String
    
    arrAdvice = Split(strSubItem, ";")
    For i = 0 To UBound(arrAdvice)
        intCount = 0
        strSQL = _
            "Select Sum(Decode(从属项目,1,1,0)) as 从项数," & _
            " Max(Decode(从属项目,1,NULL,收费项目ID)) as 主项ID" & _
            " From (Select C.从属项目,C.收费项目ID,C.适用科室id" & _
            " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
            " From 病人医嘱记录 A,诊疗收费关系 C" & _
            " Where A.ID=[1] And A.诊疗项目ID+0=C.诊疗项目ID And Nvl(C.费用性质,0)=[2]" & _
            "       And (A.相关ID is Null And A.执行标记 IN(1,2) And C.费用性质=1" & _
            "       Or A.标本部位=C.检查部位 And A.检查方法=C.检查方法 And Nvl(C.费用性质,0)=0" & _
            "       Or (A.检查方法 is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(C.费用性质,0)=0 And C.检查部位 is Null And C.检查方法 is Null)" & _
            "       And (C.适用科室ID is Null or C.适用科室ID = Nvl(A.执行科室id,[3]) And C.病人来源 = 2)" & _
            ") Where Nvl(适用科室id, 0) = Top"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Split(arrAdvice(i), ",")(0)), Val(Split(arrAdvice(i), ",")(1)), mlng病区ID)
        If Not rsTmp.EOF Then intCount = NVL(rsTmp!从项数, 0)
        If intCount = 0 Then
            '如果现有计价没有从属项目，则取消所有从项属性
            mrsPrice.Filter = "医嘱ID=" & Val(Split(arrAdvice(i), ",")(0)) & " And 费用性质=" & Val(Split(arrAdvice(i), ",")(1))
            Do While Not mrsPrice.EOF
                If mrsPrice!从项 = 1 Then
                    mrsPrice!从项 = 0
                    mrsPrice.Update
                    blnChange = True
                End If
                mrsPrice.MoveNext
            Loop
        Else
            '如果存在从属项目，则除主项以外全部设置为从项
            mrsPrice.Filter = "医嘱ID=" & Val(Split(arrAdvice(i), ",")(0)) & " And 费用性质=" & Val(Split(arrAdvice(i), ",")(1))
            Do While Not mrsPrice.EOF
                If mrsPrice!收费细目ID = Val(NVL(rsTmp!主项ID, 0)) Then '为什么一定要加Val?
                    If mrsPrice!从项 = 1 Then
                        mrsPrice!从项 = 0 '主项肯定有且只有一个
                        mrsPrice.Update
                        blnChange = True
                    End If
                Else
                    If mrsPrice!从项 = 0 Then
                        mrsPrice!从项 = 1
                        mrsPrice.Update
                        blnChange = True
                    End If
                End If
                mrsPrice.MoveNext
            Loop
        End If
    Next
    
    AdjustSubPrice = blnChange
End Function

Private Sub AppendPriceItem()
'功能：补充无对应计价内容项目的记录
    Dim arrPrice As Variant, strPrice As String, i As Long
    Dim lng相关ID As Long, str类别 As String
    Dim lng项目id As Long, str部位 As String
    Dim str方法 As String, int执行标记 As Integer

    mrsPrice.Filter = 0
    Do While Not mrsPrice.EOF
        If mrsPrice!诊疗类别 = "D" And IsNull(mrsPrice!相关ID) Then '检查床旁或术中是一种加收情况
            '加入应有的
            If InStr(strPrice, mrsPrice!医嘱ID & "_") = 0 Then
                If NVL(mrsPrice!执行标记, 0) <> 0 Then
                    '当为床旁或术中执行时，才可以设置加收计价
                    strPrice = strPrice & "," & mrsPrice!医嘱ID & "_0," & mrsPrice!医嘱ID & "_1"
                Else
                    strPrice = strPrice & "," & mrsPrice!医嘱ID & "_0"
                End If
            End If
            '去掉已有的
            If InStr(strPrice, "," & mrsPrice!医嘱ID & "_" & NVL(mrsPrice!费用性质, 0)) > 0 Then
                strPrice = Replace(strPrice, "," & mrsPrice!医嘱ID & "_" & NVL(mrsPrice!费用性质, 0), "")
            End If
        End If
        mrsPrice.MoveNext
    Loop
    
    '剩余的就是没有的
    If strPrice <> "" Then
        arrPrice = Split(Mid(strPrice, 2), ",")
        For i = 0 To UBound(arrPrice)
            mrsPrice.Filter = "医嘱ID=" & Split(arrPrice(i), "_")(0) '该条医嘱可能对应有多种费用性质的计价记录
            If Not mrsPrice.EOF Then
                lng相关ID = NVL(mrsPrice!相关ID, 0)
                str类别 = mrsPrice!诊疗类别
                lng项目id = mrsPrice!诊疗项目ID
                str部位 = NVL(mrsPrice!标本部位)
                str方法 = NVL(mrsPrice!检查方法)
                int执行标记 = NVL(mrsPrice!执行标记, 0)
                
                mrsPrice.AddNew
                mrsPrice!医嘱ID = Val(Split(arrPrice(i), "_")(0))
                If lng相关ID <> 0 Then mrsPrice!相关ID = lng相关ID
                mrsPrice!诊疗类别 = str类别
                mrsPrice!诊疗项目ID = lng项目id
                If str部位 <> "" Then mrsPrice!标本部位 = str部位
                If str方法 <> "" Then mrsPrice!检查方法 = str方法
                mrsPrice!执行标记 = int执行标记
                mrsPrice!费用性质 = Val(Split(arrPrice(i), "_")(1))
                mrsPrice!固定 = 0
                mrsPrice.Update
            End If
        Next
    End If
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'功能：显示当前医嘱行的计价内容(包含相关医嘱的计价项目),同时设置一些编辑属性
    Dim rs诊疗项目 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    Dim str诊疗项目IDs As String, str收费细目IDs As String
    Dim strSQL As String, strAllow As String
    Dim str计价医嘱 As String, i As Long, j As Long
    Dim blnNoFirst As Boolean, lngBegin As Long
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lngComboData As Long, strCombo As String
    Dim strPriceType As String
    
    On Error GoTo errH
    
    With vsPrice
        .Redraw = False
        '清除价目表格
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Editable = flexEDNone
        
        '是否一并给药中的非第一药品行
        If RowIn一并给药(lngRow, lngBegin, 0) Then
            If lngRow > lngBegin Then blnNoFirst = True
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            If blnNoFirst Then
                '一并给药时仅第一行显示给药途径的计价
                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
            Else
                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                    " Or 医嘱ID=" & Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
            End If
        Else
            mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                " Or 相关ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
        End If
        
        If Not mrsPrice.EOF Then
'            If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
'                mrsPrice.Sort = "诊疗类别" '一并给药时显示顺序要求药品在前
'            Else
'                mrsPrice.Sort = ""
'            End If
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(vsAdvice.TextMatrix(lngRow, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        
            '获取诊疗项目,收费细目,价格信息
            For i = 1 To mrsPrice.RecordCount
                str诊疗项目IDs = str诊疗项目IDs & "," & mrsPrice!诊疗项目ID
                If Not IsNull(mrsPrice!收费细目ID) Then
                    str收费细目IDs = str收费细目IDs & "," & mrsPrice!收费细目ID
                End If
                
                '收集允许设置计价的医嘱
                If Not IsNull(mrsPrice!收费细目ID) Then
                    lngComboData = mrsPrice!医嘱ID
                    If NVL(mrsPrice!费用性质, 0) <> 0 Then '负数表示前面附加了一位加收费用性质
                        lngComboData = -1 * Val(mrsPrice!费用性质 & lngComboData)
                    End If
                    '存放:医嘱ID_是否全为固定
                    If InStr(strAllow, "," & lngComboData & "_") = 0 Then
                        strAllow = strAllow & "," & lngComboData & "_" & mrsPrice!固定
                    ElseIf mrsPrice!固定 = 0 Then
                        strAllow = Replace(strAllow, "," & lngComboData & "_1", "," & lngComboData & "_0")
                    End If
                End If
                
                mrsPrice.MoveNext
            Next
            str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
            str收费细目IDs = Mid(str收费细目IDs, 2)
                        
            strSQL = "Select /*+ Rule*/ A.ID,B.名称 as 类别名称,A.名称" & _
                " From 诊疗项目目录 A,诊疗项目类别 B" & _
                " Where A.类别=B.编码 And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
            Set rs诊疗项目 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊疗项目IDs)
            
            '读取是否变价及变价范围等项目信息
            If str收费细目IDs <> "" Then
                strSQL = _
                    " Select A.ID,C.名称 as 类别名称,A.编码,A.名称,A.规格," & _
                    " A.产地,A.计算单位,Nvl(D.住院单位,A.计算单位) as 住院单位," & _
                    " A.费用类型,A.是否变价,Nvl(D.住院包装,1) as 住院包装,A.类别" & _
                    " From 收费项目目录 A,收费项目类别 C,药品规格 D" & _
                    " Where A.类别=C.编码 And A.ID=D.药品ID" & _
                    " And A.类别 IN('5','6','7') And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
                '含卫材
                strSQL = strSQL & " Union ALL " & _
                    " Select A.ID,C.名称 as 类别名称,A.编码,A.名称,A.规格,A.产地," & _
                    " A.计算单位,NULL as 住院单位,A.费用类型,A.是否变价,-Null as 住院包装,A.类别" & _
                    " From 收费项目目录 A,收费项目类别 C" & _
                    " Where A.类别=C.编码 And A.类别 Not IN('5','6','7')" & _
                    " And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
                
                strSQL = _
                    " Select A.ID,A.类别名称,A.编码,A.名称,A.规格,A.产地,A.计算单位,A.住院单位,A.费用类型," & _
                    " A.是否变价,A.住院包装,Sum(B.原价) as 原价,Sum(B.现价) as 现价,Sum(B.缺省价格) as 缺省价格" & _
                    " From (" & strSQL & ") A,收费价目 B Where A.ID=B.收费细目ID" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "3", "4", "5") & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Group by A.ID,A.类别名称,A.编码,A.名称,A.规格,A.产地,A.计算单位,A.住院包装,A.费用类型,A.是否变价,A.住院单位"

                strSQL = _
                    " Select /*+ Rule*/ A.ID,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称,A.规格,A.产地," & _
                    " A.计算单位,A.住院单位,A.费用类型,A.是否变价,A.原价,A.现价,A.缺省价格,A.住院包装" & _
                    " From (" & strSQL & ") A,收费项目别名 B" & _
                    " Where A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[2]"
                Set rs收费细目 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str收费细目IDs, IIF(gbyt药品名称显示 = 0, 1, 3), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级) 'In
            End If
                        
            '确定显示行数
            If str收费细目IDs <> "" Then
                .Rows = .FixedRows + UBound(Split(str收费细目IDs, ",")) + 1
            End If
                                    
            '显示每行内容
            j = .FixedRows
            mrsPrice.MoveFirst
            For i = 1 To mrsPrice.RecordCount
                '确定计价医嘱内容
                rs诊疗项目.Filter = "ID=" & mrsPrice!诊疗项目ID
                If mrsPrice!诊疗类别 = "4" Then
                    str计价医嘱 = "卫生材料-" & rs诊疗项目!名称
                ElseIf InStr(",5,6,7,", mrsPrice!诊疗类别) > 0 Then
                    str计价医嘱 = "药品医嘱-" & rs诊疗项目!名称
                ElseIf mrsPrice!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    str计价医嘱 = "给药途径-" & rs诊疗项目!名称
                ElseIf mrsPrice!诊疗类别 = "E" And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                    str计价医嘱 = "输血途径-" & rs诊疗项目!名称
                ElseIf mrsPrice!诊疗类别 = "E" And InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_类型))) > 0 Then
                    If vsAdvice.TextMatrix(lngRow, COL_类型) = "2" Then
                        str计价医嘱 = "采集方法-" & rs诊疗项目!名称
                    ElseIf Not IsNull(mrsPrice!相关ID) Then
                        str计价医嘱 = "中药煎法-" & rs诊疗项目!名称
                    Else
                        str计价医嘱 = "中药用法-" & rs诊疗项目!名称
                    End If
                ElseIf Not IsNull(mrsPrice!相关ID) Then
                    If mrsPrice!诊疗类别 = "C" Then
                        str计价医嘱 = "检验项目-" & rs诊疗项目!名称
                    ElseIf mrsPrice!诊疗类别 = "D" Then
                        '部位及方法
                        str计价医嘱 = "检查部位-" & NVL(mrsPrice!标本部位) & "(" & NVL(mrsPrice!检查方法) & ")"
                    ElseIf mrsPrice!诊疗类别 = "F" Then
                        str计价医嘱 = "附加手术-" & rs诊疗项目!名称
                    ElseIf mrsPrice!诊疗类别 = "G" Then
                        str计价医嘱 = "麻醉项目-" & rs诊疗项目!名称
                    End If
                Else
                    If NVL(mrsPrice!费用性质, 0) = 1 Then
                        '床旁或术中加收费用
                        str计价医嘱 = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称 & "(" & Decode(NVL(mrsPrice!执行标记, 0), 1, "床旁", 2, "术中", "") & "加收)"
                    Else
                        str计价医嘱 = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称
                    End If
                End If
                
                '可以选择的计价医嘱
                If InStr(",4,5,6,7,", mrsPrice!诊疗类别) = 0 Then
                    lngComboData = mrsPrice!医嘱ID
                    If NVL(mrsPrice!费用性质, 0) <> 0 Then
                        lngComboData = -1 * Val(mrsPrice!费用性质 & lngComboData)
                    End If
                    '条件:没有设置任何收费项目 Or 存在非固定的收费项目(不是全部固定)
                    If InStr(strAllow, "," & lngComboData & "_") = 0 _
                        Or InStr(strAllow, "," & lngComboData & "_0") > 0 Then
                        If InStr(strCombo, "|#" & lngComboData & ";" & str计价医嘱) = 0 Then
                            strCombo = strCombo & "|#" & lngComboData & ";" & str计价医嘱
                        End If
                    End If
                End If
                
                '暂未设置收费关系的不显示,但可以选择
                If Not IsNull(mrsPrice!收费细目ID) Then
                    rs收费细目.Filter = "ID=" & mrsPrice!收费细目ID
                    
                    '显示计价的医嘱内容
                    .TextMatrix(j, COLP_计价医嘱) = str计价医嘱
                    .TextMatrix(j, COLP_医嘱ID) = mrsPrice!医嘱ID
                    .TextMatrix(j, COLP_费用性质) = NVL(mrsPrice!费用性质, 0)
                    .TextMatrix(j, COLP_收费方式) = getChargeMode(Val(NVL(mrsPrice!收费方式, 0)))
                        .Cell(flexcpData, j, COLP_收费方式) = Val(NVL(mrsPrice!收费方式, 0))
                    .TextMatrix(j, COLP_相关ID) = NVL(mrsPrice!相关ID)
                    .TextMatrix(j, COLP_诊疗类别) = mrsPrice!诊疗类别
                    .TextMatrix(j, COLP_诊疗项目ID) = mrsPrice!诊疗项目ID
                        
                    '显示具体计价的项目
                    .TextMatrix(j, COLP_收费类别) = mrsPrice!收费类别
                    .TextMatrix(j, COLP_收费细目ID) = mrsPrice!收费细目ID
                    .TextMatrix(j, COLP_类别) = rs收费细目!类别名称
                    .TextMatrix(j, COLP_收费项目) = rs收费细目!名称
                    If Not IsNull(rs收费细目!产地) Then
                        .TextMatrix(j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目) & "(" & rs收费细目!产地 & ")"
                    End If
                    If Not IsNull(rs收费细目!规格) Then
                        .TextMatrix(j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目) & " " & rs收费细目!规格
                    End If
                    
                    If InStr(",5,6,7,", mrsPrice!诊疗类别) > 0 Then
                        '药品医嘱本身的药品
                        .TextMatrix(j, COLP_单位) = NVL(rs收费细目!住院单位)
                    Else
                        '含其他的药品、卫材计价
                        .TextMatrix(j, COLP_单位) = NVL(rs收费细目!计算单位)
                    End If
                    '药嘱缺省为1,非药嘱药品可设置(售价单位)
                    .TextMatrix(j, COLP_数量) = FormatEx(mrsPrice!数量, 5)
                    
                    '药嘱药品为按1个住院单位计算的价格
                    .TextMatrix(j, COLP_单价) = Format(NVL(mrsPrice!单价), gstrDecPrice)
                    
                    If mrsPrice!收费类别 = "4" Then
                        .TextMatrix(j, COLP_跟踪在用) = Val(NVL(mrsPrice!在用, 0))
                    End If
                    
                    '执行科室
                    lng执行科室ID = NVL(mrsPrice!执行科室ID, 0)
                    '非药嘱药品或跟踪在用的卫材计价可以设置执行科室
                    If InStr(",4,5,6,7,", mrsPrice!诊疗类别) = 0 _
                        And (mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1 Or InStr(",5,6,7,", mrsPrice!收费类别) > 0) Then
                        '以当前值作为缺省重新取有效的执行科室
                        lng病人科室ID = Val(vsAdvice.TextMatrix(lngRow, COL_病人科室ID))
                        lng执行科室ID = Get收费执行科室ID(Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(vsAdvice.TextMatrix(lngRow, COL_主页ID)), _
                            mrsPrice!收费类别, rs收费细目!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
                        '记录是否跟踪在用
                        .Editable = flexEDKbdMouse
                    End If
                    If lng执行科室ID <> 0 Then
                        mrsDept.Filter = "ID=" & lng执行科室ID
                        If Not mrsDept.EOF Then
                            .TextMatrix(j, COLP_执行科室) = mrsDept!名称
                        End If
                    End If
                    .TextMatrix(j, COLP_执行科室ID) = lng执行科室ID
                                        
                    '变价的处理
                    If NVL(rs收费细目!是否变价, 0) = 1 Then
                        If InStr(",5,6,7,", mrsPrice!收费类别) > 0 Then
                            If InStr(",5,6,7,", mrsPrice!诊疗类别) > 0 Then
                                '药嘱药品计算1个住院单位的时价
                                .TextMatrix(j, COLP_单价) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rs收费细目!住院包装, 1), , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                                .TextMatrix(j, COLP_单价) = Format(Val(.TextMatrix(j, COLP_单价)) * NVL(rs收费细目!住院包装, 1), gstrDecPrice)
                            Else
                                '非药嘱药品按零售单位计算
                                .TextMatrix(j, COLP_单价) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, mrsPrice!数量, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                            End If
                        ElseIf mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1 Then
                            '时价卫材价格的药品一样计算
                            .TextMatrix(j, COLP_单价) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, mrsPrice!数量, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else
                            '记录可以输入的价格范围
                            .Cell(flexcpData, j, 0) = 1 '标识为变价(药品不管)
                            .Cell(flexcpData, j, 1) = NVL(rs收费细目!原价, 0)
                            .Cell(flexcpData, j, 2) = NVL(rs收费细目!现价, 0)
                            '也许以前定了变价,现在变价范围变了
                            If .TextMatrix(j, COLP_单价) = "" Then
                                .TextMatrix(j, COLP_单价) = Format(NVL(rs收费细目!缺省价格), gstrDecPrice)
                            ElseIf .TextMatrix(j, COLP_单价) <> "" Then
                                If CheckScope(NVL(rs收费细目!原价, 0), NVL(rs收费细目!现价, 0), NVL(mrsPrice!单价, 0)) <> "" Then
                                    .TextMatrix(j, COLP_单价) = Format(NVL(rs收费细目!缺省价格), gstrDecPrice)
                                End If
                            End If
                            '变价即使固定也可以编辑(包括非跟踪在用的时价卫材医嘱)
                            .Editable = flexEDKbdMouse
                        End If
                    End If

                    '显示医保费用类型
                    If Val(mrsPrice!收费细目ID & "") <> 0 Then
                        strPriceType = GetPriceType(Val(mlng病人ID), Val(mrsPrice!收费细目ID & ""), Val(mint险类), mlng病人性质 = 1)
                    End If
                    '费用类型
                    If strPriceType = "" Then
                        .TextMatrix(j, COLP_费用类型) = NVL(rs收费细目!费用类型)
                    Else
                        .TextMatrix(j, COLP_费用类型) = strPriceType
                    End If
                    
                    .TextMatrix(j, COLP_固定) = mrsPrice!固定
                    .TextMatrix(j, COLP_从项) = IIF(NVL(mrsPrice!从项, 0) = 0, "", "√")
                    
                    '记录用于恢复输入
                    .Cell(flexcpData, j, COLP_计价医嘱) = .TextMatrix(j, COLP_计价医嘱)
                    .Cell(flexcpData, j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目)
                    .Cell(flexcpData, j, COLP_数量) = .TextMatrix(j, COLP_数量)
                    .Cell(flexcpData, j, COLP_单价) = .TextMatrix(j, COLP_单价)
                    .Cell(flexcpData, j, COLP_执行科室) = .TextMatrix(j, COLP_执行科室)
                    
                    '标识固定对照为灰色
                    If mrsPrice!固定 <> 0 Then
                        .Cell(flexcpBackColor, j, .FixedCols, j, .Cols - 1) = &HE0E0E0
                    End If
                    
                    j = j + 1
                End If
                
                mrsPrice.MoveNext
            Next
            
            '设置编辑数据
            '------------------------------------------------------------------
            '需要计价的医嘱选择
            If strCombo <> "" Then
                .ColData(COLP_计价医嘱) = Mid(strCombo, 2)
                .Editable = flexEDKbdMouse '可以选择则可以编辑
            Else
                .ColData(COLP_计价医嘱) = ""
            End If
        End If
        .Row = .FixedRows: .Col = COLP_计价医嘱
        
        '缺省选择计价医嘱(如果可能)
        Call ShowDefaultRow
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub SetSameTime(ByVal lngRow As Long)
'功能：设置其它医嘱行为相同的停止、确认停止、校对,暂停,启用时间
    Dim strTime As String, vPause As Date, strCur As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        strTime = Format(.TextMatrix(lngRow, COL_输入), "yyyy-MM-dd HH:mm")
        strCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        For i = .FixedRows To .Rows - 1
            If i <> lngRow Then
                blnDo = True
                If mint类型 = 3 Then
                    blnDo = .Cell(flexcpData, i, COL_选择) <> Empty
                Else
                    blnDo = Val(.TextMatrix(i, COL_选择)) <> 0
                End If
                
                If blnDo Then
                    If (mint类型 = 1 Or mint类型 = 7) Then  '停止
                        '应>开始执行时间
                        If strTime <= Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        '应>=开嘱时间
                        If strTime < Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        
                        '应>=上次执行时间,因为该时间点已执行,超期收回提前停止的除外
                        If blnDo And .TextMatrix(i, COL_上次执行) <> "" Then
                            If strTime < Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") And _
                                strCur > Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                    ElseIf mint类型 = 2 Then    '确认停止
                        '应>=终止时间
                        If .TextMatrix(i, COL_终止时间) <> "" Then
                            If strTime < Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint类型 = 3 Then    '医嘱校对
                        '应>=min(开嘱时间,开始时间)
                        If Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                            If strTime < Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        Else
                            If strTime < Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint类型 = 5 Then    '暂停医嘱
                        '应>=开始执行时间,因为该时间点尚未执行
                        If strTime < Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        '应>上次执行时间,因为该时间点已执行
                        If .TextMatrix(i, COL_上次执行) <> "" Then
                            If strTime <= Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        '应<执行终止时间,因为该时间点执行有效
                        If .TextMatrix(i, COL_终止时间) <> "" Then
                            If strTime >= Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                        vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 7)
                        If vPause <> CDate(0) Then
                            If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint类型 = 6 Then    '启用医嘱
                        '应>暂停时间
                        vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 6)
                        If vPause <> CDate(0) Then
                            If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        '应<=执行终止时间
                        If .TextMatrix(i, COL_终止时间) <> "" Then
                            If strTime > Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                    End If
                End If
                If blnDo Then
                    .TextMatrix(i, COL_输入) = strTime
                    .Cell(flexcpData, i, COL_输入) = strTime
                End If
            End If
        Next
    End With
End Sub

Private Function GetPauseTime(ByVal lng医嘱ID As Long, ByVal int状态 As Integer) As Date
'功能：读取指定医嘱的暂停时间(该医嘱当前应已暂停)或上次启用时间(如果有)
'参数：int状态=6-暂停,7-启用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(操作时间) as 上次时间 From 病人医嘱状态 Where 操作类型=[2] And 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, int状态)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!上次时间) Then
            GetPauseTime = rsTmp!上次时间
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Pass
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean
    If Button = 2 And mint类型 = 3 Then
        If cbsMain Is Nothing Then Exit Sub
        If mblnPass Then
            blnDo = gobjPass.PassType = G_PASS_TYPE.DT Or gobjPass.PassType = G_PASS_TYPE.YWS Or (gobjPass.PassType = G_PASS_TYPE.MK And gobjPass.PassVersion = "4.0")
        Else
            blnDo = True
        End If
        If blnDo Then Exit Sub
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Function Get病人护理等级医嘱id(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng婴儿序号 As Long, ByVal lng医嘱ID As Long) As Long
'功能：取除开本次作废的护理等级医嘱外的最近的自动停止的护理等级医嘱id
'说明：在作废护理等级时调用，65092需求
'参数：
'      lng病人id
'      lng主页id
'      lng婴儿序号
'      lng医嘱id 本次作废的护理等级医嘱id
'返回：护理等级医嘱id
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    
    On Error GoTo errH
    
    '1.判断校对当前这条医嘱时是不是自动停止前一次的护理等级医嘱，
    '当本次查询有记录并且 护理等级id 为空，则表明是手动停止；其它情况都视为自动停止。
    strSQL = "Select a.护理等级id" & vbNewLine & _
        "From 病人变动记录 A" & vbNewLine & _
        "Where a.病人id = [1] And a.主页id = [2] And a.附加床位 = 0 And" & vbNewLine & _
        "      终止时间 =" & vbNewLine & _
        "      (Select MIN(c.开始时间)" & vbNewLine & _
        "       From 病人医嘱记录 B, 病人变动记录 C" & vbNewLine & _
        "       Where b.病人id = c.病人id And b.主页id = c.主页id And Trunc(c.开始时间, 'MI') = b.开始执行时间 And c.开始原因 = 6 And b.Id = [3])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, lng医嘱ID)
    If Not rsTmp.EOF Then
        If Val(rsTmp!护理等级id & "") = 0 Then Exit Function
    End If
    
    '2.先取出可以将其启用的 护理等级医嘱 的 医嘱id
    strSQL = "Select 医嘱id" & vbNewLine & _
        "From (Select a.Id As 医嘱id" & vbNewLine & _
        "       From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
        "       Where a.诊疗项目id = b.Id And a.诊疗类别 = 'H' And b.操作类型 = '1' And a.病人id = [1] And a.主页id = [2] And Nvl(a.婴儿, 0) = [3] And" & vbNewLine & _
        "             a.医嘱状态 In (8, 9) And a.Id <> [4] And a.开始执行时间 < (Select 开始执行时间 From 病人医嘱记录 Where ID = [4])" & vbNewLine & _
        "       Order By a.开始执行时间 Desc)" & vbNewLine & _
        "Where Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, lng婴儿序号, lng医嘱ID)
    If rsTmp.EOF Then Exit Function
   
    Get病人护理等级医嘱id = Val(rsTmp!医嘱ID & "")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub zlPASSMap()
'功能:设置Pass VsAdvie及列映射
'注意:删除或修改下面列中数据时，请检查合理用药部件中的关联处理。
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "合理用药监测", True)
        mblnPass = Not gobjPass Is Nothing And Not mobjPassMap Is Nothing
    End If
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_护士校对
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .objCmdBar = cmdAlley
            Set .VSCOL = .GetVSCOL(COL_ID, COL_相关ID, COL_诊疗类别, COL_诊疗项目ID, COL_收费细目ID, col_医嘱内容, COL_期效, COL_单量, COL_单量单位, _
                        COL_用法, , COL_婴儿, COL_开嘱时间, COL_开嘱医生, COL_开始时间, COL_开嘱科室ID, COL_终止时间, COL_频率, COL_频率次数, COL_频率间隔, _
                        COL_间隔单位, COL_警示, COL_序号, , , COL_病人ID, COL_主页ID, COL_选择, COL_执行性质, COL_标本部位)
        End With
        mblnPass = gobjPass.zlPassCheck(mobjPassMap)
    End If
End Sub

Private Sub GetDefaultDeptPar(ByVal lng病人科室ID As Long)
'功能：获取缺省参数
    mlng中药房 = Val(zlDatabase.GetPara("住院缺省中药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
    mlng西药房 = Val(zlDatabase.GetPara("住院缺省西药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
    mlng成药房 = Val(zlDatabase.GetPara("住院缺省成药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
    mlng发料部门 = Val(zlDatabase.GetPara("住院缺省发料部门", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
End Sub

Private Sub SetFilterTime()
'功能：设置界面参数的过滤条件初始状态
    Dim strTmp As String
    
    If mint类型 = 3 Then
        strTmp = mstr缺省校对时间
        If Left(strTmp, 1) = "0" Then
            optOper(e当前时间).value = True '当前时间
            lblS.Visible = False
            lblB.Visible = False
            cboTime(e早于).Visible = False
            cboTime(e晚于).Visible = False
        Else
            optOper(e开始时间).value = True
            lblS.Visible = True
            lblB.Visible = True
            cboTime(e早于).Visible = True
            cboTime(e晚于).Visible = True
        End If
        cboTime(e早于).ListIndex = IIF(Mid(strTmp, 2, 1) = "1", 1, 0)
        cboTime(e晚于).ListIndex = IIF(Mid(strTmp, 3, 1) = "1", 1, 0)
    ElseIf (mint类型 = 1 Or mint类型 = 7) Then
        strTmp = mstr缺省停止时间
        If Left(strTmp, 1) = "1" Then
            optStop(e上次执行时间).value = True '上次执行时间
            chkNoSend.Visible = False
            chkRollSend.Visible = False
        Else
            optStop(e指定时间).value = True
            chkNoSend.value = IIF(Mid(strTmp, 2, 1) = "1", 1, 0)
            chkRollSend.value = IIF(Mid(strTmp, 3, 1) = "1", 1, 0)
        End If
    End If
End Sub

Private Sub SetSame原因(ByVal lngRow As Long)
'功能：设置其它医嘱行为相同的终止原因
    Dim str原因 As String
    Dim i As Long
    
    Call vsAdvice_AfterEdit(lngRow, COL_终止原因)
    
    If Not VsfOnlySelOneRow(lngRow) Then
    
        str原因 = vsAdvice.TextMatrix(lngRow, COL_终止原因)
        
        If MsgBox("要设置所有已选择的医嘱都设为这个停嘱原因：" & str原因 & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If i <> lngRow Then
                    If Val(.TextMatrix(i, COL_选择)) <> 0 Then
                         .TextMatrix(i, COL_终止原因) = str原因
                        .Cell(flexcpData, i, COL_终止原因) = str原因
                    End If
                End If
            Next
        End With
    End If
End Sub

Private Sub InitRecordSet(ByRef rsAdviceTmp As ADODB.Recordset, ByRef rsMsgRow As ADODB.Recordset, ByRef rs输血 As ADODB.Recordset)
'功能：初始化本地记录集
    Set rsAdviceTmp = New ADODB.Recordset
    rsAdviceTmp.Fields.Append "病人ID", adBigInt
    rsAdviceTmp.Fields.Append "主页ID", adBigInt
    rsAdviceTmp.Fields.Append "医嘱IDs", adVarChar, 4000
    rsAdviceTmp.CursorLocation = adUseClient
    rsAdviceTmp.LockType = adLockOptimistic
    rsAdviceTmp.CursorType = adOpenStatic
    
    Set rsMsgRow = New ADODB.Recordset
    rsMsgRow.Fields.Append "病人ID", adBigInt
    rsMsgRow.Fields.Append "主页ID", adBigInt
    rsMsgRow.Fields.Append "行号", adBigInt
    rsMsgRow.Fields.Append "操作类型", adBigInt '1－停止，2－作废，3－校对通过，4-校对疑问
    rsMsgRow.Fields.Append "当前病情", adVarChar, 4000
    rsMsgRow.CursorLocation = adUseClient
    rsMsgRow.LockType = adLockOptimistic
    rsMsgRow.CursorType = adOpenStatic
    rsMsgRow.Open
    
    Set rs输血 = New ADODB.Recordset
    With rs输血
        .Fields.Append "医嘱ID", adBigInt
        .Fields.Append "类型", adInteger '3－校对，4－作废
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Private Sub Check处方审查()
'功能：调用审方接口判断当前医嘱是不是允许发送
    Dim i As Long
    Dim str给药IDs As String '传入到接口中的参数
    Dim strOut医嘱IDs As String '不能够发送的主医嘱ID
    Dim strErr As String
    Dim lng医嘱ID As Long
    Dim str医嘱内容 As String
    Dim str病人 As String
    Dim lngLastPatiID As Long
    Dim lngLastPageID As Long
    Dim rsTmp As ADODB.Recordset
    Dim j As Long
    Dim str药行医嘱IDs As String
    
    On Error GoTo errH
    
    If Not gbln审方系统 Then Exit Sub
    
    With vsAdvice
        '校对存在多病人模式
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty Then
                If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Then
                    If lngLastPatiID <> Val(.TextMatrix(i, COL_病人ID)) Then
                        lngLastPatiID = Val(.TextMatrix(i, COL_病人ID))
                        lngLastPageID = Val(.TextMatrix(i, COL_主页ID))
                    
                        Set rsTmp = Nothing
                        Call gobjPass.ZLPharmReviewResultView(lngLastPatiID, lngLastPageID, rsTmp, strErr)
                        If Not rsTmp Is Nothing Then
                            If Not rsTmp.EOF Then
                                For j = 1 To rsTmp.RecordCount
                                    If InStr("," & strOut医嘱IDs & ",", "," & rsTmp!相关ID & ",") = 0 Then
                                        strOut医嘱IDs = strOut医嘱IDs & "," & rsTmp!相关ID
                                    End If
                                    str药行医嘱IDs = str药行医嘱IDs & "," & rsTmp!医嘱ID
                                    rsTmp.MoveNext
                                Next
                            End If
                        End If
                        
                        
                    End If
                End If
            End If
        Next
   
        If strOut医嘱IDs <> "" Then
            '取消选择
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) <> 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    lng医嘱ID = IIF(0 = Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    If InStr("," & strOut医嘱IDs & ",", "," & lng医嘱ID & ",") > 0 Then
                        Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                        .Cell(flexcpData, i, COL_选择) = 0
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            If InStr("," & str药行医嘱IDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                                str医嘱内容 = str医嘱内容 & vbCrLf & .TextMatrix(i, col_医嘱内容)
                            End If
                        End If
                    End If
                End If
            Next
            If str医嘱内容 <> "" Then
                Call MsgBox("以下医嘱未通过处方审查，不能校对：" & str医嘱内容, vbInformation, Me.Caption)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
