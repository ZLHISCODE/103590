VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBloodReceivesRecord 
   Caption         =   "血液接收登记"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16050
   Icon            =   "frmBloodReceivesRecord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   16050
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "门诊"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   34
         Top             =   -10
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "住院"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   33
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "使用场合"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox pic11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   4440
      ScaleHeight     =   3135
      ScaleWidth      =   11055
      TabIndex        =   28
      Top             =   3120
      Width           =   11055
      Begin VB.PictureBox picTransit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10815
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   10815
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            ScaleHeight     =   255
            ScaleWidth      =   2175
            TabIndex        =   37
            Top             =   120
            Width           =   2175
            Begin VB.OptionButton optTransit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   39
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optTransit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "当前时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpTransit 
            Height          =   330
            Left            =   3240
            TabIndex        =   40
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42618
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "转接时间"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.PictureBox pic4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   105
         ScaleHeight     =   1695
         ScaleWidth      =   3615
         TabIndex        =   29
         Top             =   240
         Width           =   3615
         Begin VB.CheckBox chk2 
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Width           =   255
         End
         Begin VSFlex8Ctl.VSFlexGrid VSF2 
            Height          =   1575
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   2580
            _cx             =   4551
            _cy             =   2778
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
   End
   Begin VB.PictureBox pic7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2295
      ScaleWidth      =   3255
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin XtremeSuiteControls.TabControl tbcthis 
         Height          =   1335
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   2355
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox Pic5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7920
      ScaleHeight     =   2535
      ScaleWidth      =   7455
      TabIndex        =   14
      Top             =   360
      Width           =   7455
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10815
         TabIndex        =   16
         Top             =   0
         Width           =   10815
         Begin VB.PictureBox pic8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            ScaleHeight     =   255
            ScaleWidth      =   2175
            TabIndex        =   25
            Top             =   120
            Width           =   2175
            Begin VB.OptionButton opt1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "当前时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opt1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   26
               Top             =   0
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker DTP3 
            Height          =   330
            Left            =   3240
            TabIndex        =   17
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42618
         End
         Begin VB.Label lbl6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "接收时间"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   0
         ScaleHeight     =   2055
         ScaleWidth      =   5175
         TabIndex        =   15
         Top             =   1320
         Width           =   5175
         Begin VB.CheckBox chk3 
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin VSFlex8Ctl.VSFlexGrid VSF1 
            Height          =   1575
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   2580
            _cx             =   4551
            _cy             =   2778
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   360
      ScaleHeight     =   7455
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      Begin zlPublicBlood.usrCardPeople UCP 
         Height          =   3975
         Left            =   240
         TabIndex        =   42
         Top             =   3240
         Width           =   3015
         _extentx        =   5318
         _extenty        =   7011
      End
      Begin VB.Frame Fra1 
         Height          =   2895
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   3855
         Begin VB.CommandButton cmdOper 
            Height          =   240
            Left            =   3330
            Picture         =   "frmBloodReceivesRecord.frx":07AA
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "编辑(F4)"
            Top             =   1710
            Width           =   255
         End
         Begin VB.TextBox txtOper 
            Height          =   300
            Left            =   960
            TabIndex        =   5
            Top             =   1680
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpEdtime 
            Height          =   300
            Left            =   960
            TabIndex        =   4
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42635
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            Text            =   "今天内"
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "批量提取"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2520
            TabIndex        =   9
            Top             =   2430
            Width           =   1100
         End
         Begin VB.CheckBox chk1 
            Caption         =   "批量处理"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2490
            Width           =   1095
         End
         Begin VB.CommandButton cmd1 
            Caption         =   "刷新"
            Height          =   350
            Left            =   2520
            TabIndex        =   7
            Tag             =   "√"
            Top             =   2040
            Width           =   1100
         End
         Begin VB.ComboBox cboDepart 
            Height          =   300
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpSttime 
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin VB.Label lbl5 
            Caption         =   "取 血 人"
            Height          =   255
            Left            =   165
            TabIndex        =   22
            Top             =   1725
            Width           =   735
         End
         Begin VB.Label lbl1 
            Caption         =   "科    室"
            Height          =   255
            Left            =   165
            TabIndex        =   13
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lbl2 
            Caption         =   "取血时间"
            Height          =   255
            Left            =   165
            TabIndex        =   12
            Top             =   645
            Width           =   735
         End
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "病人列表"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3375
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   9795
      Width           =   16050
      _ExtentX        =   28310
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2461
            MinWidth        =   882
            Picture         =   "frmBloodReceivesRecord.frx":08A0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22357
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   1200
      Top             =   120
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
            Picture         =   "frmBloodReceivesRecord.frx":1386
            Key             =   "拒绝转接"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":1720
            Key             =   "完成接收"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":7F82
            Key             =   "拒绝接收"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":E7E4
            Key             =   "等待用血"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPeoPle 
      Bindings        =   "frmBloodReceivesRecord.frx":15046
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBloodReceivesRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSys As Long   '调用模块号
Private mlngMoudle As Long '模块号
Private mstr开始时间 As String
Private mstr结束时间 As String
Private mstr填写人 As String
Private mstrPrivs As String '权限串
Private mblnButtonChecked As Boolean
Private mblnTextChecked As Boolean
Private mblnSizeChecked As Boolean
Private mblnStatuChecked As Boolean
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private WithEvents mclsvsf1 As clsVsf
Attribute mclsvsf1.VB_VarHelpID = -1
Private mRsBR(0 To 1) As ADODB.Recordset '查询到的病人的记录集
Private mtbcThisIndex As Long '当前选中的选项卡
Private mintOutPreTime As Long  '当前选中的时间
Private marrPreCardID(0 To 1) As String
Private mfrmMain As Object
Private mbln转接 As Boolean
Private mint场合 As Integer '0:门诊和住院;1-门诊;2-住院
Public mblnBloodReceivesIsOpen As Boolean '非模态状态下，判断窗体是否开启
Private mrs部门 As ADODB.Recordset
Private mblnHavePrivs As Boolean    '是否有所有科室的权限，且病人属于门诊或住院部门
'刷新数据时的过滤条件
Private Type Type_Filter
    DeptID As Long
    TimeIndex As Integer
    BeginTime As String
    EndTime As String
    Oper As String
End Type
Private marrFilter(0 To 1) As Type_Filter

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始化处理
    
    Call CommandBarInit(cbsMain)
    '菜单定义:包括公共部份
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '编辑
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "接收", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "回退")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Transfer, "转接", True)
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    
    mblnButtonChecked = True
    mblnTextChecked = True
    mblnSizeChecked = True
    mblnStatuChecked = True
    
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "接收"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "转接"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        Set objCustom = .Add(xtpControlCustom, conMenu_View_FindType, "场合")
        objCustom.Handle = fraType.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyF, conMenu_View_Find           '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext         '继续查找
        .Add 0, vbKeyF5, conMenu_View_Refresh      '刷新
    End With
    
    Call gobjDatabase.ShowReportMenu(Me, 100, p血液接收登记, mstrPrivs)
    InitCommandBar = True
    
    Exit Function
ErrHand:
End Function

Private Sub cboDepart_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    Dim lngi As Long
    Dim blnisread  As Boolean
    Dim rs部门 As ADODB.Recordset
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub '不允许输入"'"
    blnisread = False
    If KeyAscii = vbKeyReturn Then
        For lngi = 0 To cboDepart.ListCount - 1
            If cboDepart.List(lngi) Like cboDepart.Text & "*" Or cboDepart.List(lngi) Like "*" & cboDepart.Text & "*" Then
                cboDepart.Text = cboDepart.List(lngi)
                cboDepart.Tag = cboDepart.ListIndex
                cboDepart.ListIndex = lngi
                blnisread = True
                Exit For
            End If
        Next
        If blnisread = False Then
            Call CopyRecord(mrs部门, rs部门)
            rs部门.Filter = "简码 like '" & cboDepart.Text & "%'"
            If rs部门.RecordCount > 0 Then
                For lngi = 0 To cboDepart.ListCount - 1
                    If cboDepart.ItemData(lngi) = rs部门!id Then
                        cboDepart.ListIndex = lngi
                        cboDepart.Tag = lngi
                        blnisread = True
                        Exit For
                    End If
                Next
            End If
        End If
        If blnisread = False Then
            cboDepart.ListIndex = IIf(Val(cboDepart.Tag) < 0, 0, Val(cboDepart.Tag))
            Exit Sub
        End If
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cboDepart_Validate(Cancel As Boolean)
    '检查科室是否正确
    Dim lngi As Long
    Dim blnIsSelect As Boolean
    blnIsSelect = False
    If mblnHavePrivs = False Then Exit Sub '如果病人没有相关权限则退出
    For lngi = 0 To cboDepart.ListCount - 1
        If cboDepart.Text = cboDepart.List(lngi) Or cboDepart.Text = "所有部门" Then
            blnIsSelect = True
            cboDepart.Tag = cboDepart.ListIndex
            Exit For
        End If
    Next
    If blnIsSelect = False Then
        cboDepart.ListIndex = IIf(Val(cboDepart.Tag) < 0, 0, Val(cboDepart.Tag))
        Exit Sub
    End If
End Sub

Private Sub cboTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    If cboTime.ListIndex < 0 Then Exit Sub
    If mintOutPreTime = cboTime.ListIndex And cboTime.ListIndex <> cboTime.ListCount - 1 Then Exit Sub
    intDateCount = cboTime.ItemData(cboTime.ListIndex)
    datCurr = Format(gobjDatabase.Currentdate, "yyyy-MM-dd")
    If Me.Visible Then
        If intDateCount = -1 Then
        ElseIf intDateCount = 0 Then
            mstr开始时间 = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mstr结束时间 = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mstr开始时间 = Format(datCurr - intDateCount, "yyyy-MM-dd 00:00:00")
            mstr结束时间 = Format(datCurr, "yyyy-MM-dd 23:59:59")
        End If
        dtpSttime.Value = Format(mstr开始时间, "YYYY-MM-DD HH:mm")
        dtpEdtime.Value = Format(mstr结束时间, "YYYY-MM-DD HH:mm")
        dtpSttime.Enabled = (cboTime.ItemData(cboTime.ListIndex) = -1): dtpEdtime.Enabled = dtpSttime.Enabled
    End If
    mintOutPreTime = cboTime.ListIndex
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng病人ID As Long, lng主页id As Long
    Dim strTmp As String
    
    Select Case Control.id
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(2, IIf(mtbcThisIndex = 0, VSF1, VSF2), IIf(mtbcThisIndex = 0, "待接受血液", "已接收血液"))
        Case conMenu_File_Print
            Call zlRptPrint(1, IIf(mtbcThisIndex = 0, VSF1, VSF2), IIf(mtbcThisIndex = 0, "待接受血液", "已接收血液"))
        Case conMenu_Edit_Audit: '接收
            If mbln转接 = False Then '正常接收
                If ExecuteCommand("接收数据") = True Then Call ExecuteCommand("刷新数据")
                chk3.Value = 0
            Else '转接接收
                If ExecuteCommand("转接") = True Then Call ExecuteCommand("刷新数据")
            End If
        Case conMenu_Edit_Untread: '回退
            If ExecuteCommand("回退") = True Then Call ExecuteCommand("刷新数据")
            chk2.Value = 0
        Case conMenu_View_Refresh: '刷新
            cmd1_Click
         Case conMenu_View_Find, conMenu_View_FindNext '查找，继续查找
            Call UCP.FindPatiByVbKey(Control.id = conMenu_View_FindNext)
        Case conMenu_View_ToolBar_Button '标准按钮
            mblnButtonChecked = Not mblnButtonChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Text '文本标签
            mblnTextChecked = Not mblnTextChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Size '大图标
            mblnSizeChecked = Not mblnSizeChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_StatusBar '状态栏
            mblnStatuChecked = Not mblnStatuChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_Help_Help              '帮助主题
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((100) / 100))
        Case conMenu_Help_Web_Home 'Web上的中联
            Call gobjComlib.zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum         'Web上的论坛
            Call gobjComlib.zlWebForum(Me.hWnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call gobjComlib.zlMailTo(Me.hWnd)
        Case conMenu_Help_About '关于
            Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Exit '退出
            Unload Me
        Case conMenu_Manage_Transfer '转接
            mbln转接 = Not mbln转接 '转换接收模式
            If mbln转接 Then optTransit(0).Value = True
            Call SetVsf2State
        Case Else
            If Between(Control.id, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                strTmp = UCP.strReturn
                If strTmp <> "" Then
                    strTmp = Split(strTmp, "'")(1)
                    lng病人ID = Split(strTmp, "-")(0)
                    lng主页id = Split(strTmp, "-")(1)
                    '执行发布到当前模块的报表
                    Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "病人ID=" & lng病人ID & ",就诊ID=" & lng主页id)
                End If
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_Edit_Audit: '接收
            If mbln转接 = False Then
                Control.Caption = "接收"
                Control.Visible = IsPrivs(mstrPrivs, "接收数据")
                Control.Enabled = IIf(mtbcThisIndex = 0, True, False) And Control.Visible
            Else
                Control.Caption = "保存"
                Control.Enabled = IsClick
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '回退
            Control.Visible = IsPrivs(mstrPrivs, "回退数据")
            Control.Enabled = IIf(mtbcThisIndex = 0, False, True) And Control.Visible And Not mbln转接
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Transfer: '转接
            Control.Visible = IsPrivs(mstrPrivs, "转接数据")
            Control.Enabled = IIf(mtbcThisIndex = 0, False, True) And Control.Visible
            Control.Checked = mbln转接
            picTransit.Visible = mbln转接
            Call pic11_Resize
        Case conMenu_View_ToolBar_Button
            Control.Checked = mblnButtonChecked
        Case conMenu_View_ToolBar_Text
            Control.Checked = mblnTextChecked
        Case conMenu_View_ToolBar_Size
            Control.Checked = mblnSizeChecked
        Case conMenu_View_StatusBar
            stbThis.Visible = mblnStatuChecked
            Control.Checked = mblnStatuChecked
    End Select
End Sub

Private Sub chk1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub chk2_Click()
    '全选、全清
    Dim lngi As Long
    
    For lngi = 1 To VSF2.Rows - 1
        If mbln转接 = False Then
            If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("病人id"))) <> 0 And (UserInfo.姓名 = VSF2.TextMatrix(lngi, VSF2.ColIndex("接收人")) Or UserInfo.姓名 = VSF2.TextMatrix(lngi, VSF2.ColIndex("核收人"))) Then
                VSF2.TextMatrix(lngi, VSF2.ColIndex("选择")) = chk2.Value
            End If
        Else
            If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("状态"))) <> 1 And Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("状态"))) <> 3 And Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("病人id"))) <> 0 Then
                VSF2.TextMatrix(lngi, VSF2.ColIndex("选择")) = chk2.Value
            End If
        End If
    Next
End Sub

Private Sub chk2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '提示信息
    Dim strInfo As String
    If mbln转接 = True Then
        strInfo = "可对已接收的血液进行转接" & vbCrLf & "条件：接收时间大于上次时间；病人当前科室发生变化"
    Else
        strInfo = "可对自己接收或核对的数据进行回退"
    End If
    gobjCommFun.ShowTipInfo chk2.hWnd, strInfo
End Sub

Private Sub chk3_Click()
    '全选、全清
    Dim lngi As Long
    For lngi = 1 To VSF1.Rows - 1
        If Val(VSF1.TextMatrix(lngi, VSF1.ColIndex("病人id"))) <> 0 Then
            VSF1.TextMatrix(lngi, VSF1.ColIndex("选择")) = chk3.Value
        End If
    Next
End Sub

Private Sub cmd1_Click()
    '刷新数据
    If Format(dtpSttime.Value, "YYYY-MM-DD HH:mm") > Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm") Then
        MsgBox "过滤条件中的开始时间不能大于结束时间，请调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnHavePrivs = False Then Exit Sub
    mstr开始时间 = Format(dtpSttime.Value, "YYYY-MM-DD HH:mm")
    mstr结束时间 = Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm")
    
    With marrFilter(mtbcThisIndex)
        .DeptID = cboDepart.ItemData(cboDepart.ListIndex)
        .TimeIndex = cboTime.ListIndex
        .BeginTime = mstr开始时间
        .EndTime = mstr结束时间
        .Oper = txtOper.Text
    End With
    '清除VSF1和VSF2上的数据
    If mtbcThisIndex = 0 Then
        VSF1.Rows = 2
        VSF1.RowData(1) = 0
        VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
    Else
        VSF2.Rows = 2
        VSF2.RowData(1) = 0
        VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
    End If
    
    Call ExecuteCommand("基础病人查询") ','' as 床号,'' as 姓名,'' as 年龄
End Sub

Private Sub ShowVsf(lngstatu As Long, Optional rs As ADODB.Recordset, Optional strP As String)
    '功能：在vsf上显示响应数据，分为多选和单选，处理方式不同
    '参数：lngstatu:0-待接收血液、1-已接收血液，rs和strp都是ucp的返回数据，根据情况选择,两个可选参数只能选择其中一个，两个参数都存在，则不做处理
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strSql1 As String
    Dim lng病人ID As Long, lng主页id As Long
    Dim strValues As String, strTmp As String
    Dim blnBatch As Boolean
    
    On Error GoTo ErrHand
    
    If chk1.Value = Checked Then  '多选的情况下
        blnBatch = True
        If rs Is Nothing Then Exit Sub
        If rs.State = adStateClosed Then Exit Sub
        If rs.RecordCount = 0 Then Exit Sub
    Else
        blnBatch = False
        If strP = "" Then Exit Sub
    End If
    
    '选项卡不同查询内容不同
    If lngstatu = 0 Then '不同的选项卡过滤条件不相同
        strSql1 = " And e.发送时间+0 between [3] and [4] And nvl(e.接收状态,0) =0  "
        If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
            strSql1 = strSql1 & " And e.取血人=[6]"
        End If
    Else
        strSql1 = " And e.接收时间+0 between [3] and [4] And nvl(e.接收状态,0)<>0 "
        If marrFilter(mtbcThisIndex).DeptID <> -1 Then
            strSql1 = strSql1 & " And E.执行科室id=[5]"
        End If
        If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
            strSql1 = strSql1 & " And e.接收人=[6]"
        End If
    End If
    
    If optOccasion(0).Value = True Then
        strSQL = _
            " Select " & IIf(blnBatch = True, " /*+ CARDINALITY(T 10) */ ", "") & " e.收发id, d.病人id, d.主页id, 0 As 选择, g.出院病床 As 床号, g.姓名, g.性别,g.出院科室ID 当前科室ID, a.名称 As 血液名称, a.规格, f.血袋编号, f.Abo, f.Rh," & vbNewLine & _
            "       Decode(e.接收状态, 0, '正在接收', 1, '已经接收', 2, '拒绝接收', '转科接收') As 接收状态, " & IIf(lngstatu = 1, "h.名称", "''") & " As 执行部门, e.接收人, to_char(e.接收时间,'yyyy-mm-dd HH24:mi:ss') as 接收时间, e.核收人, e.执行科室id," & vbNewLine & _
            "       e.取血人, e.拒绝原因" & vbNewLine & _
            " From 收费项目目录 a, 血液品种 b, 血液规格 c" & IIf(lngstatu = 1, ",部门表 h", "") & ", 血液收发记录 f, 血液发送记录 e, 血液配血记录 d, 病案主页 g" & IIf(blnBatch = True, ",Table(f_Num2list2([1])) T", "") & vbNewLine & _
            " Where a.Id = c.规格id And c.品种id = b.品种id And c.规格id = f.血液id And f.Id = e.收发id" & IIf(lngstatu = 1, " And E.执行科室id=h.id(+) ", " ") & " And e.配发id = d.Id And d.病人id = g.病人id And" & vbNewLine & _
            "      d.主页id = g.主页id "
        If blnBatch = True Then
            strSQL = strSQL & " And g.病人ID=T.C1 and g.主页ID=T.C2"
        Else
            strSQL = strSQL & " And g.病人id = [1] And g.主页id = [2]"
        End If
    Else
        strSQL = _
            " Select " & IIf(blnBatch = True, " /*+ CARDINALITY(T 10) */ ", "") & " e.收发id, d.病人id, d.Id 主页id, 0 As 选择, '' As 床号, g.姓名, g.性别,g.执行部门ID 当前科室ID, a.名称 As 血液名称, a.规格, f.血袋编号, f.Abo, f.Rh," & vbNewLine & _
            "       Decode(e.接收状态, 0, '正在接收', 1, '已经接收', 2, '拒绝接收', '转科接收') As 接收状态, " & IIf(lngstatu = 1, "h.名称", "''") & " As 执行部门, e.接收人, to_char(e.接收时间,'yyyy-mm-dd HH24:mi:ss') as 接收时间, e.核收人, e.执行科室id," & vbNewLine & _
            "       e.取血人, e.拒绝原因" & vbNewLine & _
            " From 收费项目目录 a, 血液品种 b, 血液规格 c" & IIf(lngstatu = 1, ",部门表 h", "") & ", 血液收发记录 f, 血液发送记录 e, 血液配血记录 d, 病人医嘱记录 k, 病人挂号记录 g" & IIf(blnBatch = True, ",Table(f_Num2list2([1])) T", "") & vbNewLine & _
            " Where a.Id = c.规格id And c.品种id = b.品种id And c.规格id = f.血液id And f.Id = e.收发id " & IIf(lngstatu = 1, " And E.执行科室id=h.id(+) ", " ") & " And e.配发id = d.Id And d.申请id = k.Id And" & vbNewLine & _
            "       k.挂号单 = g.No And k.病人id = g.病人id And k.诊疗类别 = 'K' "
        If blnBatch = True Then
            strSQL = strSQL & " And g.病人ID=T.C1 and g.Id=T.C2"
        Else
            strSQL = strSQL & " And g.病人id = [1] And g.Id = [2]"
        End If
    End If
    strSQL = strSQL & strSql1
    Screen.MousePointer = 11
    If blnBatch Then '多选的情况下
        strValues = ""
        Do While Not rs.EOF
            strTmp = rs.Fields("ID").Value
            lng病人ID = Val(Split(strTmp, "-")(0))
            lng主页id = Val(Split(strTmp, "-")(1))
            strValues = strValues & "," & lng病人ID & ":" & lng主页id
            rs.MoveNext
        Loop
        If Left(strValues, 1) = "," Then strValues = Mid(strValues, 2)
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "血液信息", strValues, "", CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), marrFilter(mtbcThisIndex).DeptID, Trim(marrFilter(mtbcThisIndex).Oper))
    Else
        lng病人ID = Val(Split(Split(strP, "'")(1), "-")(0))
        lng主页id = Val(Split(Split(strP, "'")(1), "-")(1))
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "血液信息", lng病人ID, lng主页id, CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), marrFilter(mtbcThisIndex).DeptID, Trim(marrFilter(mtbcThisIndex).Oper))
    End If
    If lngstatu = 0 Then
        Call mclsVsf.LoadGrid(rsTemp)
    Else
        Call mclsvsf1.LoadGrid(rsTemp)
        Call SetVsf2State
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetVsf2State()
    '根据vsf2上的数据的情况，确认其数据状态，还有根据状态改变每条数据的图标。
    Dim lngi As Long, lng状态 As Long
    For lngi = 1 To VSF2.Rows - 1
        VSF2.TextMatrix(lngi, VSF2.ColIndex("选择")) = 0
        VSF2.Cell(flexcpPicture, lngi, 4, lngi, 4) = Nothing
        If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("病人id"))) <> 0 Then
            If VSF2.TextMatrix(lngi, VSF2.ColIndex("接收状态")) = "拒绝接收" Then
                lng状态 = 3
            ElseIf VSF2.TextMatrix(lngi, VSF2.ColIndex("接收状态")) = "已经接收" Then
                lng状态 = 2
            ElseIf VSF2.TextMatrix(lngi, VSF2.ColIndex("接收状态")) = "转科接收" Then
                lng状态 = 4
            End If
            '当接收状态为正常接收和转科接收，但执行科室ID未改变时(病人未转科)，不允许转接
            If mbln转接 = True Then '转接模式会将不允许转接的数据标注为"!"
                If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("执行科室id"))) = Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("当前科室ID"))) Then
                    lng状态 = 1
                    VSF2.Cell(flexcpForeColor, lngi, 1, lngi, VSF2.Cols - 1) = vbGrayText
                End If
            Else
                VSF2.Cell(flexcpForeColor, lngi, 1, lngi, VSF2.Cols - 1) = vbBlack
            End If
            VSF2.TextMatrix(lngi, VSF2.ColIndex("状态")) = lng状态
            If lng状态 <> 0 Then
                VSF2.Cell(flexcpPicture, lngi, VSF2.ColIndex("图标"), lngi, VSF2.ColIndex("图标")) = ils16.ListImages(lng状态).Picture
            End If
        End If
    Next
End Sub

Private Sub GetSendorReceivePeople(ByVal objControl As TextBox, Optional ByVal StrInput As String = "")
    '功能：根据时间范围查找取血人和接收人
    Dim strSQL As String, strSQLNew As String
    Dim rsUser  As ADODB.Recordset
    Dim lngDeptID As Long, strWhere As String
    Dim vPoint As RECT, blnCancel As Boolean
    On Error GoTo ErrHand
    
    '过滤条件，判断按照什么方式过滤数据：姓名、编号、简码
    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And C.编号 Like [4]"
         ElseIf gobjCommFun.IsCharAlpha(StrInput) Then
            strWhere = " And C.简码 Like [4]"
            StrInput = UCase(StrInput)
         Else
            strWhere = " And C.姓名 Like [4]"
         End If
    End If
    
    '查询语句
    If mtbcThisIndex = 0 Then
        strSQL = "Select Distinct 取血人 as 姓名 From 血液发送记录 where 发送时间 between [1] And [2] And NVL(接收状态,0)=0"
    Else
        strSQL = "Select Distinct 接收人 as 姓名 From 血液发送记录 where 接收时间 between [1] And [2] And NVL(接收状态,0)<>0"
    End If
    
    If IsPrivs(mstrPrivs, "所有科室") And Val(cboDepart.ItemData(cboDepart.ListIndex)) = -1 Then
        lngDeptID = 0
        strSQLNew = _
        " Select distinct c.id,c.姓名,C.简码" & vbNewLine & _
        " From 人员表 c,(" & strSQL & ") d" & vbNewLine & _
        " Where c.姓名=d.姓名" & strWhere
    Else
        lngDeptID = Val(cboDepart.ItemData(cboDepart.ListIndex))
        strSQLNew = _
        " Select distinct c.id,c.姓名,C.简码" & vbNewLine & _
        " From 部门表 a, 部门人员 b, 人员表 c,(" & strSQL & ") d" & vbNewLine & _
        " Where a.Id = b.部门id And b.人员id = c.Id And a.Id = [3] And c.姓名=d.姓名" & strWhere
    End If
    vPoint = GetControlRect(objControl.hWnd)
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQLNew, 0, "", False, "", "请选择一个" & IIf(mtbcThisIndex = 0, "取血", "接收") & "人员", False, False, True, vPoint.Left, vPoint.Top, objControl.Height, blnCancel, False, False, CDate(Format(dtpSttime.Value, "YYYY-MM-DD HH:mm")), CDate(Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm")), lngDeptID, StrInput & "%")
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Sub
            objControl.Text = rsUser!姓名 & ""
            objControl.Tag = objControl.Text
        End If
    End If
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Sub BloodReceives(frmMain As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional strPrivs As String, Optional lngisModul As Long = 0, Optional int场合 As Integer = 0)
    '功能：血液接收登记入口函数
    Dim lngi As Long
    Dim strSQL As String
    Dim str取血人 As String
    Dim rs取血人 As ADODB.Recordset
    Dim rs人员
    Dim objPane As Pane

    If mblnBloodReceivesIsOpen = True Then GoTo TOSHOW
    
    mlngSys = lngSys
    mlngMoudle = lngMoudle
    mstrPrivs = strPrivs
    mint场合 = int场合
    mblnHavePrivs = False
    Set mfrmMain = frmMain
    
    InitCommandBar '初始化commandbar
'    '初始化DockingPane
    Call DockPannelInit(dkpPeoPle)
    dkpPeoPle.SetCommandBars cbsMain
    Set objPane = dkpPeoPle.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "病人": objPane.Options = PaneNoCaption
    Set objPane = dkpPeoPle.CreatePane(2, 800, 100, DockRightOf, Nothing): objPane.Title = "记录": objPane.Options = PaneNoCaption
    
    '初始化tbcthis
    mtbcThisIndex = 0
    With tbcthis
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ShowIcons = False
        End With
        .InsertItem(0, "待接收血液", Pic5.hWnd, 0).Tag = "待接收血液"
        .InsertItem(1, "已接收血液", pic11.hWnd, 0).Tag = "已接收血液"
        .Item(0).Selected = True
        Call tbcThis_SelectedChanged(tbcthis.Selected)
    End With
TOSHOW:
    mblnBloodReceivesIsOpen = True
    If IsObject(frmMain) Then
        If frmMain Is Nothing Then
            Me.Show lngisModul
        Else
            Me.Show lngisModul, frmMain
        End If
    Else
        gobjComlib.os.ShowChildWindow Me.hWnd, Val(frmMain)
    End If
End Sub

Private Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
End Function


Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '功能：新建ToRs记录集，将RsProm的结构复制到ToRs上
    '参数：RsProm-原记录集，ToRs-新建的记录集
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '初始化rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long, lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim StrSqlSAD As String
    Dim strSQL As String, strSql1 As String, strSql2 As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColor As Long
    Dim strABORH As String
    Dim rsBR As ADODB.Recordset
    Dim blnSelect As Boolean, lngDeptID As Long, strOpter As String
    Dim strCurDate As String, strRows As String
    
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "基础病人查询":
                If mtbcThisIndex = 0 Then
                    strSql1 = "Select Distinct 配发id From 血液发送记录 Where 发送时间 Between [1] And [2] And nvl(接收状态,0)=0"
                    If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
                        strSql1 = strSql1 & " And 取血人=[3]"
                    End If
                Else
                    strSql1 = "Select Distinct 配发id From 血液发送记录 Where 接收时间 Between [1] And [2] And nvl(接收状态,0)<>0"
                    If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
                        strSql1 = strSql1 & " And 接收人=[3]"
                    End If
                End If
                
                If optOccasion(0).Value = True Then '住院病人
                    If marrFilter(mtbcThisIndex).DeptID <> -1 And mtbcThisIndex = 0 Then
                        '某个部门查看，需考虑病人转科的情况，转入前和转入后的科室都可以接收血液
                        strSql1 = "Select Distinct 配发id,发送时间 From 血液发送记录 Where 发送时间 Between [1] And [2] And nvl(接收状态,0)=0"
                        strSQL = _
                            " Select a.病人id || '-' || a.主页ID As Id,a.病人ID, a.主页id, a.住院号 病历号,  Decode(Nvl(a.病人性质, 0), 0, '住', '留')As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别年龄, a.出院科室id As 科室id, d.名称," & vbNewLine & _
                            "       a.入院日期 As 日期, a.出院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, '' As Aborh" & vbNewLine & _
                            " From 病案主页 a, 部门表 d," & vbNewLine & _
                            "     (Select Distinct b.病人id, b.主页id" & vbNewLine & _
                            "       From 血液配血记录 b, (" & strSql1 & ") c" & vbNewLine & _
                            "       Where b.Id = c.配发id And b.主页id Is Not Null And Exists" & vbNewLine & _
                            "        (Select 1" & vbNewLine & _
                            "              From 病人变动记录" & vbNewLine & _
                            "              Where b.病人id = 病人id And b.主页id = 主页id And 科室id = [4] And" & vbNewLine & _
                            "                    Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) >= c.发送时间)) b" & vbNewLine & _
                            " Where a.病人id = b.病人id And a.主页id = b.主页id And a.出院科室id = d.Id(+)"
                    Else
                        If marrFilter(mtbcThisIndex).DeptID <> -1 Then
                            strSql1 = strSql1 & " And 执行科室ID=[4]"
                        End If
                        strSQL = _
                            " Select Distinct a.病人id || '-' || a.主页ID As Id,a.病人ID, a.主页id, a.住院号 病历号,  Decode(Nvl(a.病人性质, 0), 0, '住', '留')As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别年龄, a.出院科室id As 科室id, d.名称," & vbNewLine & _
                            "                a.入院日期 As 日期, a.出院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, '' As Aborh" & vbNewLine & _
                            " From 病案主页 a, 部门表 d, 血液配血记录 b, (" & strSql1 & ") c" & vbNewLine & _
                            " Where a.病人id = b.病人id And a.主页id = b.主页id And a.出院科室id = d.Id(+) And b.Id = c.配发id And b.主页id Is Not Null"
                    End If
                Else '门诊病人
                    If marrFilter(mtbcThisIndex).DeptID <> -1 Then
                        If mtbcThisIndex = 0 Then
                            strSql2 = strSql2 & " And a.执行部门id=[4]"
                        Else
                            strSql1 = strSql1 & " And 执行科室ID=[4]"
                        End If
                    End If
                    
                    strSQL = _
                        " Select Distinct a.病人id || '-' || A.id As Id, A.病人ID,A.ID 主页ID, a.门诊号 病历号, Decode(a.急诊,1,'急',decode(a.复诊,1,'复','普')) As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别年龄, a.执行部门id As 科室id, e.名称 科室名称," & vbNewLine & _
                        "                a.执行时间 As 日期, '' As 床号, a.险类, '' As 类型, 255 As 颜色, a.执行人, '' As Aborh" & vbNewLine & _
                        " From 病人挂号记录 a, 部门表 e, 病人医嘱记录 b, 血液配血记录 c,(" & strSql1 & ") d" & vbNewLine & _
                        " Where a.No = b.挂号单" & strSql2 & " And a.执行部门id = e.Id(+) And b.Id = c.申请id And c.Id = d.配发id And c.主页id Is Null"
                End If
                
                Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), Trim(marrFilter(mtbcThisIndex).Oper), marrFilter(mtbcThisIndex).DeptID)
                Call RsTitelCopy(rsBR, mRsBR(mtbcThisIndex))  '新建记录集mRsBR，其结构复制rsBR
                '下面是对查询到的数据进行处理，将处理后的数据赋值给通过RsTitelCopy函数生成的记录集
                With mRsBR(mtbcThisIndex)
                    If rsBR.RecordCount > 0 Then
                        For lngi = 0 To rsBR.RecordCount - 1
                            .AddNew
                            For lngj = 0 To rsBR.Fields.Count - 1
                                .Fields(lngj).Value = rsBR.Fields(lngj).Value
                                
                                If .Fields(lngj).name = "日期" Then
                                    .Fields(lngj).Value = Format(rsBR.Fields("日期").Value, "YYYY-MM-DD HH:mm")
                                End If
                                
                                strABORH = ""
                                If .Fields(lngj).name = "ABORH" Then '重新给ABORH赋值
                                    Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "ABO")
                                    If rsTmp.BOF = False Then strABORH = rsTmp("信息值").Value
                                    If strABORH = "" Then '门诊病人查询血型
                                        Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "血型")
                                        If rsTmp.BOF = False Then strABORH = rsTmp("信息值").Value
                                    End If
                                    Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "RH")
                                    If rsTmp.BOF = False Then strABORH = strABORH & rsTmp("信息值").Value 'ABO&RH
                                     .Fields("ABORH").Value = strABORH
                                End If
                                
                                If .Fields(lngj).name = "颜色" Then '重新根据类型和险类分配颜色
                                    If Not IsNull(rsBR!险类) And rsBR!类型 & "" = "" Then
                                        '病人颜色
                                        lngColor = &HC0&
                                    Else
                                        lngColor = gobjDatabase.GetPatiColor(Nvl(rsBR!类型))
                                    End If
                                    .Fields("颜色").Value = lngColor
                                End If
                            Next
                            .Update
                            rsBR.MoveNext
                        Next
                        .MoveFirst
                    End If
                End With
                If mblnHavePrivs = True Then
                    UCP.ShowPeople mRsBR(mtbcThisIndex)
                    If marrPreCardID(mtbcThisIndex) <> "" Then Call UCP.SetCardFocus("ID", marrPreCardID(mtbcThisIndex))
                End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "初始表格"
            
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSF1, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("收发id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("病人id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("主页id", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", "选择", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", "图标", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", "状态", True)
                    Call .AppendColumn("床号", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("姓名", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("性别", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("血液名称", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("规格", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("血袋编号", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ABO", 600, flexAlignLeftCenter, flexDTString, , "ABO", True)
                    Call .AppendColumn("Rh", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
                    Call .AppendColumn("接收状态", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("执行部门", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("接收人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("接收时间", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("核收人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("执行科室id", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("取血人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("拒绝原因", 2000, flexAlignLeftCenter, flexDTString, "拒绝原因", "", True)
                    Call .AppendColumn("当前科室ID", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    .AppendRows = False
                    .SysHidden(.ColIndex("收发id")) = True
                    .SysHidden(.ColIndex("病人id")) = True
                    .SysHidden(.ColIndex("主页id")) = True

                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    Call .InitializeEditColumn(.ColIndex("接收状态"), True, vbVsfEditCombox, "正在接收|拒绝接收")
                    Call .InitializeEditColumn(.ColIndex("拒绝原因"), True, vbVsfEditText)
                    
                End With
                
                Set mclsvsf1 = New clsVsf
                With mclsvsf1
                    Call .Initialize(Me.Controls, VSF2, True, True) ', frmPubResource.GetImageList(16)由于没有frmpubresource所以这里将之去掉
                    Call .ClearColumn
                    Call .AppendColumn("收发id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("病人id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("主页id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", "选择", True)
                    Call .AppendColumn("", 300, flexAlignLeftCenter, flexDTString, "图标", "图标", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "状态", "状态", True)
                    Call .AppendColumn("床号", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("姓名", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("性别", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("血液名称", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("规格", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("血袋编号", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ABO", 600, flexAlignLeftCenter, flexDTString, , "ABO", True)
                    Call .AppendColumn("Rh", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
                    Call .AppendColumn("接收状态", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("执行部门", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("接收人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("接收时间", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("核收人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("执行科室id", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("取血人", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("拒绝原因", 2000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("当前科室ID", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    .AppendRows = False
                    .SysHidden(.ColIndex("收发id")) = True
                    .SysHidden(.ColIndex("病人id")) = True
                    .SysHidden(.ColIndex("主页id")) = True
                    
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    Call .InitializeEditColumn(.ColIndex("拒绝原因"), True, vbVsfEditText)
                End With
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "接收数据"
                blnSelect = False
                With VSF1
                    '查询有无选中的数据
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    '无选中数据则提示
                    If blnSelect = False Then
                        MsgBox "请选择要接收的血液数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    '核对数据
                    If frmBloodVerification.ShowCheck(Me, True) = False Then ExecuteCommand = False: Exit Function
                    strOpter = frmBloodVerification.str核收人
                    '保存选中数据
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then  '处于选中状态
                            If marrFilter(mtbcThisIndex).DeptID = -1 Then
                                lngDeptID = Val(.TextMatrix(lngi, .ColIndex("当前科室id")))
                            Else
                                lngDeptID = marrFilter(mtbcThisIndex).DeptID
                            End If
                            If opt1(0).Value = True Then
                                StrSqlSAD = "Zl_血液接收登记_Receive(" & Val(.TextMatrix(lngi, .ColIndex("收发id"))) & ",'" & UserInfo.姓名 & "'," & IIf(.TextMatrix(lngi, .ColIndex("接收状态")) = "正在接收", 1, 2) & ",'" & .TextMatrix(lngi, .ColIndex("拒绝原因")) & "',null,'" & strOpter & "'," & lngDeptID & ")"  '接收人为登陆者
                            Else
                                StrSqlSAD = "Zl_血液接收登记_Receive(" & Val(.TextMatrix(lngi, .ColIndex("收发id"))) & ",'" & UserInfo.姓名 & "'," & IIf(.TextMatrix(lngi, .ColIndex("接收状态")) = "正在接收", 1, 2) & ",'" & .TextMatrix(lngi, .ColIndex("拒绝原因")) & "',To_Date('" & DTP3.Value & "','YYYY-MM-DD hh24:mi'),'" & strOpter & "'," & lngDeptID & ")"   '接收人为登陆者
                            End If
                            Call SQLRecordAdd(rsSAD, StrSqlSAD)
                        End If
                    Next
                    Call SQLRecordExecute(rsSAD)
                End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "回退"
                With VSF2
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    If blnSelect = False Then
                        MsgBox "请选择要回退的血液数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End With
                For lngi = 1 To VSF2.Rows - 1
                    '回退时要考虑当前登陆用户和接收人是否相同，如果不同不允许回退，谁接收，谁才可以回退
                    If Abs(Val(VSF2.TextMatrix(lngi, 3))) = 1 Then '处于选中状态
                        
                        StrSqlSAD = "Zl_血液接收登记_fallback(" & Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("收发id"))) & ")" 'fallback,Unreceive
                        Call SQLRecordAdd(rsSAD, StrSqlSAD)
                    End If
                Next
                Call SQLRecordExecute(rsSAD)
            Case "转接"
                blnSelect = False
                With VSF2
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    If blnSelect = False Then
                        MsgBox "请选择要转接的血液数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If frmBloodVerification.ShowCheck(Me, True) = False Then ExecuteCommand = False: Exit Function
                    strOpter = frmBloodVerification.str核收人
                    If optTransit(0).Value = True Then
                        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:MM")
                    Else
                        strCurDate = Format(dtpTransit.Value, "YYYY-MM-DD HH:MM")
                    End If
                    strRows = ""
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then  '处于选中状态
                            '转接之前要判断转接时间是否大于接收时间，如果小于接收时间，则要提示且不予以接收
                            If strCurDate > Format(VSF2.TextMatrix(lngi, VSF2.ColIndex("接收时间")), "YYYY-MM-DD HH:mm") Then
                                lngDeptID = Val(.TextMatrix(lngi, .ColIndex("当前科室id")))
                                StrSqlSAD = "Zl_血液接收登记_Transfer(" & Val(.TextMatrix(lngi, .ColIndex("收发id"))) & ",3,'" & .TextMatrix(lngi, .ColIndex("接收人")) & "',To_Date('" & strCurDate & "','YYYY-MM-DD hh24:mi'),'" & strOpter & "'," & lngDeptID & ",NULL)"
                                Call SQLRecordAdd(rsSAD, StrSqlSAD)
                            Else
                                strRows = strRows & "," & lngi
                            End If
                        End If
                    Next
                    If strRows <> "" Then
                        If MsgBox("第[" & Mid(strRows, 2) & "]行数据的接收时间[" & strCurDate & "]小于上次接收时间，以上数据本次将不能转接。" & vbCrLf & "请问是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    End If
                    Call SQLRecordExecute(rsSAD)
                End With
            Case "刷新数据"
                '根据当前选中的病人刷新vsf上的数据
                Select Case mtbcThisIndex
                    Case 0 '选项卡1
                        If chk1.Value = Unchecked Then '单选
                            Call ShowVsf(0, , UCP.strReturn)
                        Else '多选
                            Call ShowVsf(0, UCP.GetCheckedData)
                        End If
                    Case 1 '选项卡2
                        If chk1.Value = Unchecked Then '单选
                            Call ShowVsf(1, , UCP.strReturn)
                        Else '多选
                            Call ShowVsf(1, UCP.GetCheckedData)
                        End If
                End Select
                chk2.Value = Unchecked
                chk3.Value = Unchecked
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next

    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    ExecuteCommand = False
End Function

Private Sub chk1_Click()
    '转换提取模式，一个是单个病人提取，一个是多个病人的批量提取，同时也要清空vsf上的数据
    If Me.Visible = True Then
        cmd2.Enabled = chk1.Value
        UCP.CanCheck = chk1.Value
        If mtbcThisIndex = 0 Then
            VSF1.Rows = 2
            VSF1.RowData(1) = 0
            VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        Else
            VSF2.Rows = 2
            VSF2.RowData(1) = 0
            VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        End If
    End If
End Sub

Private Sub cmd1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd2_Click()
    Call ExecuteCommand("刷新数据")
End Sub

Private Sub cmd2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdOper_Click()
    If mblnHavePrivs = False Then Exit Sub
    Call GetSendorReceivePeople(txtOper)
End Sub

Private Sub dkpPeoPle_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = pic1.hWnd
        Case 2
            Item.Handle = pic7.hWnd
    End Select
End Sub

Private Sub dtpEdtime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpSttime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Set gobjFScrollBar = UCP.FScrollBar
    glngBooldPepWinProc = GetWindowLong(UCP.objPicBack.hWnd, GWL_WNDPROC)
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, glngBooldPepWinProc
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHand
    marrPreCardID(0) = ""
    marrPreCardID(1) = ""
    optOccasion(0).Enabled = True
    optOccasion(1).Enabled = True
    If mint场合 = 1 Or mint场合 = 2 Then
        If mint场合 = 1 Then
            optOccasion(1).Value = True
            optOccasion(0).Enabled = False
        Else
            optOccasion(0).Value = True
            optOccasion(1).Enabled = False
        End If
    End If
    mbln转接 = False
    Call InitCondFilter
    DTP3.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    dtpTransit.Value = DTP3.Value
    
    Call ExecuteCommand("初始表格")
'    Set mfrmBloodPeoPle = New frmBloodPeoPle
    '初始化UCP控件
    UCP.UserInit Me, "颜色|ID|1||||255;住院情况|主页ID;床号;姓名;病历号;性别年龄;日期;ABORH;图标", ils16, p血液接收登记
    Call LoadDeptAndCard '初始化部门还有病人信息
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDeptAndCard()
    Dim strSQL As String
    Dim bln门诊 As Boolean, i As Integer
'    Dim rs部门 As New ADODB.Recordset
    On Error GoTo ErrHand
    '部门信息
    bln门诊 = optOccasion(1).Value
    Set mrs部门 = GetDeptList("临床", IIf(bln门诊 = True, 1, 2), IsPrivs(mstrPrivs, "所有科室"))
    If mrs部门.RecordCount <= 0 And IsPrivs(mstrPrivs, "所有科室") = False Then
        MsgBox "您不属于" & IIf(bln门诊 = True, "门诊", "住院") & "部门！", vbInformation, gstrSysName
        mblnHavePrivs = False
        Exit Sub
    End If
    cboDepart.Clear
    '所有部门
    If InStr(";" & mstrPrivs & ";", ";所有科室;") > 0 Then
        cboDepart.AddItem "所有部门"
        cboDepart.ItemData(cboDepart.NewIndex) = -1
        cboDepart.Tag = cboDepart.Text
    End If
    
    For i = 1 To mrs部门.RecordCount
        cboDepart.AddItem mrs部门!编码 & "-" & mrs部门!名称
        cboDepart.ItemData(cboDepart.NewIndex) = mrs部门!id
        '所属缺省
        If IsPrivs(mstrPrivs, "所有科室") = False Then
            If mrs部门!缺省 = 1 Then
                Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, cboDepart.NewIndex)
            End If
        End If
        mrs部门.MoveNext
    Next
    If mrs部门.RecordCount > 0 Then
        mrs部门.MoveFirst
    End If
    
    If cboDepart.ListIndex = -1 And cboDepart.ListCount > 0 Then
        Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, 0)
    End If
    cboDepart.Tag = cboDepart.Text
    '调用usrCardPeople控件
    strSQL = "Select '' 颜色,'' ID,'' 住院情况, '' 主页ID,'' 床号,'' 姓名,'' 病历号,'' 性别年龄,'' 日期,'' ABORH From dual where 1<>1"
    Set mRsBR(0) = gobjDatabase.OpenSQLRecord(strSQL, "取血人或接收人信息")
    Set mRsBR(1) = gobjComlib.Rec.CopyNew(mRsBR(0))
    UCP.ShowPeople mRsBR(0)         '向usrCardPeoPle传入参数
    mblnHavePrivs = True
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitCondFilter()
    '功能：初始化cbotime、dtpSttime、dtpEdtime控件
    Dim intDay As Long
    Dim intStart As Long
    
    mintOutPreTime = -1
    cboTime.Clear
    With cboTime
        .AddItem "今天"
        .ItemData(.NewIndex) = 0
        .AddItem "2天内"
        .ItemData(.NewIndex) = 1
        .AddItem "3天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 6
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    
    intStart = Val(gobjDatabase.GetPara("接收取血时间缺省范围", 100, p血液接收登记, "0"))
    '自定义默认定位到今天
    If InStr(1, ",0,1,2,6,", "," & intStart & ",") = 0 Then
        mstr开始时间 = GetDateTime(0, 1)
        mstr结束时间 = GetDateTime(0, 2)
        Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 0)
    Else
        mstr开始时间 = GetDateTime(intStart, 1)
        mstr结束时间 = GetDateTime(intStart, 2)
        Select Case intStart
            Case 0
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 0)
            Case 1
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 1)
            Case 2
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 2)
            Case 6
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 3)
            Case Else
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 4)
        End Select
    End If
    mintOutPreTime = cboTime.ListIndex
    dtpSttime.Value = Format(mstr开始时间, "YYYY-MM-DD HH:mm")
    dtpEdtime.Value = Format(mstr结束时间, "YYYY-MM-DD HH:mm")
    dtpSttime.Enabled = (cboTime.ItemData(cboTime.ListIndex) = -1): dtpEdtime.Enabled = dtpSttime.Enabled
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpPeoPle, 1, 260, 100, 320, Me.ScaleHeight)
    Call SetPaneRange(dkpPeoPle, 2, 100, 100, Me.ScaleWidth, Me.ScaleHeight)
    dkpPeoPle.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call gobjDatabase.SetPara("接收取血时间缺省范围", DateDiff("d", CDate(mstr开始时间), CDate(mstr结束时间)), 100, p血液接收登记)
    
    mblnBloodReceivesIsOpen = False
    Set mRsBR(0) = Nothing
    Set mRsBR(1) = Nothing
    Set mclsVsf = Nothing
    Set mclsvsf1 = Nothing
    Set mrs部门 = Nothing
'    Set marrFilter(0) = Nothing
'    Set marrFilter(1) = Nothing
End Sub

Private Sub UCP_CardChanged()
    Dim strReturn As String
    strReturn = UCP.strReturn
    If strReturn = "" Then Exit Sub
    If chk1.Value = Unchecked Then  '勾选了批量提取，卡片切换不进行数据刷新
        marrPreCardID(mtbcThisIndex) = Split(strReturn, "'")(1)
        Call ExecuteCommand("刷新数据")
    End If
End Sub

Private Sub opt1_Click(Index As Integer)
    If Me.Visible = True Then
        DTP3.Enabled = opt1(1).Value
    End If
End Sub

Private Sub optOccasion_Click(Index As Integer)
    If Me.Visible = True Then
'        If mtbcThisIndex = 0 Then
'            VSF1.Rows = 2
'            VSF1.RowData(1) = 0
'            VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
'        Else
'            VSF2.Rows = 2
'            VSF2.RowData(1) = 0
'            VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
'        End If
        '在换住院/门诊操作模式时，两个页面都要清除掉，当转换tbcthis控件的页面后，在转换操作模式，会导致页面数据残留
        VSF1.Rows = 2
        VSF1.RowData(1) = 0
        VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        VSF2.Rows = 2
        VSF2.RowData(1) = 0
        VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        Call LoadDeptAndCard
        mbln转接 = False
    End If
End Sub

Private Sub optTransit_Click(Index As Integer)
    If Me.Visible = True Then
        dtpTransit.Enabled = optTransit(1).Value
        If dtpTransit.Enabled = True Then
            dtpTransit.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
        End If
    End If
End Sub

Private Sub pic1_Resize()
    On Error Resume Next
    lbl3.Left = pic1.ScaleLeft + 50
    lbl3.Top = pic1.ScaleTop
    lbl3.Width = pic1.ScaleWidth - 100
    lbl3.Height = 260
    
    Fra1.Left = lbl3.Left
    Fra1.Top = pic1.ScaleTop + lbl3.Height
    Fra1.Width = pic1.ScaleWidth - 100
    Fra1.Height = 2895
    
    UCP.Left = lbl3.Left
    UCP.Top = Fra1.Top + Fra1.Height + 100
    UCP.Width = Fra1.Width
    If pic1.ScaleHeight - Fra1.Top - Fra1.Height - 100 > 0 Then
        UCP.Height = pic1.ScaleHeight - Fra1.Top - Fra1.Height - 100
    End If
    
    'Fra1中控件处理
    cboDepart.Width = Fra1.Width - cboDepart.Left - 120
    cboTime.Width = cboDepart.Width
    dtpSttime.Width = cboDepart.Width
    dtpEdtime.Width = cboDepart.Width
    txtOper.Width = cboDepart.Width
    cmdOper.Left = txtOper.Left + txtOper.Width - cmdOper.Width - 30
    cmd1.Left = cboDepart.Width + cboDepart.Left - cmd1.Width
    cmd2.Left = cmd1.Left
End Sub

Private Sub pic11_Resize()
    On Error Resume Next
    If picTransit.Visible = False Then
        pic4.Move pic11.ScaleLeft, pic11.ScaleTop, pic11.ScaleWidth, pic11.ScaleHeight
    Else
        picTransit.Move pic11.ScaleLeft, pic11.ScaleTop, pic11.ScaleWidth, 500
        pic4.Move pic11.ScaleLeft, picTransit.Top + picTransit.Height, pic11.ScaleWidth, pic11.ScaleHeight - picTransit.Top - picTransit.Height
    End If
End Sub

Private Sub pic3_Resize()
    On Error Resume Next
    VSF1.Move pic3.ScaleLeft + 50, pic3.ScaleTop + 50, pic3.ScaleWidth - 100, pic3.ScaleHeight - 100
    chk3.Move VSF1.Left + 10 + VSF1.ColWidth(3) / 2 - 100, VSF1.Top + 10, VSF1.ColWidth(3) / 2 + 90
End Sub

Private Sub Pic4_Resize()
    On Error Resume Next
    VSF2.Move pic4.ScaleLeft + 50, pic4.ScaleTop + 50, pic4.ScaleWidth - 100, pic4.ScaleHeight - 100
    chk2.Move VSF2.Left + 10 + VSF2.ColWidth(3) / 2 - 100, VSF2.Top + 10, VSF2.ColWidth(3) / 2 + 90
End Sub

Private Sub Pic5_Resize()
    On Error Resume Next
    pic2.Move Pic5.ScaleLeft, Pic5.ScaleTop, Pic5.ScaleWidth, 500
    pic3.Move Pic5.ScaleLeft, pic2.Top + pic2.Height, Pic5.ScaleWidth, Pic5.ScaleHeight - pic2.Top - pic2.Height
End Sub

Private Sub pic7_Resize()
    On Error Resume Next
    tbcthis.Move pic7.ScaleLeft, pic7.ScaleTop, pic7.ScaleWidth, pic7.ScaleHeight
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '在转换选项卡时刷新页面
    Dim i As Integer
    If Item.Tag = "" Then Exit Sub
    mtbcThisIndex = Item.Index
    If mtbcThisIndex = 0 Then
        lbl2.Caption = "取血时间"
        lbl5.Caption = "取 血 人"
    Else
        lbl2.Caption = "接收时间"
        lbl5.Caption = "接 收 人"
    End If
    For i = 0 To cboDepart.ListCount - 1
        If cboDepart.ItemData(i) = marrFilter(mtbcThisIndex).DeptID Then
            Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, i)
            Exit For
        End If
    Next
    cboTime.ListIndex = marrFilter(mtbcThisIndex).TimeIndex
    If cboTime.ListIndex = cboTime.ListCount - 1 Then
        dtpSttime.Value = Format(marrFilter(mtbcThisIndex).BeginTime, "YYYY-MM-DD HH:mm")
        dtpEdtime.Value = Format(marrFilter(mtbcThisIndex).EndTime, "YYYY-MM-DD HH:mm")
    End If
    txtOper.Text = marrFilter(mtbcThisIndex).Oper
    cmd1_Click
    mbln转接 = False '跳转页面后会重置转接状态
    UCP.FindStart = True '跳转页面会初始化查询
    pic11_Resize
End Sub

Private Sub txtOper_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        If txtOper.Tag <> txtOper.Text And Trim(txtOper.Text) <> "" Then
            Call GetSendorReceivePeople(txtOper, txtOper.Text)
        End If
        txtOper.Tag = txtOper.Text
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub VSF1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Trim(VSF1.TextMatrix(Row, VSF1.ColIndex("拒绝原因"))) <> "" And VSF1.TextMatrix(Row, VSF1.ColIndex("接收状态")) = "正在接收" Then
        VSF1.TextMatrix(Row, VSF1.ColIndex("接收状态")) = "拒绝接收"
'    Else
'        VSF1.TextMatrix(Row, VSF1.ColIndex("接收状态")) = "正在接收"
    End If
End Sub

Private Sub VSF1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSF1_AfterScroll(ByVal OldtopRow As Long, ByVal OldLeftCol As Long, ByVal NewtopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol > 3 Then
        chk3.Visible = False
    Else
        chk3.Visible = True
    End If
End Sub

Private Sub VSF1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    chk3.Move VSF1.Left + 10 + VSF1.ColWidth(3) / 2 - 100, VSF1.Top + 10, VSF1.ColWidth(3) / 2 + 90
End Sub

Private Sub VSF1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '不能输入符号"'"
End Sub

Private Sub VSF2_AfterScroll(ByVal OldtopRow As Long, ByVal OldLeftCol As Long, ByVal NewtopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol > 3 Then
        chk2.Visible = False
    Else
        chk2.Visible = True
    End If
End Sub

Private Sub VSF2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    chk2.Move VSF2.Left + 10 + VSF2.ColWidth(3) / 2 - 100, VSF2.Top + 10, VSF2.ColWidth(3) / 2 + 90
End Sub

Private Sub VSF1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSF1.TextMatrix(Row, VSF1.ColIndex("病人id"))) = 0 Then Cancel = True: Exit Sub
    '控制某些行不能编辑
    If Col = VSF1.ColIndex("选择") Then
        Cancel = False
    Else
        Cancel = True
    End If
    If Abs(Val(VSF1.TextMatrix(Row, VSF1.ColIndex("选择")))) = 1 Then
        If Col = VSF1.ColIndex("接收状态") Or Col = VSF1.ColIndex("拒绝原因") Then
            Cancel = False
        End If
'        If Col = VSF1.ColIndex("拒绝原因") And VSF1.TextMatrix(Row, VSF1.ColIndex("接收状态")) = "拒绝接收" Then
'            Cancel = False
'        End If
    End If
End Sub

Private Sub VSF1_DblClick()
    Call mclsVsf.DbClick
End Sub

Private Sub VSF2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsvsf1.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSF2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSF2.TextMatrix(Row, VSF2.ColIndex("病人id"))) = 0 Then Cancel = True: Exit Sub
    If Col = VSF2.ColIndex("选择") Then
        If mbln转接 = False Then
'            If (UserInfo.姓名 = VSF2.TextMatrix(Row, VSF2.ColIndex("接收人")) Or UserInfo.姓名 = VSF2.TextMatrix(Row, VSF2.ColIndex("核收人"))) Then
'                Cancel = False
'            Else
'                Cancel = True
'            End If
        Else
            If Val(VSF2.TextMatrix(Row, VSF2.ColIndex("状态"))) = 1 Or Val(VSF2.TextMatrix(Row, VSF2.ColIndex("状态"))) = 3 Then
                Cancel = True
            Else
                Cancel = False
            End If
        End If
    Else
        Cancel = True
    End If
End Sub

Private Function IsClick() As Boolean
    '功能：判断在转接模式下是否有血液登记记录被选中
    Dim lngi As Long
    IsClick = False
    For lngi = 1 To VSF2.Rows - 1
        If Abs(Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("选择")))) = 1 Then
            IsClick = True
            Exit For
        End If
    Next
End Function

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '功能：将记录集RsProm的结构还有数据都复制给ToRs
    '参数：RsProm-要赋值的记录集，ToRs-目标记录集
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '以前没有对rsbr的数据做判断会报错
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub
