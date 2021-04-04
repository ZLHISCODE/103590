VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmReady 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "接单"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12795
   Icon            =   "frmReady.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   2  'Point
   ScaleWidth      =   639.75
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   195
      Picture         =   "frmReady.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8075
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   180
      Picture         =   "frmReady.frx":0156
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   83
      Top             =   7755
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   180
      Picture         =   "frmReady.frx":06E0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   81
      Top             =   7420
      Width           =   240
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "全选(&H)"
      Height          =   300
      Index           =   1
      Left            =   10440
      TabIndex        =   88
      Top             =   2320
      Width           =   990
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "全消(&N)"
      Height          =   300
      Index           =   0
      Left            =   11520
      TabIndex        =   87
      Top             =   2320
      Width           =   990
   End
   Begin VB.Frame Fra 
      Caption         =   "查找条件"
      Height          =   840
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   12495
      Begin zlIDKind.IDKindNew idkSelect 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         ShowSortName    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
         SaveRegType     =   4
      End
      Begin VB.CommandButton cmdReadCard 
         Caption         =   "读卡"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3450
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   1455
         TabIndex        =   1
         Top             =   300
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   8970
         TabIndex        =   6
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   6675
         TabIndex        =   5
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   345
         Left            =   11085
         TabIndex        =   7
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送时间                         ～"
         Height          =   180
         Left            =   5790
         TabIndex        =   4
         Top             =   345
         Width           =   3150
      End
   End
   Begin MSComctlLib.ImageList imgPic 
      Left            =   9960
      Top             =   1920
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
            Picture         =   "frmReady.frx":0C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":1204
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":179E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":1D38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTransfusion 
      Height          =   3780
      Left            =   180
      TabIndex        =   80
      Top             =   3510
      Width           =   12375
      _cx             =   21828
      _cy             =   6667
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReady.frx":1E92
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
   Begin VB.Frame fraPrint 
      Caption         =   "请选择要打印的单据"
      Height          =   1000
      Left            =   2400
      TabIndex        =   60
      Top             =   7440
      Width           =   7170
      Begin VB.CheckBox chkWristband 
         Caption         =   "输液腕带"
         Height          =   250
         Left            =   2520
         TabIndex        =   43
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLabel 
         Caption         =   "输液瓶签"
         Height          =   250
         Left            =   1320
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkView 
         Caption         =   "预览"
         Height          =   250
         Left            =   6240
         TabIndex        =   44
         Top             =   600
         Width           =   700
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "治疗单"
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "皮试单"
         Height          =   250
         Index           =   3
         Left            =   2520
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "注射单"
         Height          =   250
         Index           =   2
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "输液单"
         Height          =   250
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   11545
      TabIndex        =   47
      Top             =   7960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   10320
      TabIndex        =   46
      Top             =   7960
      Width           =   1100
   End
   Begin VB.ComboBox cboOperator 
      Height          =   300
      Index           =   0
      Left            =   10725
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   7440
      Width           =   1920
   End
   Begin VB.Frame frmBaseInfo 
      Caption         =   "基本信息"
      Height          =   1365
      Left            =   135
      TabIndex        =   50
      Top             =   930
      Width           =   12525
      Begin VB.ComboBox cboSeating 
         Height          =   300
         Left            =   9525
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   2670
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   9
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   49
         Top             =   915
         Width           =   8530
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   8
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   48
         Top             =   915
         Width           =   1680
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   7
         Left            =   9525
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   45
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   6
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   37
         Top             =   555
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   5
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   14
         Top             =   585
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   4
         Left            =   945
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   13
         Top             =   585
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   3
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   12
         Top             =   210
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   1
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   11
         Top             =   255
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   9
         Top             =   255
         Width           =   1700
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "诊断"
         Height          =   240
         Index           =   9
         Left            =   3105
         TabIndex        =   61
         Top             =   930
         Width           =   525
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "医生"
         Height          =   240
         Index           =   8
         Left            =   210
         TabIndex        =   59
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "病人科室"
         Height          =   240
         Index           =   7
         Left            =   8730
         TabIndex        =   58
         Top             =   615
         Width           =   780
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "年龄"
         Height          =   240
         Index           =   6
         Left            =   5895
         TabIndex        =   57
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "性别"
         Height          =   240
         Index           =   5
         Left            =   3000
         TabIndex        =   56
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "姓名"
         Height          =   240
         Index           =   4
         Left            =   300
         TabIndex        =   55
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "就诊时间"
         Height          =   240
         Index           =   3
         Left            =   5745
         TabIndex        =   54
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "座位号"
         Height          =   240
         Index           =   2
         Left            =   8880
         TabIndex        =   53
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "顺序号"
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   52
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "挂号单号"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   51
         Top             =   300
         Width           =   750
      End
   End
   Begin TabDlg.SSTab stabType 
      Height          =   5010
      Left            =   90
      TabIndex        =   8
      Top             =   2370
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "(&0)治疗"
      TabPicture(0)   =   "frmReady.frx":1F2D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblForecastTime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblOperator(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl治疗"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpForecastTime(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboOperator(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt摘要(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "(&1)输液"
      TabPicture(1)   =   "frmReady.frx":1F49
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblOperator(2)"
      Tab(1).Control(1)=   "lblTransfusion(1)"
      Tab(1).Control(2)=   "lblTransfusion(2)"
      Tab(1).Control(3)=   "lblTransfusion(0)"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "lblOperator(3)"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "dtpForecastTime(1)"
      Tab(1).Control(8)=   "cboOperator(2)"
      Tab(1).Control(9)=   "txtTransfusion(0)"
      Tab(1).Control(10)=   "txtTransfusion(1)"
      Tab(1).Control(11)=   "txtTransfusion(2)"
      Tab(1).Control(12)=   "cboOperator(3)"
      Tab(1).Control(13)=   "txt摘要(1)"
      Tab(1).Control(14)=   "chk输液提醒"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "(&2)注射"
      TabPicture(2)   =   "frmReady.frx":1F65
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "lblOperator(4)"
      Tab(2).Control(2)=   "lblOperator(5)"
      Tab(2).Control(3)=   "Label11"
      Tab(2).Control(4)=   "dtpForecastTime(2)"
      Tab(2).Control(5)=   "cboOperator(4)"
      Tab(2).Control(6)=   "cboOperator(5)"
      Tab(2).Control(7)=   "txt摘要(2)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "(&3)皮试"
      TabPicture(3)   =   "frmReady.frx":1F81
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblScratchTest(0)"
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(2)=   "lblOperator(6)"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "dtpForecastTime(3)"
      Tab(3).Control(5)=   "txtScratchTest"
      Tab(3).Control(6)=   "chk皮试提醒"
      Tab(3).Control(7)=   "cboOperator(6)"
      Tab(3).Control(8)=   "txt摘要(3)"
      Tab(3).ControlCount=   9
      Begin VB.TextBox txt摘要 
         Height          =   300
         Index           =   0
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   17
         Top             =   765
         Width           =   9735
      End
      Begin VB.CheckBox chk输液提醒 
         Caption         =   "到时提醒"
         Height          =   330
         Left            =   -66405
         TabIndex        =   20
         Top             =   405
         Width           =   1050
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Index           =   2
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   29
         Top             =   765
         Width           =   9735
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   5
         Left            =   -71070
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   420
         Width           =   1830
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Index           =   3
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   34
         Top             =   765
         Width           =   9735
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Index           =   1
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   22
         Top             =   765
         Width           =   7260
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   1
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   4
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   6
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   3
         Left            =   -65850
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   765
         Width           =   1830
      End
      Begin VB.TextBox txtTransfusion 
         Height          =   300
         Index           =   2
         Left            =   -67365
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   25
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtTransfusion 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   -69540
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   420
         Width           =   990
      End
      Begin VB.TextBox txtTransfusion 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   -71265
         MaxLength       =   2
         TabIndex        =   19
         ToolTipText     =   "滴系数的取值为10，15，20"
         Top             =   420
         Width           =   495
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   2
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.CheckBox chk皮试提醒 
         Caption         =   "到时提醒"
         Height          =   330
         Left            =   -69690
         TabIndex        =   32
         Top             =   405
         Width           =   1050
      End
      Begin VB.TextBox txtScratchTest 
         Height          =   300
         Left            =   -70725
         MaxLength       =   5
         TabIndex        =   31
         Top             =   420
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   15
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   1
         Left            =   -73770
         TabIndex        =   18
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   2
         Left            =   -73770
         TabIndex        =   26
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   3
         Left            =   -73770
         TabIndex        =   30
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin VB.Label lbl治疗 
         Caption         =   "执行摘要"
         Height          =   255
         Left            =   450
         TabIndex        =   63
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "执行摘要"
         Height          =   255
         Left            =   -74550
         TabIndex        =   79
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblOperator 
         Caption         =   "配药人"
         Height          =   240
         Index           =   5
         Left            =   -71670
         TabIndex        =   78
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label9 
         Caption         =   "执行摘要"
         Height          =   255
         Left            =   -74550
         TabIndex        =   77
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label8 
         Caption         =   "执行摘要"
         Height          =   255
         Left            =   -74550
         TabIndex        =   76
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblOperator 
         Caption         =   "执行人"
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   75
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "执行人"
         Height          =   240
         Index           =   4
         Left            =   -74400
         TabIndex        =   74
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "执行人"
         Height          =   240
         Index           =   6
         Left            =   -74400
         TabIndex        =   73
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "配药人"
         Height          =   240
         Index           =   3
         Left            =   -66450
         TabIndex        =   72
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "预计开始时间"
         Height          =   255
         Left            =   -74895
         TabIndex        =   71
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "预计开始时间"
         Height          =   255
         Left            =   -74895
         TabIndex        =   70
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "预计开始时间"
         Height          =   255
         Left            =   -74895
         TabIndex        =   69
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblForecastTime 
         Caption         =   "预计开始时间"
         Height          =   255
         Left            =   105
         TabIndex        =   68
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "滴系数"
         Height          =   255
         Index           =   0
         Left            =   -71880
         TabIndex        =   67
         Top             =   495
         Width           =   555
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "预计时间(分)"
         Height          =   255
         Index           =   2
         Left            =   -68475
         TabIndex        =   66
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "液体总量(ml)"
         Height          =   255
         Index           =   1
         Left            =   -70665
         TabIndex        =   65
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblOperator 
         Caption         =   "执行人"
         Height          =   240
         Index           =   2
         Left            =   -74400
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblScratchTest 
         Caption         =   "预计时间(分)"
         Height          =   255
         Index           =   0
         Left            =   -71835
         TabIndex        =   62
         Top             =   495
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已完成"
      Height          =   180
      Left            =   435
      TabIndex        =   86
      Top             =   8080
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "鼠标右键拒绝"
      Height          =   240
      Left            =   435
      TabIndex        =   84
      Top             =   7770
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "鼠标左键接单"
      Height          =   240
      Left            =   435
      TabIndex        =   82
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Label lblOperator 
      Caption         =   "接单人"
      Height          =   240
      Index           =   0
      Left            =   10050
      TabIndex        =   35
      Top             =   7500
      Width           =   630
   End
   Begin VB.Menu popMenu 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuGNo 
         Caption         =   "挂号单"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuJZK 
         Caption         =   "就诊卡"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMNo 
         Caption         =   "门诊号"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuName 
         Caption         =   "姓　名"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSFZ 
         Caption         =   "身份证"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuICK 
         Caption         =   "ＩＣ卡"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCardSquare 
         Caption         =   "一卡通"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmReady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Base
    挂号单号 = 0
    顺序号 = 1
    座位号 = 2
    就诊时间 = 3
    姓名 = 4
    性别 = 5
    年龄 = 6
    病人科室 = 7
    医生 = 8
    诊断 = 9
End Enum

Private Enum vsCol
    col_选择 = 0
    col_执行顺序 = 1
    col_上次顺序 = 2
    col_序号 = 3
    col_医嘱内容 = 4
    col_剂量 = 5
    col_单位 = 6
    col_项目金额 = 7
    col_执行频率 = 8
    col_用法 = 9
    col_滴速 = 10
    col_容量 = 11
    col_时间 = 12
    col_剩余次数 = 13
    col_收费金额 = 14
    col_医生嘱托 = 15
    col_BillKey = 16
    col_groupkey = 17
    col_执行计费状态 = 18
    col_明细计费状态 = 19
End Enum

Private Enum Trans
    滴系数 = 0
    液体总量 = 1
    预计时间 = 2
End Enum

Private Enum cbo
    接单人 = 0
    治疗执行人 = 1
    输液执行人 = 2
    输液配药人 = 3
    注射执行人 = 4
    注射配药人 = 5
    皮试执行人 = 6
End Enum
Private Enum sType
    治疗 = 0
    输液 = 1
    注射 = 2
    皮试 = 3
End Enum

Private mPatients As cPatients   '病人记录集
Private mPatient As cPatient     '病人

Private mSeatings As Seatings    '座位记录集
Private mOutNurses As OutNurses  '护士记录集
Private mGps治疗 As Groups
Private mGps输液 As Groups
Private mGps注射 As Groups
Private mGps皮试 As Groups

Private mlng部门ID As Long
Private mstr执行科室 As String
Private mstr座位 As String      '在主界面中选中的座位

Private mDateBegin As Date
Private mdateEnd As Date
Private mblnHaveData As Boolean '是否有数据
Private mblnOk As Boolean

Private mstrPrivs As String                 '权限
Private mobjSquareCard  As Object           '一卡通部件 add by 2011-12-23

'Private mintFindType As Integer             '查找类型 0-就诊卡,1-门诊号,2-挂号单,3-姓名,4-身份证,5-IC卡
'Private mstrIDCard As String                '最近自动刷出来的身份证号
'Private WithEvents mobjIDCard As clsIDCard  '身份证对象
Private mobjICCard As Object                'IC卡对象
Private mblnLoad As Boolean                  '是否第一次启动窗体
Private mblnLiquid As Boolean               '是否有配液流程
Private mstrKeyType As String               '接单查询类型
Private marrKey As Variant
Private mblnImmediatePuncture As Boolean    '接单后直接进入穿刺状态
Private mblnActivate As Boolean             '是否已经执行Activate事件
Private mstrSquareCards As String           '一卡通信息
Private mintLabelState As Integer           '上次输液瓶签的状态
Private mintWristband As Integer            '上次输液腕带标签的状态
Private mblnReadCard As Boolean
Private mbytType As Byte                    '接单方式
Private mptiInfo As PatiIdentify            '病人信息控件

'接单方式
'格式说明：
Private Const MSTR_MODE As String = "挂|挂号单|0;就|就诊卡|1;门|门诊号|0;姓|姓名|0;身|身份证号|0;IC|IC卡|1"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChoose_Click(Index As Integer)
    If vsTransfusion.Rows <= 1 Then Exit Sub
    
    Dim i As Integer, intRow As Integer
    
    With vsTransfusion
        intRow = .Row
        .Redraw = False
        For i = 1 To .Rows - 1
            .Row = i
            If .RowData(i) = 0 Or .RowData(i) = 3 Then
                If Index = 1 Then
                    '通过执行顺序判断是否选择
                    If Val(.TextMatrix(i, col_执行顺序)) <= 0 Then
                        Call vsTransfusion_KeyDown(vbKeySpace, 0)
                    End If
                Else
                    If Val(.TextMatrix(i, col_执行顺序)) > 0 Then
                        Call vsTransfusion_KeyDown(vbKeySpace, 0)
                    End If
                End If
            End If
        Next
        .Redraw = True
        .Row = intRow
    End With
    
    vsTransfusion.SetFocus
End Sub

Private Sub cmdOk_Click()
    Dim strStat As String      '准备改为何种状态
    Dim blnNoCall As Boolean
    Dim strSeqNo As String, strErr As String
    Dim blnUpdate As Boolean
    
    '安排座位
    If cboSeating.List(cboSeating.ListIndex) <> "<无座>" And (mPatient.座位号 = "" Or mPatient.座位号 = "无") Then
        If mSeatings.SetSeating(mPatient.病人ID, mPatient.挂号单, cboSeating.List(cboSeating.ListIndex)) Then
            cboSeating.RemoveItem cboSeating.ListIndex '请除被占用座位
        Else
            Exit Sub
        End If
    End If
    
    '填排队记录
'    If txtBase(Base.顺序号).Tag <> "" Then
'        '需要写入
'        Call mPatient.AddQueue(mlng部门ID)
'    End If

    On Error GoTo errHandle
    
    '填写执行记录
    If Not mGps治疗 Is Nothing Then
        If mGps治疗.选择组数 > 0 Then
            
            mGps治疗.p执行摘要 = txt摘要(0)
            mGps治疗.p本次执行时间 = dtpForecastTime(0).Value
            mGps治疗.p接单人 = cboOperator(接单人)
            mGps治疗.SelectGroupThingNew mlng部门ID, chkPrint(0).Value = 1, 0, Me, chkView.Value = 1
            strStat = "7-执行中" '治疗类的改为7－执行中
        End If
    End If
    
    If Not mGps皮试 Is Nothing Then
        If mGps皮试.选择组数 > 0 Then
            mGps皮试.p执行摘要 = txt摘要(3)
            mGps皮试.p本次执行时间 = dtpForecastTime(3).Value
            mGps皮试.p接单人 = cboOperator(接单人)
            mGps皮试.p耗时 = txtTransfusion(预计时间)
            If chk皮试提醒.Value = 1 Then
                If mGps皮试.p耗时 <= 0 Then mGps皮试.p耗时 = 5
                mGps皮试.p提醒 = Val(zldatabase.GetPara("皮试提醒提前时间", glngSys, 1264))
                If mGps皮试.p提醒 < 0 Or mGps皮试.p提醒 > 60 Then mGps皮试.p提醒 = 0
                mGps皮试.p提醒 = mGps皮试.p提醒
            Else
                mGps皮试.p提醒 = -1
            End If
            mGps皮试.SelectGroupThingNew mlng部门ID, chkPrint(3).Value = 1, 3, Me, chkView.Value = 1
            strStat = "7-执行中" '皮试类的改为7－执行中
        End If
    End If
    
    
    If Not mGps注射 Is Nothing Then
        If mGps注射.选择组数 > 0 Then
            mGps注射.p执行摘要 = txt摘要(2)
            mGps注射.p本次执行时间 = dtpForecastTime(2).Value
            mGps注射.p接单人 = cboOperator(接单人)
            mGps注射.SelectGroupThingNew mlng部门ID, chkPrint(2).Value = 1, 2, Me, chkView.Value = 1
            
            strStat = "7-执行中" '注射类的改为7－执行中
        End If
    End If
    
    If Not mGps输液 Is Nothing Then
        If mGps输液.选择组数 > 0 Then
            
            mGps输液.p执行摘要 = txt摘要(1)
            mGps输液.p本次执行时间 = dtpForecastTime(1).Value
            mGps输液.p滴系数 = txtTransfusion(滴系数)
            
            If Not mblnLiquid Then
                If mblnImmediatePuncture Then
                    '直接进入穿刺状态
                    strStat = "7-执行中"
                Else
                    blnNoCall = CurDayNoCall(mlng部门ID, mPatients, mPatient)
                    If Not blnNoCall Then
                        strStat = "1-待配液"  '输液类的改为 5－待穿刺
                    Else
                        strStat = "7-执行中"  '输液类的改为 7－执行中
                    End If
                End If
                mGps输液.p配药人 = cboOperator(输液配药人)
            Else
                strStat = "1-待配液"  '有配液流程，进1-待配液  不填配药人
            End If
            
            mGps输液.p接单人 = cboOperator(接单人)
            mGps输液.p耗时 = txtTransfusion(预计时间)
            If chk输液提醒.Value = 1 Then
                If mGps输液.p耗时 <= 0 Then mGps输液.p耗时 = 5
                mGps输液.p提醒 = Val(zldatabase.GetPara("输液提醒提前时间", glngSys, 1264))
                If mGps输液.p提醒 < 0 Or mGps输液.p提醒 > 60 Then mGps输液.p提醒 = 3
                mGps输液.p提醒 = mGps输液.p提醒
            Else
                mGps输液.p提醒 = -1
            End If
            mGps输液.SelectGroupThingNew mlng部门ID, chkPrint(1).Value = 1, 1, Me, chkView.Value = 1, chkLabel.Value = 1, chkWristband.Value = 1
    
        End If
    End If
    
    On Error GoTo 0

    '当天接过输液的单，就不调 状态了   2014-08-21 取消
    '应为接过单，但没有执行结束的病人，不调整状态。执行结束的病人走正常状态的流程

    blnUpdate = Not CurDayHaveItem(mPatient, mlng部门ID)
    
    'blnUpdate=True：当前没有接过单；排队状态=4：结束；=3：退号；=2：弃号
    If blnUpdate Or Val(mPatient.排队状态) = 4 Or Val(mPatient.排队状态) = 3 Or Val(mPatient.排队状态) = 2 Then
        mPatient.UpdateState strStat, mlng部门ID, False
        
        SaveOperLog mlng部门ID, mPatient, QUEUE, "接单后更改队列状态为" & strStat
        
        '有呼叫流程的，或无“配液”流程均分配穿刺台
        If strStat = "1-待配液" Then
            Call AllocationDesks(mlng部门ID, mPatient, strSeqNo, strErr)
        End If
        If Not mblnLiquid And strStat = "1-待配液" Then
            '无“配液”流程时自动执行配液操作，更改状态
            strStat = Liquid(mlng部门ID, mPatient.Key, mPatients, strErr)
            If Trim(strStat) = "" Then
                strStat = "5-待穿刺"
            End If
            mPatient.UpdateState strStat, mlng部门ID, False
            SaveOperLog mlng部门ID, mPatient, QUEUE, "接单后更改队列状态为" & strStat
            If strStat = "5-待穿刺" Then
                Call QueueCall("输液类", mlng部门ID, mPatient)
            End If
        End If
    End If
   
    If mbytType = Val("1-自动接单") Then
        Unload Me
    Else
        '刷新数据
        Call initObject
        mblnOk = True
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        Call SaveErrLog
    End If
End Sub

Private Sub cmdReadCard_Click()
    If Not mobjICCard Is Nothing Then
        txtNo.Text = mobjICCard.Read_Card(Me)
        If txtNo.Text <> "" Then Call cmdRefresh_Click
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dim strFindType As String, strFindTxt As String
    Dim strFindNo As String, blnFind As Boolean
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single, iRow As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rsVariable As ADODB.Recordset, strSQL As String
    Dim objSeating As Seating, i As Integer
    Dim strInfo As String, strTmp As String
    Dim intFrom As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If Trim(txtNo.Text) = "" Then
        MsgBox "请填写“" & idkSelect.GetCurCard.名称 & "”！", vbInformation, gstrSysName
        txtNo.SetFocus
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '更新各页面的“预计开始时间”
    For i = dtpForecastTime.LBound To dtpForecastTime.UBound
        dtpForecastTime(i).Value = zldatabase.Currentdate
    Next
    
    'If mDateBegin <> dtpBegin.Value Or mdateEnd <> DtpEnd.Value Then
    Me.cmdRefresh.Enabled = False
    Set mPatient = Nothing
    Call initObject '初始化数据集
    Call RefPatiData
    cmdOk.Enabled = False
    
    '--初始化选择器要用的记录集
'    SaveLog "1-初始化选择器要用的记录集"
'    Set rsTmp = New ADODB.Recordset
'    With rsTmp
'        .Fields.Append "ID", adVarChar, 20
'        .Fields.Append "Key", adVarChar, 40
'        .Fields.Append "挂号单", adVarChar, 20
'        .Fields.Append "挂号时间", adVarChar, 20
'        .Fields.Append "开嘱时间", adVarChar, 20
'        .Fields.Append "病人科室", adVarChar, 100
'        .CursorLocation = adUseClient
'        .LockType = adLockOptimistic
'        .CursorType = adOpenStatic
'        .Open
'    End With

    mDateBegin = dtpBegin.Value
    mdateEnd = dtpEnd.Value
    
    '获取数据
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "2-获取参数"
    strFindType = idkSelect.GetCurCard.名称
    If strFindType = "挂号单" Then txtNo.Text = GetFullNO(txtNo.Text, 12)   '12－挂号收据号
    strFindTxt = Trim(txtNo)
    strFindNo = ""
    
'    If strFindType = "挂号单↓" Then
'        '按挂号单查找，可以确定时间,其他方式需要手工指定时间段
'        strSQL = "Select 执行时间 From 病人挂号记录 Where No=[1]"
'        Set rsVariable = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFindTxt)
'        If Not rsVariable.EOF Then
'            mDateBegin = Format(rsVariable!执行时间, "yyyy-MM-dd 00:00:00")
'            mdateEnd = Format(mDateBegin, "yyyy-MM-dd 23:59:59")
'        End If
'    End If
    
'    If InStr(";就诊卡;门诊号;单据号;姓名;身份证;IC卡;ＩＣ卡;", ";" & Replace(strFindType, "↓", "") & ";") = 0 Then

    'strInfo格式：卡类别(1-6固定卡；7以上为一卡通卡)|卡号|一卡通类别ID
    '卡类别
    Select Case strFindType
        Case "就诊卡"
            strInfo = "1"
        Case "门诊号"
            strInfo = "2"
        Case "挂号单", ""
            strInfo = "3"
        Case "姓名"
            strInfo = "4"
        Case "身份证号", "二代身份证"
            strInfo = "5"
        Case Else
            '一卡通
            strInfo = "6"
    End Select
    '卡号
    strInfo = strInfo & "|" & strFindTxt
    '一卡通类别ID
    If Val(strInfo) >= 6 Or Val(strInfo) = 1 Then
        strTmp = GetSquareCardInfo(mstrSquareCards, strFindType, enuCardProperty.卡类别ID)
        strInfo = strInfo & "|" & strTmp
    Else
        strInfo = strInfo & "|"
    End If
    
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "3-获取数据开始"
    Call mPatients.FetchPatients(mlng部门ID, mDateBegin, mdateEnd, True, strInfo, , , mobjSquareCard)
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "4-获取数据结束"
    
    For Each mPatient In mPatients
        blnFind = False
        If strFindType = "挂号单" Then
            If mPatient.挂号单 = strFindTxt Then blnFind = True
        ElseIf strFindType = "门诊号" Then
            If mPatient.门诊号 = strFindTxt Then blnFind = True
        ElseIf strFindType = "姓名" Then
            If mPatient.姓名 = strFindTxt Then blnFind = True
        ElseIf strFindType = "身份证号" Or strFindType = "二代身份证" Then
            If mPatient.身份证号 = strFindTxt Then blnFind = True
        Else
            '一卡通
            blnFind = True
        End If
        
        If blnFind Then
            LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "5-找到对应的挂号单"
            '获取未执行完成，并且未过执行终止时期的挂号单
            If ExecutionComplete(mlng部门ID, mPatient) Then
                strFindNo = strFindNo & "," & mPatient.Key
'                rsTmp.AddNew
'                iRow = iRow + 1
'                rsTmp.Fields("ID").Value = iRow
'                rsTmp.Fields("Key").Value = mPatient.Key
'                rsTmp.Fields("挂号单").Value = mPatient.挂号单
'                rsTmp.Fields("挂号时间").Value = Format(mPatient.挂号时间, "yyyy-MM-dd hh:mm")
'                rsTmp.Fields("开嘱时间").Value = Format(mPatient.挂号时间, "yyyy-MM-dd hh:mm")
'                rsTmp.Fields("病人科室").Value = GetClinicDept(mPatient.病人ID, mPatient.Key)
'                rsTmp.Update
            End If
            intFrom = mPatient.病人来源
        End If
    Next
    
    If strFindNo = "" Then
        LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "5-未找到对应的挂号单"
        Set mPatient = Nothing
        MsgBox "在本科室未找到此人！", vbInformation, gstrSysName
        GoTo hExit
    Else
        strFindNo = Mid(strFindNo, 2)
        
        '门诊病人，门诊留观病人不存在多张挂号单
        If InStr(strFindNo, ",") > 0 And intFrom <> 1 Then
'            '有两个以上的记录,要选择一个挂号单
'            SaveLog "6-有两个以上的记录，要选择一个挂号单"
'            Call ClientToScreen(txtNo.hwnd, objPoint)
'            sglX = objPoint.X * 15 - 30
'            sglY = objPoint.Y * 15 + 300
'
'            SaveLog "7-显示选择窗体"
'            If intFrom = 1 Then
'                strTmp = "id,0,0,0;KEY,0,0,0;开嘱时间,1500,0,0;病人科室,2200,0,0"
'            Else
'                strTmp = "id,0,0,0;KEY,0,0,0;挂号单,1200,0,0;挂号时间,1500,0,0;病人科室,2200,0,0"
'            End If
'            If frmSelect.ShowSelect(Me, rsTmp, strTmp, sglX, sglY, 5000, 3000, Me.Name & "\选择", "请选") Then
'                strFindNo = Trim("" & rsTmp!Key)
'                SaveLog "8-显示选择了" & strFindNo
'            Else
'                SaveLog "8-未选择"
'                GoTo hExit
'            End If

            '有两个以上的记录,要选择一个挂号单
            vRect = zlControl.GetControlRect(txtNo.hwnd)
            
            '不能用Null，因此选择器发现第一行为Null的列会屏蔽列显示
            strSQL = "Select a.Id, Null 上级id, 0 末级, a.No 名称, '' 开嘱科室, '' 开嘱医生, '' 用法, '' 医嘱内容, 0 剂量, " & _
                     "    '' 单位, '' 执行频率, 0 医嘱id, 0 剩余数次 " & vbNewLine & _
                     "From 病人挂号记录 A, Table(f_Str2list([1], ',')) B " & vbNewLine & _
                     "Where a.No = b.Column_Value " & vbNewLine & _
                     "Union All " & vbNewLine & _
                     "Select a.*, b.剩余数次 " & vbNewLine & _
                     "From (Select Rownum * -1 ID, c.Id 上级id, 1 末级, a.挂号单, e.名称 开嘱科室, a.开嘱医生, b.医嘱内容 用法, " & _
                     "          a.医嘱内容, a.单次用量, d.计算单位, a.执行频次, a.相关id " & vbNewLine & _
                     "      From 病人医嘱记录 A, 病人医嘱记录 B, 病人挂号记录 C, 诊疗项目目录 D, 部门表 E, Table(f_Str2list([1], ',')) F " & vbNewLine & _
                     "      Where a.相关id = b.Id And a.挂号单 = c.No And a.诊疗项目id = d.Id And a.开嘱科室id = e.Id " & vbNewLine & _
                     "          And a.挂号单 = f.Column_Value And a.诊疗类别 In ('5', '6', '7')) A," & vbNewLine & _
                     "     (Select a.Id, Nvl(Avg(b.发送数次), 0) - Nvl(Sum(c.本次数次), 0) 剩余数次 " & vbNewLine & _
                     "      From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱执行 C, Table(f_Str2list([1], ',')) D " & vbNewLine & _
                     "      Where a.挂号单 = d.Column_Value And a.Id = b.医嘱id And b.医嘱id = c.医嘱id(+) And b.发送号 = c.发送号(+) " & _
                     "          And a.诊疗类别 = 'E' " & vbNewLine & _
                     "      Group By ID) B " & vbNewLine & _
                     "Where a.相关id = b.Id "
            Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 2, "选择挂号单", False, False, "", False, False, True, _
                                                    vRect.Left, vRect.Bottom, 0, blnCancel, False, False, _
                                                    strFindNo)
            If blnCancel Then
                LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "取消选择挂号单"
                GoTo hExit
            End If
            If rsTmp.EOF Then
                LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "未查询出挂号单数据"
                GoTo hExit
            End If
            
            strFindNo = rsTmp!名称
                        
        End If
    End If
    
    Set mPatient = mPatients(strFindNo)
    
    '-- 初始化座位
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "9-初始化座位"
    cboSeating.Clear
    If mPatient.座位号 = "" Or mPatient.座位号 = "无" Then
        cboSeating.AddItem "<无座>"
        
        For Each objSeating In mSeatings
            If objSeating.病人ID = 0 Then
                cboSeating.AddItem objSeating.类别 & "-" & objSeating.编号
            End If
        Next
        mstr座位 = Replace(mstr座位, "_", "-")
        If cboSeating.ListCount > 0 Then
            For i = 0 To cboSeating.ListCount - 1
                If mstr座位 = "" Then Exit For
                If cboSeating.List(i) = mstr座位 Then
                    Exit For
                End If
            Next
            If i < cboSeating.ListCount Then
                cboSeating.ListIndex = i
            Else
                cboSeating.ListIndex = 0
            End If
        End If
    Else
        cboSeating.AddItem mPatient.座位号
        cboSeating.ListIndex = 0
        cboSeating.Enabled = False
    End If
    If cboSeating.Enabled Then
        cboSeating.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0
    End If
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "10-刷新显示"
    Call InceptBill '刷新显示
    
hExit:
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "9-退出"
    Me.cmdRefresh.Enabled = True
    Exit Sub
    
hErr:
    LogWrite "输液接单的调试日志", "" & glngModul, "cmdRefresh_Click", "单击刷新，第" & CStr(Erl()) & "行，" & Err.Description
    Exit Sub
    
errSQL:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    If mblnActivate Then Exit Sub
    
    If mblnLoad Then
        If mbytType = Val("1-自动接单") Then
            idkSelect.IDKind = mptiInfo.IDKindIDX
            txtNo.Text = mptiInfo.Text
        Else
            For i = 1 To idkSelect.ListCount
                If idkSelect.Cards(i).名称 = mstrKeyType Then
                    idkSelect.IDKind = i
                    Exit For
                End If
            Next
            txtNo.Text = marrKey(0)
        End If
        
        mblnLoad = False
        If txtNo <> "" Then cmdRefresh_Click
    End If
    
    mblnActivate = True
    
    Exit Sub
    
errHandle:
    mblnActivate = True
    Call ErrCenter
End Sub


Private Sub RefPatiData()
    '显示要接单病人的信息
    
    If mPatient Is Nothing Then
        txtBase(Base.挂号单号) = ""
        txtBase(Base.就诊时间) = ""
        txtBase(Base.姓名) = ""
        txtBase(Base.性别) = ""
        txtBase(Base.年龄) = ""
        txtBase(Base.医生) = ""
        txtBase(Base.病人科室) = ""
        txtBase(Base.诊断) = ""
        
        '--  初始化顺序号
        txtBase(Base.顺序号).Tag = ""
        txtBase(Base.顺序号) = ""
        Call GroupToVsFlex(1)
        Exit Sub
    Else
        txtBase(Base.挂号单号) = mPatient.挂号单
        txtBase(Base.就诊时间) = mPatient.挂号时间
        txtBase(Base.姓名) = mPatient.姓名
        txtBase(Base.性别) = mPatient.性别
        txtBase(Base.年龄) = mPatient.年龄
        txtBase(Base.医生) = mPatient.医生
        txtBase(Base.病人科室) = mPatient.病人科室
        txtBase(Base.诊断) = mPatient.门诊诊断
        
        '--  初始化顺序号
        If mPatient.顺序号 = "0" Then
            txtBase(Base.顺序号).Tag = mPatient.Get顺序号
            txtBase(Base.顺序号) = txtBase(Base.顺序号).Tag
        Else
            txtBase(Base.顺序号) = mPatient.顺序号
        End If
    End If
    
    If mGps治疗.Count <= 0 Then
        stabType.TabVisible(0) = False
        chkPrint(0).Value = 0
        chkPrint(0).Visible = False
    Else
        stabType.TabVisible(0) = True
        chkPrint(0).Visible = True
        stabType.Tab = 0
    End If

    If mGps输液.Count <= 0 Then
        stabType.TabVisible(1) = False
        chkPrint(1).Value = 0
        chkPrint(1).Visible = False
        chkLabel.Visible = False
        chkWristband.Visible = False
    Else
        stabType.TabVisible(1) = True
        chkPrint(1).Value = 1
        chkPrint(1).Visible = True
        chkLabel.Visible = True
        chkWristband.Visible = True
        stabType.Tab = 1
    End If

    If mGps注射.Count <= 0 Then
        stabType.TabVisible(2) = False
        chkPrint(2).Value = 0
        chkPrint(2).Visible = False
    Else
        stabType.TabVisible(2) = True
        chkPrint(2).Visible = True
        stabType.Tab = 2
    End If

    If mGps皮试.Count <= 0 Then
        stabType.TabVisible(3) = False
        chkPrint(3).Value = 0
        chkPrint(3).Visible = False
    Else
        stabType.TabVisible(3) = True
        chkPrint(3).Visible = True
        stabType.Tab = 3
    End If

    If mblnHaveData Then
        
        Call stabType_Click(-1)
    End If
    
End Sub
Private Sub Form_Load()
    Dim curDate As Date, i As Integer, ObjOutNurse As OutNurse, Y As Integer
    Dim strPara As String
    
    mstrPrivs = gstrPrivs
    
    '调整为接24小时内的医嘱
    dtpBegin = mDateBegin
    dtpEnd = mdateEnd
    
    mstrKeyType = Trim(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "接单查询类型", ""))
    mintLabelState = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印输液瓶签", "0"))
    mintWristband = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印输液腕带", "0"))
    
    '预计开始时间
    curDate = zldatabase.Currentdate
    For i = 0 To dtpForecastTime.Count - 1
        dtpForecastTime(i).Value = curDate
    Next


    '--- 初始化护士列表
    For i = 0 To cboOperator.Count - 1
        cboOperator(i).Clear
    Next
    
    For Each ObjOutNurse In mOutNurses
        For i = 0 To cboOperator.Count - 1
            cboOperator(i).AddItem ObjOutNurse.姓名
        Next
    Next
    For Y = 0 To cboOperator.Count - 1
        If cboOperator(Y).ListCount > 0 Then
            For i = 0 To cboOperator(Y).ListCount - 1
                If cboOperator(Y).List(i) = UserInfo.姓名 Then
                    cboOperator(Y).ListIndex = i
                    Exit For
                Else
                    cboOperator(Y).ListIndex = 0
                End If
            Next
            
        End If
    Next
    
    '接单后直接穿刺
    mblnImmediatePuncture = zldatabase.GetPara("接单直接穿刺", glngSys, glngModul, "0") = "1"
            
    '85046
    'mblnLiquid = GetDeptInListPara("无线输液_配液科室列表", mlng部门ID)
    strPara = zldatabase.GetPara("待配液科室列表", glngSys, 1264, "")
    mblnLiquid = InStr("," & strPara & ",", "," & mlng部门ID & ",") > 0
    
    If mblnLiquid Then
        Me.cboOperator(3).Visible = False
        Me.lblOperator(3).Visible = False
    Else
        Me.cboOperator(3).Visible = True
        Me.lblOperator(3).Visible = True
    End If
    Me.Caption = "接单 " & mstr执行科室
    '提该病人,该科室,指定时间内的 医嘱发送单
    
'    mnuGNo.Checked = True
'    mnuMNo.Checked = False
'    mnuName.Checked = False
'    mnuSFZ.Checked = False
'    mnuICK.Checked = False
'    mnuJZK.Checked = False
    
    cmdOk.Enabled = False
    '创建/初始化一卡通部件
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle) Then
        Set mobjSquareCard = Nothing
        MsgBox "医疗卡部件（zl9CardSquare）初始化失败！", vbInformation, gstrSysName
    Else
        mstrSquareCards = mobjSquareCard.zlGetIDKindStr(mstrSquareCards)
    End If

'    '读卡部件
'    Set mobjIDCard = New clsIDCard
'    On Error Resume Next
'    Set mobjICCard = CreateObject("zlICCard.clsICCard")
'    On Error GoTo 0
'
'    '初始化弹出菜单，必须在mstrSquareCards后面
'    If mstrSquareCards <> "" Then
'        Dim arrVal As Variant
'        Dim strName As String
'        Dim objMenuPopItem As Object
'        Dim blnAdd As Boolean
'
'        arrVal = Split(mstrSquareCards, ";")
'        For i = LBound(arrVal) To UBound(arrVal)
'            strName = Split(arrVal(i), "|")(enuCardProperty.全名)
'            If InStr(";就诊卡;门诊号;单据号;姓名;身份证;IC卡;ＩＣ卡;", ";" & strName & ";") = 0 Then
'                If blnAdd = False Then
'                    '第一个菜单项
'                    blnAdd = True
'                Else
'                    Load mnuCardSquare(mnuCardSquare.UBound + 1)
'                End If
'                '设置菜单项
'                With mnuCardSquare(mnuCardSquare.UBound)
'                    .Caption = strName
'                    .Tag = arrVal(i)
'                    .Visible = True
'                End With
'            End If
'        Next
'    End If
    
    '读卡改用IDKindNew控件
    idkSelect.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE, txtNo
    idkSelect.IDKind = 1
    For i = 1 To idkSelect.ListCount
        If idkSelect.Cards(i).名称 = mstrKeyType Then
            idkSelect.IDKind = i
            Exit For
        End If
    Next
    
    chkLabel.Value = IIf(mintLabelState = 1, 1, 0)
    chkWristband.Value = IIf(mintWristband = 1, 1, 0)
    
    mblnLoad = True
End Sub

Public Function ShowIncepBill(ByVal bytType As Byte, ByVal lng部门ID As Long, ByVal str执行部门 As String, ByVal str座位 As String, _
                              ByVal DateBegin As Date, ByVal DateEnd As Date, ByRef objPatients As cPatients, _
                              objOutNurses As OutNurses, frmMain As Form, _
                              Optional strNO As String, Optional strJZK As String, _
                              Optional strName As String, Optional ByVal ptiVar As PatiIdentify) As Boolean
'功能：接单功能接口
'参数：
'  bytType：0-点击接单按钮方式；1-自动调用接单（待接单病人）

    mbytType = bytType
    Set mptiInfo = ptiVar
    
    '新的接单入口
    Set mPatients = objPatients
    Set mSeatings = objPatients.mSeatings
    Set mOutNurses = objOutNurses
    
    mlng部门ID = lng部门ID
    mstr执行科室 = str执行部门
    mstr座位 = str座位
    If DateDiff("d", DateBegin, DateEnd) > 7 Then
        '大于7天，强制10天间隔，但可以人为调整界面的时间，不过不能超30天。
        mDateBegin = Format(DateEnd - 9, "yyyy-MM-dd 00:00:00")
    Else
        mDateBegin = DateBegin
    End If
    mdateEnd = Format(DateEnd, "yyyy-MM-dd 23:59:59")
    mblnOk = False
    
    ReDim marrKey(3)
    marrKey(0) = strNO          '病人信息文本
    marrKey(1) = strJZK         '就诊卡号
    marrKey(2) = strName        '病人姓名
    
    Me.Show vbModal, frmMain
    
    ShowIncepBill = mblnOk
    
End Function

Public Function InceptBill() As Boolean

    '接单窗体启动函数
    Dim strPar As String
    Dim dateS As Date, dateE As Date

    mblnHaveData = False
    strPar = zldatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    
    dateS = dtpBegin.Value
    dateE = dtpEnd.Value
    
    If mGps治疗 Is Nothing Then Set mGps治疗 = New Groups
    If Val(Split(strPar, ",")(0)) = 1 Then
        Call mGps治疗.GetGroups(mPatient.病人ID, mlng部门ID, 0, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源)
        If mGps治疗.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGps输液 Is Nothing Then Set mGps输液 = New Groups
    If Val(Split(strPar, ",")(1)) = 1 Then
        Call mGps输液.GetGroups(mPatient.病人ID, mlng部门ID, 1, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源)
        If mGps输液.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGps注射 Is Nothing Then Set mGps注射 = New Groups
    If Val(Split(strPar, ",")(2)) = 1 Then
        Call mGps注射.GetGroups(mPatient.病人ID, mlng部门ID, 2, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源)
        If mGps注射.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGps皮试 Is Nothing Then Set mGps皮试 = New Groups
    If Val(Split(strPar, ",")(3)) = 1 Then
        Call mGps皮试.GetGroups(mPatient.病人ID, mlng部门ID, 3, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源)
        If mGps皮试.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    Call RefPatiData    '刷新界面显示数据
    
    If Not mblnHaveData Then
        MsgBox "没有可接单据！", vbInformation, gstrSysName
    End If
    
    cmdOk.Enabled = mblnHaveData

End Function

Private Sub initObject()
    Set mGps输液 = Nothing
    Set mGps治疗 = Nothing
    Set mGps注射 = Nothing
    Set mGps皮试 = Nothing
    If Not mPatient Is Nothing Then Call GroupToVsFlex(-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMode As String

    Call initObject
    
    strMode = idkSelect.GetCurCard.名称
    If mbytType = 0 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "接单查询类型", strMode
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印输液瓶签", IIf(chkLabel.Value = 1, "1", "0")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印输液腕带", IIf(chkWristband.Value = 1, "1", "0")
    
    Erase marrKey
    Set mPatient = Nothing
    Set mSeatings = Nothing
    Set mOutNurses = Nothing
    Set mobjSquareCard = Nothing
'    Set mobjIDCard = Nothing
    mblnHaveData = False
    mblnActivate = False
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PopupMenu popMenu
End Sub

Private Sub idkSelect_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtNo.Enabled And txtNo.Visible Then
        txtNo.Text = ""
        txtNo.SetFocus
    End If
End Sub

Private Sub idkSelect_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtNo.Text = objPatiInfor.卡号
    mblnReadCard = True
    Call txtNo_KeyPress(0)
End Sub

Private Sub stabType_Click(PreviousTab As Integer)
    Call GroupToVsFlex(stabType.Tab)
End Sub

Private Sub GroupToVsFlex(ByVal intType As Integer)
    Dim ObjGroups As Groups
    Dim objGroup As Group
    Dim objBIll As Bill
    Dim lng序号 As Long, lng总容量 As Long
    Dim strHead As String
    Dim dateS As Date, dateE As Date
    
    If InStr(",10,15,20,", "," & Val(txtTransfusion(滴系数).Text) & ",") <= 0 Then
        txtTransfusion(滴系数).Text = Val(zldatabase.GetPara("默认滴系数", glngSys, 1264))
        If InStr(",10,15,20,", "," & Val(txtTransfusion(滴系数).Text) & ",") <= 0 Then txtTransfusion(滴系数).Text = 20
    End If
    
    strHead = ",280,4;顺序,450,4;上次,450,4;序号,450,4;内容,2800,1;剂量,800,7;单位,450,4;金额,800,7;执行频率,1200,1;用法,1000,1;滴速,450,7;容量(ml),500,7;时间(分),500,4;剩余次数,450,7;输液费,700,7;备注,1200,1;billKey,0,1;GroupKey,0,1;计费状态,0,1;明细计费状态,0,1"
    If Not mPatient Is Nothing Then
        dateS = dtpBegin.Value
        dateE = dtpEnd.Value
        
        Select Case intType
        Case 0
            If mGps治疗 Is Nothing Then
                Set mGps治疗 = New Groups
                If mGps治疗.GetGroups(mPatient.病人ID, mlng部门ID, 0, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源) = True Then
                    Set ObjGroups = mGps治疗
                End If
            Else
                Set ObjGroups = mGps治疗
            End If
            '1 左对齐 4 居中 7 右对齐
            strHead = ",280,4;顺序,0,4;上次,0,4;序号,450,4;内容,2800,1;单量,800,7;单位,450,4;金额,0,7;执行频率,1200,1;用法,1000,1;滴速,0,7;容量(ml),0,7;时间(分),0,4;剩余次数,450,7;治疗费,700,7;备注,1200,1;billKey,0,1;GroupKey,0,1;计费状态,0,1;明细计费状态,0,1"
        Case 1
            If mGps输液 Is Nothing Then
                Set mGps输液 = New Groups
                mGps输液.p滴系数 = Val(txtTransfusion(滴系数))
                If mGps输液.GetGroups(mPatient.病人ID, mlng部门ID, 1, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源) = True Then
                    Set ObjGroups = mGps输液
                End If
            Else
                Set ObjGroups = mGps输液
            End If
            '1 左对齐 4 居中 7 右对齐
            strHead = ",280,4;顺序,450,4;上次,450,4;序号,450,4;内容,2800,1;剂量,800,7;单位,450,4;金额,800,7;执行频率,1200,1;用法,1000,1;滴速,450,7;容量(ml),500,7;时间(分),500,4;剩余次数,450,7;输液费,700,7;备注,1200,1;billKey,0,1;GroupKey,0,1;计费状态,0,1;明细计费状态,0,1"
        
        Case 2
            If mGps注射 Is Nothing Then
                Set mGps注射 = New Groups
                If mGps注射.GetGroups(mPatient.病人ID, mlng部门ID, 2, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源) = True Then
                    Set ObjGroups = mGps注射
                End If
            Else
                Set ObjGroups = mGps注射
            End If
            '1 左对齐 4 居中 7 右对齐
            strHead = ",280,4;顺序,0,4;上次,0,4;序号,450,4;内容,2800,1;剂量,800,7;单位,450,4;金额,800,7;执行频率,1200,1;用法,1000,1;滴速,0,7;容量(ml),0,7;时间(分),0,4;剩余次数,450,7;注射费,700,7;备注,1200,1;billKey,0,1;GroupKey,0,1;计费状态,0,1;明细计费状态,0,1"
        
        Case Else
            If mGps皮试 Is Nothing Then
                Set mGps皮试 = New Groups
                If mGps皮试.GetGroups(mPatient.病人ID, mlng部门ID, 3, dateS, dateE, mPatient.挂号单, mPatient.Key, mPatient.病人来源) = True Then
                    Set ObjGroups = mGps皮试
                End If
            Else
                Set ObjGroups = mGps皮试
            End If
            '1 左对齐 4 居中 7 右对齐
            strHead = ",280,4;顺序,0,4;上次,0,4;序号,450,4;内容,2800,1;剂量,0,7;单位,0,4;金额,0,7;执行频率,1200,1;用法,1000,1;滴速,0,7;容量(ml),0,7;时间(分),0,4;剩余次数,450,7;皮试费,700,7;备注,1200,1;billKey,0,1;GroupKey,0,1;计费状态,0,1;明细计费状态,0,1"
        End Select
    End If
    vsTransfusion.Redraw = flexRDNone
    vsTransfusion.Rows = 2
    vsTransfusion.Clear
    Call SetVsFlexGridHead(strHead, vsTransfusion)
    
    
    lng总容量 = 0
    
    If ObjGroups Is Nothing Then Exit Sub
    For Each objGroup In ObjGroups
        lng序号 = 0
        With vsTransfusion
            For Each objBIll In objGroup.BillsItem(objGroup.执行医嘱ID & "_" & objGroup.发送号)
                lng序号 = lng序号 + 1

                .TextMatrix(.Rows - 1, col_执行顺序) = IIf(objGroup.组次 = 0, "", objGroup.组次)
                .RowData(.Rows - 1) = objGroup.执行状态
                '状态 0-未执行;1-完全执行;2-拒绝执行;3-正在执行
                Call ShowPic(.Rows - 1, objGroup.组次)   ' objGroup.组次 固定显示未选择，让用户必须再选择一次，以便验证一卡通收费
                .TextMatrix(.Rows - 1, col_上次顺序) = objGroup.上次组次
                
                .TextMatrix(.Rows - 1, col_序号) = lng序号
                .TextMatrix(.Rows - 1, col_医嘱内容) = objBIll.医嘱内容
                .TextMatrix(.Rows - 1, col_剂量) = IIf(Left(CStr(objBIll.单量), 1) = ".", "0", "") & CStr(objBIll.单量)
                .TextMatrix(.Rows - 1, col_单位) = objBIll.单位
                .TextMatrix(.Rows - 1, col_项目金额) = IIf(Format(objBIll.金额, "0.00") = 0, "", Format(objBIll.金额, "0.00"))
                
                If objBIll.明细计费状态 = -1 Then
                    .TextMatrix(.Rows - 1, col_项目金额) = "不计费"
                ElseIf objBIll.明细计费状态 = -2 Then
                    If objBIll.金额 = 0 Then .TextMatrix(.Rows - 1, col_项目金额) = "零费用"
                ElseIf objBIll.明细计费状态 = -3 Then
                    .TextMatrix(.Rows - 1, col_项目金额) = "已退费"
                End If
                .TextMatrix(.Rows - 1, col_执行频率) = objGroup.执行频次 '要合并单元
                .TextMatrix(.Rows - 1, col_用法) = objGroup.用法 '要合并单元
                .TextMatrix(.Rows - 1, col_滴速) = objGroup.滴速 '要合并单元
                .TextMatrix(.Rows - 1, col_容量) = objBIll.容量
                
                 objBIll.时间 = CacleTransTime(objBIll.容量, Val(txtTransfusion(滴系数)), ObjGroups.Item(objGroup.执行医嘱ID & "_" & objGroup.发送号).滴速)
                .TextMatrix(.Rows - 1, col_时间) = objBIll.时间
                .TextMatrix(.Rows - 1, col_医生嘱托) = objBIll.医生嘱托
                .TextMatrix(.Rows - 1, col_剩余次数) = objGroup.发送数次 - objGroup.已执行数次 '要合并单元
                
                If objGroup.收费金额 > 0 Then .TextMatrix(.Rows - 1, col_收费金额) = Format(objGroup.收费金额, "0.00") ''要合并单元
                If objGroup.计费状态 = -1 Then
                    .TextMatrix(.Rows - 1, col_收费金额) = "不计费"
                ElseIf objGroup.计费状态 = -2 Then
                    If objGroup.收费金额 = 0 Then .TextMatrix(.Rows - 1, col_收费金额) = "零费用"
                ElseIf objGroup.计费状态 = -3 Then
                     .TextMatrix(.Rows - 1, col_收费金额) = "已退费"
                End If
                .TextMatrix(.Rows - 1, col_BillKey) = objGroup.执行医嘱ID & "_" & objBIll.医嘱ID
                .TextMatrix(.Rows - 1, col_groupkey) = objGroup.执行医嘱ID & "_" & objGroup.发送号
                
                .TextMatrix(.Rows - 1, col_执行计费状态) = objGroup.计费状态
                .TextMatrix(.Rows - 1, col_明细计费状态) = objBIll.明细计费状态
                lng总容量 = lng总容量 + objBIll.容量
                .Rows = .Rows + 1
            Next
            
        End With
    Next
    
    If vsTransfusion.Rows > 2 Then
        vsTransfusion.RemoveItem vsTransfusion.Rows - 1
    End If
    
    txtTransfusion(液体总量) = lng总容量     '可由用户修改
    txtTransfusion(液体总量).Tag = lng总容量 '保存原始计算数据,可用于恢复
    
    '合计总时间
    txtTransfusion(预计时间) = vsTransfusion.Aggregate(flexSTSum, 1, col_时间, vsTransfusion.Rows, col_时间)
    
    '单元格合并
    With vsTransfusion
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(col_剂量)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, .Cols - 1
        .Editable = flexEDKbdMouse
        .Cell(flexcpBackColor, 1, col_滴速, .Rows - 1, col_容量) = VsModiBackColor
        .Cell(flexcpBackColor, 1, col_医生嘱托, .Rows - 1, col_医生嘱托) = VsModiBackColor
        .Redraw = True
    End With
    
End Sub

Private Sub txtNo_Change()
    idkSelect.SetAutoReadCard Trim(txtNo.Text) = ""
End Sub

'Private Sub txtNo_Change()
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtNo.Text = "" And Me.ActiveControl Is txtNo
'    End If
'End Sub

Private Sub txtNo_GotFocus()
    Call zlControl.TxtSelAll(txtNo)
    idkSelect.SetAutoReadCard Trim(txtNo.Text) = ""
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    '按回车
    Dim strCard As String
    
    On Error GoTo hErr
    
    strCard = idkSelect.Cards(idkSelect.IDKind).名称
   
    If mblnReadCard Or KeyAscii = 13 Then
'        If KeyAscii <> 13 Then
'            txtNo.Text = txtNo.Text & Chr(KeyAscii)
'            txtNo.SelStart = Len(txtNo.Text)
'        End If
'        KeyAscii = 0
        Call cmdRefresh_Click
        Call zlControl.TxtSelAll(txtNo)
    Else
        Select Case strCard
            Case "门诊号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "挂号单"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtNo.Text = "" Or txtNo.SelLength = Len(txtNo.Text)) _
                       And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "身份证号", "二代身份证"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
    mblnReadCard = False
    
    Exit Sub

hErr:
    mblnReadCard = False
    LogWrite "输液接单的调试日志", "" & glngModul, "txtNo_KeyPress", "单据输入框，第" & CStr(Erl()) & "行，" & Err.Description
End Sub

Private Sub txtNo_LostFocus()
    idkSelect.SetAutoReadCard False
End Sub

'Private Sub txtNo_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Private Sub txtTransfusion_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 滴系数 Then
        If KeyAscii = vbKeyReturn Then
            Call zlcommfun.PressKey(vbKeyTab)
        ElseIf InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTransfusion_LostFocus(Index As Integer)
    If Index = 滴系数 Then
        Call stabType_Click(-1)
    End If
End Sub

Private Sub txtTransfusion_Validate(Index As Integer, Cancel As Boolean)
'    If Index = 滴系数 Then
'        If txtTransfusion(Index).Text < 10 Or txtTransfusion(Index).Text > 50 Then
'            txtTransfusion(Index).Text = 20
'        End If
'    End If
End Sub

Private Sub vsTransfusion_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strGroupKey As String, strBillKey As String
    Dim objBIll As Bill
    
    With vsTransfusion
    strGroupKey = .TextMatrix(Row, col_groupkey)
    strBillKey = .TextMatrix(Row, col_BillKey)
    
    Select Case Col
    
    Case col_滴速
        If stabType.Tab = 1 Then
            If mGps输液.Count > 0 Then
                mGps输液.Item(strGroupKey).滴速 = Val(.TextMatrix(Row, Col))
                For Each objBIll In mGps输液.Item(strGroupKey).BillsItem(strGroupKey)
                    objBIll.时间 = CacleTransTime(objBIll.容量, Val(txtTransfusion(滴系数)), mGps输液.Item(strGroupKey).滴速)
                    strBillKey = mGps输液.Item(strGroupKey).执行医嘱ID & "_" & objBIll.医嘱ID
                    mGps输液.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).时间 = objBIll.时间
                Next
                Call stabType_Click(-1)
                
            End If
        End If
    Case col_容量
        If stabType.Tab = 1 Then
            If mGps输液.Count > 0 Then
                mGps输液.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).容量 = Val(.TextMatrix(Row, Col))
                mGps输液.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).时间 = _
                CacleTransTime(mGps输液.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).容量, Val(txtTransfusion(滴系数)), mGps输液.Item(strGroupKey).滴速)
                Call stabType_Click(-1)
            End If
        End If
    Case col_医生嘱托
        Select Case stabType.Tab
        Case 0
            If mGps治疗.Count > 0 Then
                mGps治疗.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).医生嘱托 = .TextMatrix(Row, Col)
            End If
        Case 1
            If mGps输液.Count > 0 Then
                mGps输液.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).医生嘱托 = .TextMatrix(Row, Col)
            End If
        Case 2
            If mGps注射.Count > 0 Then
                mGps注射.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).医生嘱托 = .TextMatrix(Row, Col)
            End If
        Case Else
            If mGps皮试.Count > 0 Then
                mGps皮试.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).医生嘱托 = .TextMatrix(Row, Col)
            End If
        End Select
    End Select
    End With
End Sub

Private Sub vsTransfusion_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    vsTransfusion.AutoSize 1, col_groupkey
End Sub

Private Sub vsTransfusion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col_滴速 Or Col = col_容量 Or Col = col_医生嘱托) Then Cancel = True
End Sub

Private Sub vsTransfusion_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)

    Dim LeftCol As Long, RightCol As Long, topRow As Long, BottomRow As Long
    
    If Not MergeRow(Row, topRow, BottomRow) Then Exit Sub '非合并行,退出
    If topRow = BottomRow Then Exit Sub
    
    LeftCol = col_执行顺序: RightCol = col_上次顺序
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
    LeftCol = col_执行频率: RightCol = col_滴速
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
    LeftCol = col_剩余次数: RightCol = col_医生嘱托
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
End Sub

Private Sub vsTransfusion_EnterCell()
    If vsTransfusion.Col = col_滴速 Or vsTransfusion.Col = col_容量 Or vsTransfusion.Col = col_医生嘱托 Then
        Call vsTransfusion.CellBorder(vsTransfusion.GridColor, 2, 2, 3, 3, 0, 0)
    End If
End Sub

Private Sub vsTransfusion_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsTransfusion
        If Not (.Col = col_滴速 Or .Col = col_容量 Or .Col = col_医生嘱托) Then
            If KeyCode = vbKeySpace And .Row > 0 Then
                Call CheckGroup(.Row, col_选择, 1)
            ElseIf KeyCode = vbKeyDelete Then
                Call CheckGroup(.Row, col_选择, 2)
            End If
        End If
    End With
End Sub

Private Sub vsTransfusion_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim NextCol As Long
    If KeyAscii = vbKeyReturn Then
        NextCol = EditCol(Col)
        With vsTransfusion
            If NextCol = -1 Then
                If Row + 1 <= .Rows - 1 Then
                    If .TextMatrix(Row + 1, col_序号) = 1 Then
                        .Select Row + 1, col_滴速
                    Else
                        .Select Row + 1, col_容量
                    End If
                Else
                    If .TextMatrix(Row, col_序号) = 1 Then
                        .Select Row, col_滴速
                    Else
                        .Select Row, col_容量
                    End If
                End If
            Else
                .Select Row, NextCol
            End If
        End With
    End If
End Sub

Private Function EditCol(ByVal Col As Long) As Long
    '反回当前列之后的可编辑列,-1表示当前列之后无可编辑行.
    Dim lngCol As Long
    
    If Col + 1 > vsTransfusion.Cols - 1 Then
        EditCol = -1
        Exit Function
    End If
    
    For lngCol = Col + 1 To vsTransfusion.Cols - 1
        If InStr(",8,9,13,", "," & CStr(lngCol) & ",") > 0 Then
            EditCol = lngCol
            Exit Function
        End If
    Next
    If EditCol = 0 Then EditCol = -1
    
End Function

Private Sub vsTransfusion_LeaveCell()
    If vsTransfusion.Col = col_滴速 Or vsTransfusion.Col = col_容量 Or vsTransfusion.Col = col_医生嘱托 Then
        On Error Resume Next
        Call vsTransfusion.CellBorder(vsTransfusion.GridColor, 0, 0, 0, 0, 0, 0)
    End If
End Sub

Private Sub vsTransfusion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsTransfusion
        Call CheckGroup(.MouseRow, .MouseCol, Button)
    End With
End Sub

Private Sub CheckGroup(ByVal Row As Long, ByVal Col As Long, ByVal Button As Integer)
    Dim blnCheck As Boolean, StrKey As String, lngRows As Long, lngCurRow As Long, i As Long
    Dim lngCol As Long, lngRow As Long, strTmpKey As String
    Dim lng医嘱ID As Long, lng发送号 As Long, intReturn As Integer    '一卡通调用时使用
    Dim blnOK As Boolean
    
    With vsTransfusion
    lngCol = Col: lngRow = Row
    If lngCol = col_选择 And lngRow > 0 Then
        
        StrKey = .TextMatrix(lngRow, col_groupkey)
        If InStr(StrKey, "_") > 0 Then
            If Button = 1 And InStr(",0,3,", "," & .RowData(lngRow) & ",") > 0 Then
                '接单
                'blnCheck = .Cell(flexcpPicture, lngRow, col_选择) = imgPic.ListImages(2).Picture
                
                blnCheck = Val(.TextMatrix(lngRow, col_执行顺序)) = 0
                '一卡通收费检查
                Select Case stabType.Tab
                Case 0
                    lng医嘱ID = mGps治疗.Item(StrKey).执行医嘱ID
                    lng发送号 = mGps治疗.Item(StrKey).发送号

                Case 1
                    lng医嘱ID = mGps输液.Item(StrKey).执行医嘱ID
                    lng发送号 = mGps输液.Item(StrKey).发送号

                Case 2
                    lng医嘱ID = mGps注射.Item(StrKey).执行医嘱ID
                    lng发送号 = mGps注射.Item(StrKey).发送号

                Case Else
                    lng医嘱ID = mGps皮试.Item(StrKey).执行医嘱ID
                    lng发送号 = mGps皮试.Item(StrKey).发送号
                   
                End Select
                '2012-11-08 检查剩余可执行次数
                If Not CheckRun(lng医嘱ID, lng发送号) Then Exit Sub

                
                If .TextMatrix(lngRow, col_明细计费状态) = -3 Then
                    MsgBox "明细项目已退费，不能执行此操作！", vbInformation, Me.Caption
                    Exit Sub
                End If
                
                intReturn = OneCardCheck(lng医嘱ID, lng发送号, Me, mobjSquareCard)
                If intReturn = 0 Then
                    '老流程
                    If InStr(mstrPrivs, "执行项目未收费接单") <= 0 And blnCheck Then
                        If .TextMatrix(lngRow, col_执行计费状态) <> -2 Then '零费用，当成已经收过费了2012－07－16
                            If (.TextMatrix(lngRow, col_执行计费状态) > -1) And Val(.TextMatrix(lngRow, col_收费金额)) = 0 Then
                                MsgBox .TextMatrix(0, col_收费金额) & "未收取，请收费后再操作！", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                    If InStr(mstrPrivs, "明细项目未收费接单") <= 0 And blnCheck Then
                        If .TextMatrix(lngRow, col_执行计费状态) <> -2 Then '零费用，当成已经收过费了2012－07－16
                            If (.TextMatrix(lngRow, col_明细计费状态) > -1) And Val(.TextMatrix(lngRow, col_项目金额)) = 0 Then
                                MsgBox "明细项目未收费，请收费后再操作！", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    
                ElseIf intReturn = 2 Then
                    '一卡通的流程失败,函数内部有提示，直接退出
                    Exit Sub
                End If
                
                
                Select Case stabType.Tab
                Case 0
                    lngRows = mGps治疗.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps治疗.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_执行顺序) = IIf(mGps治疗.Item(strTmpKey).组次 = 0, "", mGps治疗.Item(strTmpKey).组次)
                        If Val(.TextMatrix(lngCurRow, col_执行顺序)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case 1
                    lngRows = mGps输液.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps输液.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_执行顺序) = IIf(mGps输液.Item(strTmpKey).组次 = 0, "", mGps输液.Item(strTmpKey).组次)
                        If Val(.TextMatrix(lngCurRow, col_执行顺序)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case 2
                    lngRows = mGps注射.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps注射.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_执行顺序) = IIf(mGps注射.Item(strTmpKey).组次 = 0, "", mGps注射.Item(strTmpKey).组次)
                        If Val(.TextMatrix(lngCurRow, col_执行顺序)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case Else
                    lngRows = mGps皮试.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps皮试.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_执行顺序) = IIf(mGps皮试.Item(strTmpKey).组次 = 0, "", mGps皮试.Item(strTmpKey).组次)
                        If Val(.TextMatrix(lngCurRow, col_执行顺序)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                End Select
                chkPrint(stabType.Tab).Value = IIf(blnOK, 1, 0)
                
            ElseIf Button = 2 And Val(.TextMatrix(lngRow, col_执行顺序)) = 0 Then
                '右键拒绝
                Select Case stabType.Tab
                Case 0
                    If .RowData(lngRow) = 2 Then
                        '取消拒绝
                        blnOK = mGps治疗.Item(StrKey).FuncExecRestore
                    Else
                        '拒绝
                        blnOK = mGps治疗.Item(StrKey).FuncExecRefuse
                    End If
                    If blnOK = False Then Exit Sub
                    
                    lngRows = mGps治疗.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps治疗.Item(strTmpKey).执行状态
                    Next
                Case 1
                    If .RowData(lngRow) = 2 Then
                        '取消拒绝
                        blnOK = mGps输液.Item(StrKey).FuncExecRestore
                    Else
                        '拒绝
                        blnOK = mGps输液.Item(StrKey).FuncExecRefuse
                    End If
                    If blnOK = False Then Exit Sub
                    
                    lngRows = mGps输液.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps输液.Item(strTmpKey).执行状态
                    Next
                Case 2
                    If .RowData(lngRow) = 2 Then
                        '取消拒绝
                        blnOK = mGps注射.Item(StrKey).FuncExecRestore
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 0
                    Else
                        '拒绝
                        blnOK = mGps注射.Item(StrKey).FuncExecRefuse
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 2
                    End If
                    lngRows = mGps注射.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps注射.Item(strTmpKey).执行状态
                    Next
                Case 3
                    If .RowData(lngRow) = 2 Then
                        '取消拒绝
                        blnOK = mGps皮试.Item(StrKey).FuncExecRestore
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 0
                    Else
                        '拒绝
                        blnOK = mGps皮试.Item(StrKey).FuncExecRefuse
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 2
                    End If
                    lngRows = mGps皮试.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps皮试.Item(strTmpKey).执行状态
                    Next
                End Select
            End If
            '--- 更新图片
            lngCurRow = Val(.TextMatrix(lngRow, col_序号))
            
            If lngRows = lngCurRow Then
                If lngRows = 1 Then
                    Call ShowPic(lngRow, Val(.TextMatrix(lngRow, col_执行顺序)))
                    Exit Sub
                Else
                    For i = 1 To lngRows
                       Call ShowPic(lngRow - (lngCurRow - i), Val(.TextMatrix(lngRow - (lngCurRow - i), col_执行顺序)))
                    Next
                End If
            Else
                For i = 1 To lngCurRow
                    Call ShowPic(lngRow - (lngCurRow - i), Val(.TextMatrix(lngRow - (lngCurRow - i), col_执行顺序)))
                Next
                
                For i = lngCurRow To lngRows
                    Call ShowPic(lngRow + (lngRows - i), Val(.TextMatrix(lngRow + (lngRows - i), col_执行顺序)))
                Next
            End If
        End If
    End If
    .Refresh
    End With
End Sub

Private Sub vsTransfusion_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case col_滴速
        If Val(vsTransfusion.EditText) < 10 Or Val(vsTransfusion.EditText) > 100 Then
            Cancel = True
        End If
    Case col_容量
        If Val(vsTransfusion.EditText) < 0 Or Val(vsTransfusion.EditText) > 10000 Then
            Cancel = True
        End If
    End Select
End Sub

Private Function MergeRow(ByVal Row As Long, topRow, BottomRow As Long) As Boolean
    '是否合并行
    Dim strGroupKey As String, lngRow As Long
    With vsTransfusion
        If .Cols < col_groupkey Then Exit Function
        strGroupKey = .TextMatrix(Row, col_groupkey)
        topRow = Row: BottomRow = Row
        For lngRow = Row To 0 Step -1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                topRow = lngRow + 1
                Exit For
            Else
                topRow = lngRow
            End If
        Next
        
        For lngRow = Row To .Rows - 1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                BottomRow = lngRow - 1
                Exit For
            Else
                BottomRow = lngRow
            End If
        Next
    End With

    If topRow > 0 And BottomRow > 0 Then MergeRow = True
End Function

Private Sub ShowPic(ByVal Row As Long, ByVal 组次 As Long)
    '更新指定行的图片
    If Row <= 0 Then Exit Sub
    With vsTransfusion
        '状态 0-未执行;1-完全执行;2-拒绝执行;3-正在执行
    
        If .RowData(Row) = 0 Or .RowData(Row) = 3 Then
            '0 未选择 3-正在执行 也当成未执行处理,因为有 分成几次在执行的情况
            If 组次 > 0 Then
                Set .Cell(flexcpPicture, Row, col_选择) = imgPic.ListImages(2).Picture
            Else
                Set .Cell(flexcpPicture, Row, col_选择) = imgPic.ListImages(1).Picture
            End If
        ElseIf .RowData(Row) = 1 Then
            '1 完成
            Set .Cell(flexcpPicture, Row, col_选择) = imgPic.ListImages(4).Picture
        ElseIf .RowData(Row) = 2 Then
            Set .Cell(flexcpPicture, Row, col_选择) = imgPic.ListImages(3).Picture

        End If
    End With
End Sub
Private Function Get最大已销(ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng组ID As Long, ByVal str诊疗类别 As String) As Long
'功能：获取某条医嘱，或某组医嘱的最大已销帐的医嘱执行次数
'       bln单独执行 是否单独执行，检验检查类存在单据的医嘱的单独执行某一部位，某一部分检查
'       lng医嘱ID 该条医嘱ID
'       lng组ID 没有父医嘱，或者父医嘱时为医嘱ID,子医嘱为相关ID
'       str诊疗类别 该医嘱的诊疗类别
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If bln单独执行 Then
        lng组ID = lng医嘱ID
        strSQL = "Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 最大已销数" & vbNewLine & _
                "From 门诊费用记录 A, 病人医嘱计价 B" & vbNewLine & _
                "Where a.医嘱序号 = [1] And b.医嘱id = a.医嘱序号 And b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And mod(a.记录性质,10) in(1,2) And a.价格父号 Is Null And" & vbNewLine & _
                "      a.收费类别 Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1)"

    Else
        strSQL = "Select Max(c.已销数) 最大已销数" & vbNewLine & _
                "From (Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 已销数" & vbNewLine & _
                "       From 门诊费用记录 A, 病人医嘱计价 B" & vbNewLine & _
                "       Where a.医嘱序号 In (Select ID From 病人医嘱记录 Where (ID = [1] Or 相关id = [1]) And 诊疗类别 = [2]) And b.医嘱id = a.医嘱序号 And" & vbNewLine & _
                "             b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And mod(a.记录性质,10) in(1,2) And a.价格父号 Is Null And a.收费类别 Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) " & vbNewLine & _
                "       Group By  a.医嘱序号,a.收费细目id) C"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng组ID, str诊疗类别)
    If rsTmp.RecordCount <> 0 Then
        Get最大已销 = Val(rsTmp!最大已销数 & "")
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckRun(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As Boolean
    '检查医嘱是否还可以执行
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng已执行数次 As Long, lng退费数次 As Long
    On Error GoTo errH
    strSQL = "Select " & _
        " Max(执行时间) as LastDate," & _
        " Max(要求时间) as curDate," & _
        " Count(要求时间) as curCount," & _
        " Sum(本次数次) as curNum" & _
        " From 病人医嘱执行" & _
        " Where 医嘱ID=[1] And 发送号=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    If Not rsTmp.EOF Then
        lng已执行数次 = Val("" & rsTmp!curNum)
    End If
    
    '计算本次执行应该的要求时间
    strSQL = "Select A.发送数次,Nvl(B.相关id, B.ID) 组ID,C.计算单位,A.首次时间,A.末次时间,Decode(B.病人来源, 2, Decode(A.记录性质, 1, 1, Decode(A.门诊记帐, 1, 1, 2)), 1) 费用性质," & _
        " B.开始执行时间,B.执行终止时间,B.上次执行时间,B.执行时间方案," & _
        " B.执行频次,B.频率次数,B.频率间隔,B.间隔单位,B.病人ID,b.主页ID,c.类别,c.操作类型,c.执行分类" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C" & _
        " Where A.医嘱ID=B.ID And B.诊疗项目ID=C.ID" & _
        " And A.医嘱ID=[1] And A.发送号=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    If rsTmp!费用性质 = 1 Then
        lng退费数次 = Get最大已销(False, lng医嘱ID, lng医嘱ID, "" & rsTmp!类别)
    
    
        '当前实际已经执行了要求的次数,不准再执行
        If lng已执行数次 + lng退费数次 >= Val("" & rsTmp!发送数次) Then
            MsgBox "该医嘱本次发送允许执行 " & "" & rsTmp!发送数次 & IIf(lng退费数次 <> 0, " 次，" & "相关单据已经退费或销帐" & lng退费数次, "") & "次，当前已经执行了 " & lng已执行数次 & " 次，不能再接单。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckRun = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecutionComplete(ByVal lngDeptID As Long, ByVal objPati As cPatient) As Boolean
'功能：门诊医嘱（输液、皮试、注射）执行数次是否完成
'参数：
'  lngDeptID：执行部门ID
'  objPati：病人信息类
'返回：True未完，False已完
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim StrKey As String
    
    '2014-03-19：调整取“执行终止时间”未到期的病人医嘱记录
    '执行终止时间=NULL，用开始执行时间+7天。同时，为了解决急诊医嘱可能执行终止时间比较短的情况，对执行终止时间强制加1天。
    On Error GoTo errHandle
    
    If objPati.病人来源 = 1 Then
        '门诊留观
        StrKey = objPati.Key
        strSQL = "select c.发送数次, Sum(d.本次数次) 执行数次" & vbNewLine & _
                "From 病案主页 a, 病人医嘱记录 B, 病人医嘱发送 C, 病人医嘱执行 D, 诊疗项目目录 E" & vbNewLine & _
                "where a.病人id=b.病人id and a.主页id=b.主页id and b.id=c.医嘱id and c.医嘱id=d.医嘱id(+) and c.发送号=d.发送号(+) " & vbNewLine & _
                "   And b.诊疗项目id = e.Id And b.诊疗类别 = 'E' And (e.执行分类 In (0, 1, 2, 3) or e.执行分类 is null) " & vbNewLine & _
                "   And c.执行部门id = [1] And a.出院日期 is null " & vbNewLine & _
                "   And b.病人id = [2] and b.主页id = [3] " & vbNewLine & _
                "Group By c.医嘱id, c.发送数次" & vbNewLine & _
                "Having Nvl(c.发送数次, 0) - Nvl(Sum(d.本次数次), 0) > 0 "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "医嘱执行数次", lngDeptID, Val(StrKey), Val(Split(StrKey, "_")(1)))
    Else
        StrKey = objPati.挂号单
        strSQL = "Select c.发送数次, Sum(d.本次数次) 执行数次 " & _
                 "From 病人医嘱记录 B, 病人医嘱发送 C, 病人医嘱执行 D, 诊疗项目目录 E " & _
                 "Where b.Id = c.医嘱id And c.医嘱id = d.医嘱id(+) And c.发送号 = d.发送号(+) " & _
                 "  And b.诊疗项目id = e.Id And b.诊疗类别 = 'E' And (e.执行分类 In (0, 1, 2, 3) or e.执行分类 is null) " & _
                 "  And b.开始执行时间 + 7 > Sysdate " & _
                 "  And c.执行部门id = [1] And b.挂号单 = [2] " & _
                 "Group By c.医嘱id, c.发送数次 " & _
                 "Having Nvl(c.发送数次, 0) - Nvl(Sum(d.本次数次), 0) > 0 "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "医嘱执行数次", lngDeptID, StrKey)
    End If
    Do While Not rsTemp.EOF
        If zlcommfun.NVL(rsTemp!发送数次, 0) - zlcommfun.NVL(rsTemp!执行数次, 0) > 0 Then
            ExecutionComplete = True
            Exit Do
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

'Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, _
'    ByVal datBirthDay As Date, ByVal strAddress As String)
''功能：身份证识别成功后激活
'    mstrIDCard = strID
'
'    If idkSelect.GetCurCard.名称 = "身份证号" Then
'        txtNo.Text = mstrIDCard
'    Else
'        txtNo.Text = "" '否则清除(目前是在已清除情况下才能激活)。
'    End If
'
'    If txtNo.Text <> "" Then Call cmdRefresh_Click
'End Sub
'
'
Private Function GetClinicDept(ByVal lngPatiId As Long, ByVal StrKey As String) As String
'功能：获取病人医嘱的开立科室
'参数：
'  lngPatiID：病人ID
'  strKey：挂号单或“病人id_主页id”
'返回：开嘱科室

    Dim lngPageID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If StrKey Like "*_*" Then
        lngPageID = Val(Split(StrKey, "_")(1))
        strSQL = "Select Distinct b.名称 " & vbNewLine & _
                 "From 病人医嘱记录 A, 部门表 B " & vbNewLine & _
                 "Where a.开嘱科室id = b.Id And a.病人id = [1] And 主页id = [2] "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "获取病人的开嘱科室", lngPatiId, lngPageID)
    Else
        strSQL = "Select Distinct b.名称 " & vbNewLine & _
                 "From 病人医嘱记录 A, 部门表 B " & vbNewLine & _
                 "Where a.开嘱科室id = b.Id And a.挂号单 = [1] And a.病人id = [2] "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "获取病人的开嘱科室", StrKey, lngPatiId)
    End If
    If rsTemp.EOF = False Then
        GetClinicDept = zlcommfun.NVL(rsTemp!名称)
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume

End Function
