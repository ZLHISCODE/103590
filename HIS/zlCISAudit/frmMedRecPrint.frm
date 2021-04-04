VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmMedRecPrint 
   Caption         =   "电子病案打印"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmMedRecPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15120
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLine 
      Caption         =   "Frame1"
      Height          =   8655
      Left            =   5160
      MousePointer    =   9  'Size W E
      TabIndex        =   39
      Top             =   120
      Width           =   45
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   4200
      Top             =   240
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   5085
      TabIndex        =   9
      Top             =   0
      Width           =   5085
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   4815
         _Version        =   589884
         _ExtentX        =   8493
         _ExtentY        =   2778
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picShow 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   120
         MouseIcon       =   "frmMedRecPrint.frx":6852
         ScaleHeight     =   270
         ScaleWidth      =   4935
         TabIndex        =   51
         Tag             =   "0"
         Top             =   120
         Width           =   4935
         Begin VB.PictureBox picUpOrDown 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   4560
            Picture         =   "frmMedRecPrint.frx":6B5C
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   52
            Top             =   0
            Width           =   270
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "显示范围查找"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   53
            Top             =   45
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4695
         TabIndex        =   34
         Top             =   6240
         Width           =   4695
         Begin VB.CommandButton cmdSet 
            Caption         =   "报表设置"
            Height          =   300
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   1100
         End
         Begin VB.ComboBox cboPrinterName 
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   120
            Width           =   3375
         End
         Begin VB.CommandButton cmdPreView 
            Caption         =   "预览(&V)"
            Height          =   300
            Left            =   2160
            TabIndex        =   7
            Top             =   600
            Width           =   1100
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "打印(&P)"
            Height          =   300
            Left            =   3360
            TabIndex        =   8
            Top             =   600
            Width           =   1100
         End
         Begin VB.Label lblPlugIn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "扩展功能↓"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   230
            TabIndex        =   64
            Top             =   1020
            Width           =   990
         End
         Begin VB.Label lblPrint 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "输出设备"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame fraFind 
         Caption         =   "直接查找"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。"
         Top             =   2520
         Width           =   4815
         Begin zlIDKind.PatiIdentify PatiIdentifyFind 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。"
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
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
            IDKindStr       =   $"frmMedRecPrint.frx":D3AE
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
            InputAppearance =   0
            ShowSortName    =   -1  'True
            DefaultCardType =   "就诊卡"
            IDKindWidth     =   555
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
      End
      Begin VB.Frame fraScope 
         Caption         =   "范围查找"
         Height          =   1935
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   4815
         Begin VB.ComboBox cboOutTime 
            Height          =   300
            ItemData        =   "frmMedRecPrint.frx":D445
            Left            =   960
            List            =   "frmMedRecPrint.frx":D447
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3495
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   960
            TabIndex        =   0
            Text            =   "cboDept"
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton cmdFind 
            Appearance      =   0  'Flat
            Caption         =   "查找"
            Height          =   300
            Left            =   3720
            Picture         =   "frmMedRecPrint.frx":D449
            TabIndex        =   4
            Top             =   1440
            Width           =   600
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
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
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   238419971
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
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
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   238419971
            CurrentDate     =   39998.8757060185
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "出院科室"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            Caption         =   "出院时间"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   8925
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22490
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   10320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":13C9B
            Key             =   "首页正面"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":16CD5
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":1D537
            Key             =   "检查报告"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":22FF9
            Key             =   "检验报告"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":28ABB
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":2F31D
            Key             =   "Patient"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":301F7
            Key             =   "unCheckAll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":30791
            Key             =   "CheckAll"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":30D2B
            Key             =   "住院病历"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":33D65
            Key             =   "其他报表"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":36D9F
            Key             =   "疾病证明"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":39DD9
            Key             =   "首页附页一"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":3ACB3
            Key             =   "临床路径"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":3DCED
            Key             =   "首页附页二"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":40D27
            Key             =   "护理病历"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":43D61
            Key             =   "住院医嘱"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":46D9B
            Key             =   "护理记录"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":49DD5
            Key             =   "知情文件"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4CE0F
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4D449
            Key             =   "CheckFill"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4DA83
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4E0BD
            Key             =   "首页反面"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":510F7
            Key             =   "down"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":57959
            Key             =   "up"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":5E1BB
            Key             =   "住院证"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   5160
      ScaleHeight     =   7815
      ScaleWidth      =   10095
      TabIndex        =   12
      Top             =   720
      Width           =   10095
      Begin VB.PictureBox picItemInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4080
         ScaleHeight     =   1575
         ScaleWidth      =   2655
         TabIndex        =   37
         Top             =   5640
         Width           =   2655
         Begin VSFlex8Ctl.VSFlexGrid vsItemInfo 
            Bindings        =   "frmMedRecPrint.frx":6164D
            Height          =   555
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   735
            _cx             =   1296
            _cy             =   979
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
      Begin VB.Frame fraPati 
         BackColor       =   &H80000005&
         Caption         =   "病人信息"
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   10000
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   2325
            TabIndex        =   62
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "赵丽颖"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   2325
            TabIndex        =   61
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "重庆市两江新区"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   6840
            TabIndex        =   60
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院医师:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   1500
            TabIndex        =   59
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出院日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   8280
            TabIndex        =   57
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-24"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   9120
            TabIndex        =   56
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出生日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   1500
            TabIndex        =   55
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "地址:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   6360
            TabIndex        =   54
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   6840
            TabIndex        =   28
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "妇产科"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   4560
            TabIndex        =   27
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "500101198810121245"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   4560
            TabIndex        =   26
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   2325
            TabIndex        =   25
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "女"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   6840
            TabIndex        =   24
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28岁"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   9120
            TabIndex        =   23
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "降央卓玛"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   4560
            TabIndex        =   22
            Top             =   360
            Width           =   720
         End
         Begin VB.Image imgPatient 
            Height          =   1185
            Left            =   120
            Picture         =   "frmMedRecPrint.frx":61661
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "入院日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   6000
            TabIndex        =   21
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出院科室:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3720
            TabIndex        =   20
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "身份证号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   3720
            TabIndex        =   18
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "年龄:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   8640
            TabIndex        =   58
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   1680
            TabIndex        =   19
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "性别:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   6360
            TabIndex        =   17
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "姓名:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   4080
            TabIndex        =   16
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.PictureBox picCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   7  'Invert
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   10005
         TabIndex        =   13
         Top             =   1680
         Width           =   10000
         Begin VB.Frame fraCent 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2655
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar vsc 
               Height          =   2295
               Left            =   8640
               Max             =   10
               TabIndex        =   41
               Top             =   120
               Width           =   255
            End
            Begin VB.Frame fraIn 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   2295
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   8415
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FAEADA&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   3
                  Left            =   4320
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   49
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   2
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":6252B
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Height          =   720
                     Index           =   5
                     Left            =   240
                     Picture         =   "frmMedRecPrint.frx":62B55
                     Top             =   240
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "住院病历"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   3
                     Left            =   240
                     TabIndex        =   50
                     Top             =   1080
                     Width           =   720
                  End
               End
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   2
                  Left            =   2760
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   47
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   1
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":65B7F
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Height          =   720
                     Index           =   4
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":661A9
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "检验/检查"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   2
                     Left            =   240
                     TabIndex        =   48
                     Top             =   1080
                     Width           =   810
                  End
               End
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   1
                  Left            =   1380
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   45
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":691D3
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   720
                     Index           =   3
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":697FD
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "住院医嘱"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   1
                     Left            =   240
                     TabIndex        =   46
                     Top             =   1080
                     Width           =   720
                  End
               End
               Begin VB.PictureBox picItem 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   0
                  Left            =   0
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   43
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image imgCHK 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":6C827
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image imgItem 
                     Appearance      =   0  'Flat
                     Height          =   720
                     Index           =   0
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":6CE51
                     Stretch         =   -1  'True
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "首页"
                     ForeColor       =   &H80000001&
                     Height          =   180
                     Index           =   0
                     Left            =   480
                     TabIndex        =   44
                     Top             =   1080
                     Width           =   360
                  End
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   3
                  Visible         =   0   'False
                  X1              =   0
                  X2              =   360
                  Y1              =   0
                  Y2              =   360
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   2
                  Visible         =   0   'False
                  X1              =   1560
                  X2              =   1320
                  Y1              =   1680
                  Y2              =   1500
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   1
                  Visible         =   0   'False
                  X1              =   960
                  X2              =   1200
                  Y1              =   1680
                  Y2              =   2040
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   0
                  Visible         =   0   'False
                  X1              =   1320
                  X2              =   1680
                  Y1              =   1920
                  Y2              =   2280
               End
            End
         End
         Begin VB.Frame fraSplit 
            BackColor       =   &H80000005&
            Height          =   45
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   10480
         End
         Begin VB.Line LineB 
            BorderColor     =   &H80000010&
            X1              =   2160
            X2              =   6480
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line LineR 
            BorderColor     =   &H00E0E0E0&
            X1              =   9480
            X2              =   9480
            Y1              =   2760
            Y2              =   480
         End
         Begin VB.Line LineL 
            BorderColor     =   &H00FFC0C0&
            X1              =   0
            X2              =   0
            Y1              =   1320
            Y2              =   3120
         End
         Begin VB.Line lineT 
            BorderColor     =   &H8000000A&
            X1              =   1920
            X2              =   7200
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   600
            TabIndex        =   33
            Top             =   120
            Width           =   90
         End
         Begin VB.Image imgAll 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   120
            Picture         =   "frmMedRecPrint.frx":6DD1B
            Top             =   60
            Width           =   300
         End
      End
      Begin XtremeSuiteControls.TabControl tbcSub 
         Height          =   495
         Left            =   600
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":6E345
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":74BA7
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7B409
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7B9A3
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7BF3D
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMedRecPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'声明
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'枚举
Private Enum ENUM_COLOR
    COLOR_HIGH = &HFCE8D7
    COLOR_ITEM = &HFDF3E9
End Enum

Private Enum PATIREPORT_COLUMN
    col_选择 = 0
    col_图标 = 1
    col_打印图标 = 2
    col_是否编目 = 3
    col_编目日期 = 4
    col_住院号 = 5
    col_姓名 = 6
    col_性别 = 7
    col_身份证号 = 8
    Col_出生日期 = 9
    col_入院日期 = 10
    col_出院日期 = 11
    col_出院科室 = 12
    coL_住院医师 = 13
    col_家庭地址 = 14
    col_年龄 = 15
    
    '隐藏列
    col_病人类型 = 16
    col_病人Id = col_病人类型 + 1        '隐藏
    col_主页ID = col_病人类型 + 2         '隐藏
    col_出院科室ID = col_病人类型 + 3   '隐藏
    col_打印记录 = col_病人类型 + 4
End Enum

Private Enum PATI_INFO
    lbl_姓名 = 0
    lbl_性别 = 1
    lbl_年龄 = 2
    lbl_身份证号 = 3
    lbl_住院号 = 4
    lbl_出院科室 = 5
    lbl_入院日期 = 6
    lbl_出院日期 = 7
    lbl_家庭地址 = 8
    lbl_住院医师 = 9
    lbl_出生日期 = 10
End Enum

Private Enum TAB_INFO
    tab_住院病历 = 0
    tab_护理病历 = 1
    tab_护理记录 = 2
    tab_知情文件 = 3
    tab_疾病证明 = 4
    tab_检验报告 = 5
    tab_检查报告 = 6
    tab_住院证 = 7
    tab_其他报表 = 8
End Enum

Public Enum Enum_Inside_Program
    p电子病历管理 = 2250
    p新版住院病历 = 2252
    p新版门诊病历 = 2251
    p疾病报告填写 = 1249
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p电子病案查阅 = 1259
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    P新版护士站 = 1265
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    p病人入出 = 1132
    p住院记帐 = 1133
    p费用查询 = 1139
    p门诊分诊管理 = 1113
    p排队叫号虚拟模块 = 1160
    p抗菌用药审核 = 1266
    p手术审核管理 = 1267
    p电子病案审查 = 1560
    p输血审核管理 = 1268
    p手麻接口 = 2425
    p手术授权管理 = 1080
    p输液配置中心 = 1345
    p电子病案打印 = 1566
End Enum
'常量
Private Const M_CON_CATE As String = "首页正面,首页反面,首页附页一,首页附页二,住院医嘱,检验报告,检查报告,住院病历,护理病历,护理记录,知情文件,疾病证明,临床路径,住院证,其他报表"
'变量
Private mlngCount As Long
Private mbytSelect As Byte    '记录勾选个数
Private mintPatiCount As Integer   '勾选病人数目
Private mblnTag As Boolean    '用于标识光标是否定位某个分类项目
Private mlngPatiID As Long    '当前病人ID
Private mlngDeptId As Long    '出院科室ID
Private mlngPatiMainID As Long      '当前病人主页ID
Private mlngInNO As Long            '住院号

Private mstrPatiName As String
Private mstrCardKind As String
Private mbytRows As Byte            '用于标记分类行数
Private mblnLoad As Boolean
Private mbytType As Byte           '用于标记最近一次所选分类,便于预览定位
Private mstrPrivs As String
Private mblnLIS As String        '是否按照新版LIS
Private mstr检验报告打印 As String        '0-老版LIS报表或病历;1-新版LIS报表方式
Private mstr检验对应报表 As String
Private mstr检查对应报表 As String
Private mcolReport As Collection
'参数
Private mbln个性化 As Boolean         '是否启用个性化风格
Private mintMecStandard As Integer    '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准,3-湖南省标准
Private mstrPrintDocIDs As String '共享病历的子文档只打印一次


'对象
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec
Private mobjSquareCard As Object     '医疗卡结算部件
Private mrsMedRec As ADODB.Recordset
Private mobjRichEMR As Object       '新版电子病历预览打印对象

'事件
Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1

Private Sub InitFace()
'功能:
    Dim i As Long
    Dim intRet As Integer
    Dim arrTmp As Variant
    Dim imgTmp As Image
    Dim lblTmp As Label
    Dim lngW As Long, lngH As Long
    Dim lngPos As Long
    Dim strTmp As String, strSelect As String
    Dim blnPath As Boolean
    Dim rsTmp As ADODB.Recordset
    
    strTmp = M_CON_CATE
    '根据权限控制显示打印项目
    If mintMecStandard = 0 Or mintMecStandard = 3 Then
        strTmp = Replace(strTmp, "首页附页一,首页附页二,", "")
    End If
    
    If GetPrivFunc(glngSys, p临床路径应用) <> "" And mlngDeptId <> 0 Then
        blnPath = gclsPackage.GetHavePath(mlngDeptId)
        If Not blnPath Then
            strTmp = Replace(strTmp, ",临床路径", "")
        End If
    End If
    
    arrTmp = Split(strTmp, ",")
    mlngCount = UBound(arrTmp)
    imgAll.Tag = "F"
    mbytSelect = 0
    strSelect = "," & GetRegister(私有模块, "打印档案", "打印内容", "1,2,3,4,5,6,7,8,9,10") & ","
    '动态加载控件数组PicItem
    For i = Lin.LBound To Lin.UBound
        Set Lin(i).Container = Me
    Next
    imgAll.Tag = "F"
    Set imgAll.Picture = imgList.ListImages("unCheck").Picture
    For i = picItem.LBound To picItem.UBound
        If i > picItem.LBound Then
            Unload lblItem(i)
            Unload imgItem(i)
            Unload imgCHK(i)
            Unload picItem(i)
        End If
    Next
    For i = 0 To mlngCount
        If i = 0 Then
            lngW = 120
            lngH = fraSplit.Top + fraSplit.Height + 120
        Else
            Load picItem(i)
            Load imgCHK(i)
            Load imgItem(i)
            Load lblItem(i)
            Set picItem(i).Container = fraIn
            Set imgCHK(i).Container = picItem(i)
            Set imgItem(i).Container = picItem(i)
            Set lblItem(i).Container = picItem(i)
            lngW = lngW + 1380  '间隔60缇
        End If
        If lngW + 1380 > picCenter.Width Then
            lngW = 120: lngH = lngH + 1500
        End If
        '容器缺省处理
        picItem(i).Move lngW, lngH, 1320, 1320
        picItem(i).Visible = True
        picItem(i).BackColor = picCenter.BackColor
        picItem(i).BorderStyle = 0 '无边框
        picItem(i).Appearance = 0
        
        '左上角图标
        imgCHK(i).Visible = False
        imgCHK(i).Tag = "F"     '标记未选中
        imgCHK(i).Move 15, 15, 300, 300
        Set imgCHK(i).Picture = imgList.ListImages("unCheck").Picture
        
        '中心图标
        imgItem(i).Visible = True
        imgItem(i).Move 300, 300, 720, 720
        Set imgItem(i).Picture = imgList.ListImages(arrTmp(i)).Picture
        
        '底部文字标识
        lblItem(i).Visible = True
        lblItem(i).AutoSize = True
        lblItem(i).BackStyle = 0 '透明
        lblItem(i).Caption = arrTmp(i)
        '记录分页签下标
        picItem(i).Tag = GetTabIndex(arrTmp(i))

        lblItem(i).Tag = ReturnItemTag(arrTmp(i))
        If InStr(strSelect, "," & lblItem(i).Tag & ",") > 0 Then
            Call SetPicItemBG(4, CInt(i))
            mblnLoad = False
            Call imgCHK_Click(CInt(i))   '缺省选中
            mblnLoad = True
        End If
        
        If lblItem(i).Caption = "其他报表" Then
            imgItem(i).ToolTipText = "报表参数固定：【病人ID】数字型 【主页ID】数字型"
        End If
        
        If lblItem(i).Width > picItem(i).Width Then
            lngPos = 0
        Else
            lngPos = (picItem(i).Width - lblItem(i).Width) / 2
        End If
        lblItem(i).Move lngPos, 1050
    Next
End Sub

Private Sub cboDept_Click()
    If Not Me.Visible Then Exit Sub
    
    cboDept.Tag = cboDept.ItemData(cboDept.ListIndex)
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strDeptIDs As String
    
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboDept.Text <> "" Then
            Set rsTmp = GetDataToDepts(cboDept.Text)
            If Not rsTmp.EOF Then
                Call cbo.SeekIndex(cboDept, rsTmp!ID)
            Else
                cboDept.ListIndex = Val(cboDept.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
        End If
    End If
End Sub

Private Sub cboOutTime_Click()
    Dim datCurr As Date
    Dim intDateCount As Integer
    
    intDateCount = cboOutTime.ItemData(cboOutTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpBegin.Enabled = False
    dtpEnd.Enabled = False
   
    If intDateCount = -1 Then
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    ElseIf intDateCount = 0 Then
        dtpBegin.Value = Format(datCurr, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(datCurr, "yyyy-MM-dd 23:59:59")
    Else
        dtpEnd.Value = datCurr
        dtpBegin.Value = datCurr - intDateCount
    End If
  
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, p电子病案打印, Control.Parameter, mlngPatiID, mlngPatiMainID, 0)
            Call zlPlugInErrH(Err, "ExecuteFunc")
            Err.Clear: On Error GoTo 0
        End If
    End Select
End Sub

Private Sub cmdFind_Click()
    Call ReadPati(2)
End Sub

Private Sub cmdPreview_Click()
    Call FuncPrintPreview
End Sub

Private Sub cmdPrint_Click()
    Call FuncPrint
End Sub

Private Sub cmdSet_Click()
    Dim i As Long
    Dim objFrm As New frmParaSet
    Dim lngRow As Long
    
    Call objFrm.ShowMe(Me, glngSys, glngModul, mstrPrivs)
    '重新加载
    Call FuncLoadReport
    
    If mrsMedRec Is Nothing Then Exit Sub
    
    Call GetRsMedRec(mlngPatiID, mlngPatiMainID, mlngDeptId, mrsMedRec)  '报表数据及新版LIS数据受参数影响需重新加载
    
    '缺省定位第一个显示且选中的页签
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub.Item(i).Visible Then
            tbcSub.Item(i).Selected = True
            Call tbcSub_SelectedChanged(tbcSub.Item(i))
            Exit Sub
        End If
    Next
    
    '项目分类加载完再界面重置
    Call picMain_Resize
End Sub

Private Sub Form_Load()
    Dim dteTime As Date
    Dim strPrinterName As String
    Dim intCount As Integer
    Dim objMenu As CommandBarPopup
    
    mblnLoad = False
    mintPatiCount = 0
    
    '读取参数
    '病案首页标准
    mintMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    mbln个性化 = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    mstr检查对应报表 = zlDatabase.GetPara("检查对应报表", glngSys, p电子病案打印)
    mstr检验对应报表 = zlDatabase.GetPara("检验对应报表", glngSys, p电子病案打印)
    mstr检验报告打印 = zlDatabase.GetPara("检验报告打印", glngSys, p电子病案打印)
    mblnLIS = sys.IsSysSetUp(2500)
    
    '权限
    mstrPrivs = GetPrivFunc(glngSys, p电子病案打印)
    '医疗卡部件
    mstrCardKind = "住|住院号|0|0|0|0|0|0;就|就诊卡|0|0|8|0|0|0;姓|姓名|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    Err.Clear: On Error GoTo 0
    If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
        Set mobjSquareCard = Nothing
        MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
    End If
    If Not mobjSquareCard Is Nothing Then Call PatiIdentifyFind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISAudit")
    'RIS接口创建
    Call CreateXWHIS(True)
    
    If mblnLIS Then Call InitObjLis(True)
    '新版电子病历
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    '-----------------------------------------------------
    
     With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(tab_住院病历, "住院病历", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_护理病历, "护理病历", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_护理记录, "护理记录", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_知情文件, "知情文件", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_疾病证明, "疾病证明", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_检验报告, "检验报告", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_检查报告, "检查报告", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_住院证, "住院证", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_其他报表, "其他报表", picItemInfo.hWnd, 0).Visible = False
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    

    Call InitVSTable
    Call ClearPatiInfo
    Call InitReportColumn
    '科室加载
    Call InitDepts
    '加载出院日期
    Call InitOutTime
    Call InitFace
    '加载打印设备
    strPrinterName = GetRegister(私有模块, "打印档案", "打印机", Printer.DeviceName)
    
    With cboPrinterName
        .Clear
        For intCount = 0 To Printers.count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = strPrinterName Then .ListIndex = intCount
        Next
    End With
    
    Call zlControl.CboSetWidth(cboPrinterName.hWnd, 3000)
    Call SetItemInfoTab
    
    '隐藏,此处只用做容器接收报表内容
    cbsMain.ActiveMenuBar.Visible = False
    Call FuncLoadReport
    
    '外挂菜单
    lblPlugIn.Visible = False
    If CreatePlugInOK(p电子病案打印) Then
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
        Set cbsMain.Icons = zlCommFun.GetPubIcons
    
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能(&K)", 0, False)
        objMenu.ID = conMenu_Tool_PlugIn
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    If mbln个性化 Then
        picShow.Tag = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Caption, "范围查找", 1)
    Else
        picShow.Tag = "0"
    End If
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, True)
    
    mblnLoad = True
    
End Sub

Private Sub SetItemInfoTab(Optional ByVal intIndex As Integer)
    Dim i As Long
    Dim idex As Long
    Dim intNum As Integer
    Dim blnTmp As Boolean
        
    '统计并记录打印类别
    With tbcSub
        For i = picItem.LBound To picItem.UBound
            If lblItem(i).Caption = "住院病历" Then
                .Item(tab_住院病历).Visible = imgCHK(i).Tag = "T"
                .Item(tab_住院病历).Tag = i
            ElseIf lblItem(i).Caption = "护理病历" Then
                .Item(tab_护理病历).Visible = imgCHK(i).Tag = "T"
                .Item(tab_护理病历).Tag = i
            ElseIf lblItem(i).Caption = "护理记录" Then
                .Item(tab_护理记录).Visible = imgCHK(i).Tag = "T"
                .Item(tab_护理记录).Tag = i
            ElseIf lblItem(i).Caption = "知情文件" Then
                .Item(tab_知情文件).Visible = imgCHK(i).Tag = "T"
                .Item(tab_知情文件).Tag = i
            ElseIf lblItem(i).Caption = "疾病证明" Then
                .Item(tab_疾病证明).Visible = imgCHK(i).Tag = "T"
                .Item(tab_疾病证明).Tag = i
            ElseIf lblItem(i).Caption = "检验报告" Then
                .Item(tab_检验报告).Visible = imgCHK(i).Tag = "T"
                .Item(tab_检验报告).Tag = i
            ElseIf lblItem(i).Caption = "检查报告" Then
                .Item(tab_检查报告).Visible = imgCHK(i).Tag = "T"
                .Item(tab_检查报告).Tag = i
            ElseIf lblItem(i).Caption = "住院证" Then
               .Item(tab_住院证).Visible = imgCHK(i).Tag = "T"
               .Item(tab_住院证).Tag = i
            ElseIf lblItem(i).Caption = "其他报表" Then
               .Item(tab_其他报表).Visible = imgCHK(i).Tag = "T"
               .Item(tab_其他报表).Tag = i
            End If
        Next
        
        '清空选择项目
        If intIndex > 0 Then
            If InStr(",住院病历,护理病历,护理记录,知情文件,疾病证明,检验报告,检查报告,住院证,其他报表,", "," & lblItem(intIndex).Caption & ",") > 0 And imgCHK(intIndex).Tag = "F" And Not mrsMedRec Is Nothing Then
                Select Case lblItem(intIndex).Caption
                Case "住院病历"
                    mrsMedRec.Filter = "上级ID='R2'"
                Case "护理病历"
                    mrsMedRec.Filter = "上级ID='R3'"
                Case "护理记录"
                    mrsMedRec.Filter = "上级ID='R4'"
                Case "知情文件"
                    mrsMedRec.Filter = "上级ID='R8'"
                Case "疾病证明"
                    mrsMedRec.Filter = "上级ID='R7'"
                Case "检验报告"
                    mrsMedRec.Filter = "上级ID='R6' And EPRId ='E' "
                Case "检查报告"
                    mrsMedRec.Filter = "上级ID='R6' And EPRId ='D' "
                Case "住院证"
                    mrsMedRec.Filter = "上级ID ='R10'"
                Case "其他报表"
                    mrsMedRec.Filter = "上级ID ='R11'"
                End Select
                If mrsMedRec.RecordCount > 0 Then
                    Do While Not mrsMedRec.EOF
                        mrsMedRec!是否选择 = 0  '缺省不选择
                        mrsMedRec.MoveNext
                    Loop
                End If
            End If
        End If
        
        For i = 0 To .ItemCount - 1
            If imgCHK(Val(.Item(i).Tag)).Tag = "T" Then
                If blnTmp = False Then blnTmp = True
                intNum = intNum + 1
            End If
        Next
        .Visible = blnTmp
        If intNum = 0 Or intNum = 1 Then picMain_Resize     '加载第一个页签或隐藏所有页签时重置界面
        
        '缺省选中当前勾选项目;如果是取消则缺省勾选Tab第一个页签
        If .Visible Then
            idex = CLng(picItem(intIndex).Tag)
            If idex >= 0 Then
                If .Item(idex).Visible Then
                    .Item(idex).Selected = True
                Else
                    For i = 0 To .ItemCount - 1
                        If .Item(i).Visible Then .Item(i).Selected = True: Exit Sub
                    Next
                End If
                Call LoadItemInfo
            End If
        End If
        
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Not mblnLoad Then Exit Sub
    picPati.Move 0, 0, (Me.ScaleWidth / 10) * 3, Me.ScaleHeight - stbThis.Height
    fraLine.Move picPati.Left + picPati.Width, 0, 45, Me.ScaleHeight - stbThis.Height
    PicMain.Move fraLine.Left + fraLine.Width, 0, Me.ScaleWidth - fraLine.Width - fraLine.Left, Me.ScaleHeight - stbThis.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbln个性化 Then SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Caption, "范围查找", picShow.Tag
    Call SaveWinState(Me, App.ProductName)
    Set mclsInOutMedRec = Nothing
    Set mrsMedRec = Nothing
    Set mcolReport = Nothing
    Set mobjRichEMR = Nothing
    If Not mclsDockAduits Is Nothing Then Set mclsDockAduits = Nothing
    If Not mobjSquareCard Is Nothing Then Set mobjSquareCard = Nothing
End Sub

Private Sub FraLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraLine.Left + x < (Me.ScaleWidth / 10) * 1 Or fraLine.Left + x > (Me.ScaleWidth / 10) * 9 Or Abs(x) < 100 Then Exit Sub
        fraLine.Left = fraLine.Left + x
        picPati.Width = picPati.Width + x
        PicMain.Left = PicMain.Left + x
        PicMain.Width = PicMain.Width - x
    End If
End Sub

Private Sub imgItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tbcSub.Visible Then
            If Val(picItem(Index).Tag) >= 0 Then
                If tbcSub.Item(picItem(Index).Tag).Visible Then
                    tbcSub.Item(picItem(Index).Tag).Selected = True
                End If
            End If
        End If
        Call FuncSetFocus(Index)
    End If
End Sub

Private Sub lblNote_Click()
    Call picShow_Click
End Sub

Private Sub lblPlugIn_Click()
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_Tool_PlugIn)
    If Not objPopup Is Nothing Then
        objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Sub PatiIdentifyFind_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strName As String
    Dim vRect As RECT
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    strSQL = ""
    strName = Trim(PatiIdentifyFind.Text)
    If objCard.名称 Like "*姓*名*" And blnCard = False And strName <> "" And InStr("-*+/", Left(strName, 1)) = 0 Then
        strSQL = "Select 1 As 排序id, a.病人id As ID, b.主页id, b.住院号, NVL(b.姓名,a.姓名) as 姓名, NVL(b.性别,a.性别) as 性别, NVL(b.年龄,a.年龄) as 年龄, a.身份证号, b.入院日期, b.出院日期, a.病人类型, a.住院次数" & vbNewLine & _
                "From 病人信息 A, 病案主页 B" & vbNewLine & _
                "Where a.病人id = b.病人id And a.姓名 Like [1] And b.出院日期 Is Not Null" & vbNewLine & _
                "Order By 排序id, 姓名, 入院日期 Desc"
    ElseIf (objCard.名称 = "住院号" Or Left(strName, 1) = "+") And IsNumeric(Mid(strName, 2)) And blnCard = False Then
        strSQL = "Select *" & vbNewLine & _
                "From (Select 1 As 排序id, a.病人id As ID, b.主页id, b.住院号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) As 年龄," & vbNewLine & _
                "              a.身份证号, b.入院日期, b.出院日期, a.病人类型, a.住院次数" & vbNewLine & _
                "       From 病人信息 A, 病案主页 B" & vbNewLine & _
                "       Where a.病人id = b.病人id And b.住院号 = [2] And b.出院日期 Is Not Null" & vbNewLine & _
                "       Order By 入院日期 Desc) A" & vbNewLine & _
                "Where Rownum < 2"

    End If
    
    If strSQL <> "" Then
        vRect = zlControl.GetControlRect(PatiIdentifyFind.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, PatiIdentifyFind.Height, blnCancel, False, True, strName & "%", strName)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '以病人ID读取
                mlngPatiID = NVL(rsTmp!ID)
                Call ReadPati(1)
                blnCancel = True: Exit Sub
            End If
        Else '取消选择
            If blnCancel = False Then
                MsgBox "没有找到符合条件的病人！", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If

    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgAll_Click()
    If imgAll.Tag = "T" Then
        imgAll.Tag = "F"
        Set imgAll.Picture = imgList.ListImages("unCheck").Picture
        Call SetPicItemBG(1)
        Call FuncSetFocus(0)
    Else
        imgAll.Tag = "T"
        Set imgAll.Picture = imgList.ListImages("Check").Picture
        Call SetPicItemBG(2)
    End If
End Sub

Private Sub imgCHK_Click(Index As Integer)
    If imgCHK(Index).Tag = "T" Then
        Set imgCHK(Index).Picture = imgList.ListImages("unCheck").Picture
        imgCHK(Index).Tag = "F"
        mbytSelect = mbytSelect - 1
        If mbytSelect = mlngCount And imgAll.Tag = "T" Then
            imgAll.Tag = "F"
            Set imgAll.Picture = imgList.ListImages("unCheck").Picture
        End If
    Else
        Set imgCHK(Index).Picture = imgList.ListImages("CheckFill").Picture
        imgCHK(Index).Tag = "T"
        mbytSelect = mbytSelect + 1
        If mbytSelect = mlngCount + 1 And imgAll.Tag = "F" Then
            imgAll.Tag = "T"
            Set imgAll.Picture = imgList.ListImages("Check").Picture
        End If
    End If
    If mblnLoad Then Call SetItemInfoTab(Index)
    Call FuncSetFocus(Index)
End Sub

Private Sub imgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetPicItemBG(3, Index)
    Call SetPicItemBG(4, Index)
End Sub

Private Sub PatiIdentifyFind_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    blnCancel = False
    If objHisPati Is Nothing Then blnCancel = True
    If blnCancel = False Then
        If objHisPati.病人ID = 0 Then blnCancel = True
    End If
    
    If blnCancel Then
        MsgBox "没有找到符合条件的病人！", vbInformation, gstrSysName
        Exit Sub
    End If

    mlngPatiID = objHisPati.病人ID
    
    Call ReadPati(1)
End Sub

Private Sub picCenter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'功能:取消选中项目
    Call SetPicItemBG(3)
End Sub

Private Sub picCenter_Resize()
    Dim i As Long
    Dim lngW As Long, lngH As Long
    Dim blnScor As Boolean
    
    On Error Resume Next
    fraSplit.Move 120, 420, picCenter.Width - 240, 45
    imgAll.Move 120, 60, 300, 300
    fraCent.Move 120, fraSplit.Top + fraSplit.Height, picCenter.Width - 135, picCenter.Height - (fraSplit.Top + fraSplit.Height) - 15
    fraIn.Move 0, 0, fraCent.Width - 255, fraCent.Height
    
    '边线颜色重置
    lineT.X1 = 0: lineT.Y1 = 0
    lineT.X2 = picCenter.Width: lineT.Y2 = 0
    lineT.BorderColor = &H80000010
    
    LineB.X1 = 0: LineB.Y1 = picCenter.Height - 15
    LineB.X2 = picCenter.Width: LineB.Y2 = picCenter.Height - 15
    LineB.BorderColor = lineT.BorderColor
    
    LineL.X1 = 0: LineL.Y1 = 0
    LineL.X2 = 0: LineL.Y2 = picCenter.Height
    LineL.BorderColor = lineT.BorderColor
    
    LineR.X1 = picCenter.Width - 15: LineR.Y1 = 0
    LineR.X2 = picCenter.Width - 15: LineR.Y2 = picCenter.Height
    LineR.BorderColor = lineT.BorderColor
    '重新排版
    mbytRows = 1
    For i = picItem.LBound To picItem.UBound
        If i = 0 Then
            lngW = 60
            lngH = 60
        Else
            lngW = lngW + 1380  '间隔120缇
        End If
        If lngW + 1380 > fraIn.Width Then
            lngW = 60: lngH = lngH + 1500
            mbytRows = mbytRows + 1
            '超过边界显示滚动条
            If lngH + picItem(i).Height + 60 > fraIn.Height Then
                fraIn.Height = fraIn.Height + picItem(i).Height + 120
                blnScor = True
            End If
        End If
        '容器缺省处理
        picItem(i).Move lngW, lngH, 1320, 1320
    Next
    vsc.Visible = blnScor
    If blnScor Then
        vsc.Max = mbytRows
        vsc.Move fraCent.Width - vsc.Width - 15, 0, 255, fraCent.Height
    End If
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call imgItem_MouseDown(Index, Button, Shift, x, y)
    End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetPicItemBG(3, Index)
    Call SetPicItemBG(4, Index)
End Sub

Private Sub SetPicItemBG(ByVal bytFunc As Byte, Optional ByVal Index As Integer = -1)
'功能:鼠标移入Pic高亮,移出显示背景色
'参数:
'   bytFunc-1 清除所有 ;2-选中所有;3-光标未选中任何项目,清空高亮状态;4-光标定位某个项目,背景色高亮
'   Index- bytFunc=4时传人
    Dim i As Integer
    
    On Error Resume Next
    If bytFunc = 1 Then
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "T" Then
                picItem(i).BackColor = picCenter.BackColor
                imgCHK(i).Visible = False
                Call imgCHK_Click(i)
            End If
        Next
        mbytSelect = 0
    ElseIf bytFunc = 2 Then
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "F" Then
                picItem(i).BackColor = COLOR_HIGH
                imgCHK(i).Visible = True
                Call imgCHK_Click(i)
            End If
        Next
        mbytSelect = mlngCount + 1
    ElseIf bytFunc = 3 Then
        '光标未定位在项目
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "F" And imgCHK(i).Visible = True Then
                If i <> Index Then
                    imgCHK(i).Visible = False
                    picItem(i).BackColor = picCenter.BackColor
                End If
            End If
        Next
        mblnTag = False
    ElseIf bytFunc = 4 And Index <> -1 Then
        If imgCHK(Index).Visible = False Then
            imgCHK(Index).Visible = True
            picItem(Index).BackColor = COLOR_HIGH
        End If
        mblnTag = True
    End If
End Sub

Private Sub picItemInfo_Resize()
    On Error Resume Next
    vsItemInfo.Move 0, 0, picItemInfo.Width, picItemInfo.Height
End Sub

Private Sub picMain_Resize()
    Dim lngW As Long
    On Error Resume Next
    
    If Not mblnLoad Then Exit Sub
    lngW = PicMain.Width - 120
    fraPati.Move 60, 60, lngW, 1575
    If tbcSub.Visible Then
        picCenter.Move 60, fraPati.Top + fraPati.Height + 120, lngW, 3650
        picCenter.Height = IIf(mbytRows > 1, 3560, 2200)
        tbcSub.Move 60, picCenter.Top + picCenter.Height, picCenter.Width, PicMain.Height - picCenter.Height - fraPati.Height - 200
    Else
        picCenter.Move 60, fraPati.Top + fraPati.Height + 120, lngW, PicMain.Height - fraPati.Top - fraPati.Height - 120
    End If
End Sub

Private Sub picPati_Resize()
    Dim lngW As Long
    Dim lngPos As Long
    
    On Error Resume Next
    If Not mblnLoad Then Exit Sub
    lngW = picPati.ScaleWidth - 240
    lngW = IIf(lngW < 4000, 4000, lngW)
    picShow.Move 120, 60, lngW, 270
    
    If picShow.Tag = "0" Then
        lblNote.Caption = "隐藏范围查找"
        Set picUpOrDown.Picture = imgList.ListImages("up").Picture
        fraScope.Visible = True
        fraScope.Move 120, picShow.Top + picShow.Height + 60, lngW, 1965
        fraFind.Move 120, fraScope.Top + fraScope.Height + 120, lngW, 975
    Else
        lblNote.Caption = "显示范围查找"
        Set picUpOrDown.Picture = imgList.ListImages("down").Picture
        fraScope.Visible = False
        fraFind.Move 120, picShow.Top + picShow.Height + 60, lngW, 975
    End If
    
    lngPos = fraFind.Top + fraFind.Height + 120
    rptPati.Move 120, lngPos, picPati.ScaleWidth - 240, picPati.Height - lngPos - 1335
    picPrint.Move 120, rptPati.Top + rptPati.Height, lngW, 1335
    
    lngW = fraScope.Width - 120 - 960
    lngW = IIf(lngW < 2700, 2700, lngW)
    
    dtpBegin.Move 960, 1080, 2175, 300
    dtpEnd.Move 960, dtpBegin.Top + dtpBegin.Height + 120, 2175, 300
    cboDept.Left = 960: cboDept.Top = 360
    cboDept.Width = lngW: cboDept.Height = 300
    cboOutTime.Left = 960: cboOutTime.Top = 720
    cboOutTime.Width = lngW: cboOutTime.Height = 300
    cmdFind.Move fraScope.Width - cmdFind.Width - 140, dtpBegin.Top + dtpBegin.Height + 120, 600, 300
 
    PatiIdentifyFind.Move 120, 360, fraFind.Width - 240, 300
End Sub

Private Sub picPrint_Resize()
    On Error Resume Next
    lblPrint.Move 120, 120, 720, 180
    cboPrinterName.Move lblPrint.Left + lblPrint.Width + 60, 60, picPrint.Width - (lblPrint.Left + lblPrint.Width + 180)
    cmdSet.Move 60, cboPrinterName.Top + cboPrinterName.Height + 200
    cmdPreView.Move picPrint.ScaleWidth - cmdPrint.Width * 2 - 240, cboPrinterName.Top + cboPrinterName.Height + 200
    cmdPrint.Move picPrint.ScaleWidth - cmdPrint.Width - 120, cboPrinterName.Top + cboPrinterName.Height + 200
    lblPlugIn.Move 140, cmdSet.Top + cmdSet.Height + 200
End Sub

Private Sub picShow_Click()
    If picShow.Tag = "1" Then
        picShow.Tag = "0"
    Else
        picShow.Tag = "1"
    End If
    Call picPati_Resize
End Sub

Private Sub picShow_Resize()
    lblNote.Move picShow.Left, 0, lblNote.Width, 270
    With picUpOrDown
        .Width = 270
        .Height = 270
        .Left = picShow.Width - picUpOrDown.Width - 120
        .Top = 0
    End With
End Sub

Private Sub picUpOrDown_Click()
    Call picShow_Click
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Item.Checked = True Then
        mintPatiCount = mintPatiCount + 1
    Else
        mintPatiCount = mintPatiCount - 1
    End If
    If mintPatiCount = rptPati.Records.count Then
        rptPati.Columns(col_选择).Icon = imgPati.ListImages("Check").Index - 1
    Else
        rptPati.Columns(col_选择).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
End Sub

Private Sub rptPati_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim hitColumn As ReportColumn
    Dim lngHit As Long
    
    If Button = 1 Then
        Set hitColumn = rptPati.HitTest(x, y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.ItemIndex = col_选择 Then
                lngHit = rptPati.HitTest(x, y).ht
                If xtpHitTestHeader = lngHit Then
                    If rptPati.Records.count = 0 Then Exit Sub  '无数据时禁止切换
                    If hitColumn.Icon = imgPati.ListImages("Check").Index - 1 Then
                        hitColumn.Icon = imgPati.ListImages("UnCheck").Index - 1
                        SelectItems 2
                    Else
                        hitColumn.Icon = imgPati.ListImages("Check").Index - 1
                        SelectItems 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim hitColumn As ReportColumn
    Dim Item As ReportRecordItem
    Dim strTipInfo As String
    Dim vPos As PointAPI
    Dim lngHwnd As Long
    
    On Error Resume Next
    Set hitColumn = rptPati.HitTest(x, y).Column
    If Not hitColumn Is Nothing Then
        If hitColumn.Index = col_打印图标 Then
            Set Item = rptPati.HitTest(x, y).Item
            If Not Item Is Nothing Then
                If Item.Record(col_打印图标).Icon <> -1 Then
                    strTipInfo = Item.Record(col_打印记录).Value
                    If strTipInfo = "" Then '如果没有获取过，则立即获取并记录在列表中
                        strTipInfo = GetPrintLog(Item.Record(col_病人Id).Value, Item.Record(col_主页ID).Value) '提取打印记录
                        Item(col_打印记录).Value = strTipInfo
                    End If
                    GetCursorPos vPos
                    lngHwnd = WindowFromPoint(vPos.x, vPos.y)
                    Call zlCommFun.ShowTipInfo(lngHwnd, strTipInfo, True)
                End If
            End If
        Else
            Call zlCommFun.ShowTipInfo(lngHwnd, "")
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    Dim lngRow As Long
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not Me.Visible Then Exit Sub
    If rptPati.SelectedRows.count = 0 Then Exit Sub          '非正常情况
    With rptPati.SelectedRows(0)
    
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            If rptPati.Tag = Val(.Record(col_病人Id).Value & "") & "_" & Val(.Record(col_主页ID).Value & "") Then Exit Sub
            '病人照片
            If Not ReadPatPricture(Val(.Record(col_病人Id).Value & ""), imgPatient) Then
               Set imgPatient.Picture = imgList.ListImages("Patient").Picture
            End If
            lblShow(lbl_姓名).Caption = .Record(col_姓名).Value
            lblShow(lbl_姓名).FontBold = True
            lblShow(lbl_姓名).ForeColor = IIf(.Record(col_病人类型).Value = "普通病人" Or .Record(col_病人类型).Value = "", &H0&, vbRed)
            lblShow(lbl_年龄).Caption = .Record(col_年龄).Value
            lblShow(lbl_性别).Caption = .Record(col_性别).Value
            lblShow(lbl_身份证号).Caption = .Record(col_身份证号).Value
            lblShow(lbl_住院号).Caption = .Record(col_住院号).Value
            lblShow(lbl_出院科室).Caption = .Record(col_出院科室).Value
            lblShow(lbl_入院日期).Caption = .Record(col_入院日期).Value
            lblShow(lbl_出院日期).Caption = .Record(col_出院日期).Value
            lblShow(lbl_出生日期).Caption = .Record(Col_出生日期).Value
            lblShow(lbl_住院医师).Caption = .Record(coL_住院医师).Value
            lblShow(lbl_家庭地址).Caption = .Record(col_家庭地址).Value
            
            mlngPatiID = Val(.Record(col_病人Id).Value & "")
            mlngPatiMainID = Val(.Record(col_主页ID).Value & "")
            mlngDeptId = Val(.Record(col_出院科室ID).Value & "")
            mlngInNO = Val(.Record(col_住院号).Value & "")
            mstrPatiName = Val(.Record(col_姓名).Value & "")
            
            If lblPlugIn.Visible Then lblPlugIn.Enabled = mlngPatiID <> 0
                
            Call GetRsMedRec(mlngPatiID, mlngPatiMainID, mlngDeptId, mrsMedRec)
            
            rptPati.Tag = mlngPatiID & "_" & mlngPatiMainID
            
            '缺省定位第一个显示且选中的页签
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    tbcSub.Item(i).Selected = True
                    Call tbcSub_SelectedChanged(tbcSub.Item(i))
                    Exit Sub
                End If
            Next
            
            '项目分类加载完再界面重置
            Call picMain_Resize
        End If
    End With
End Sub

Private Sub GetRsMedRec(ByVal lngPatiID As Long, ByVal lngPatiMainID As Long, ByVal lngDeptId As Long, ByRef rsMedRec As ADODB.Recordset, Optional ByVal blnMorePati As Boolean)
    Dim i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    Dim strErr As String
    Dim arrTmp As Variant
    Dim arrAdvice As Variant
    
    Set rsMedRec = gclsPackage.GetCISStruct(lngPatiID, lngPatiMainID, lngDeptId, False, "住院证") '切换一次病人读取一次
    Set rsMedRec = zlDatabase.CopyNewRec(rsMedRec, False, "", Array("是否选择", adInteger, 2, Empty, "预览", adInteger, 2, Empty))
    '
    '追加其他报表
    For i = 1 To mcolReport.count
        rsMedRec.AddNew
        rsMedRec!ID = mcolReport(i)
        rsMedRec!上级ID = "R11"
        rsMedRec!名称 = Split(mcolReport(i), ",")(0)
        rsMedRec!参数 = Split(mcolReport(i), ",")(0) & ";" & Split(mcolReport(i), ",")(1) & ";" & Split(mcolReport(i), ",")(2)  '报表名称,系统号,报表编号
        rsMedRec.Update
        If i = mcolReport.count Then
            rsMedRec.MoveFirst
        End If
    Next

    If Not gobjLIS Is Nothing And Val(mstr检验报告打印) = 1 Then
        strTemp = gobjLIS.GetPatientAdvice(mlngPatiID, mlngPatiMainID, strErr)  '医嘱之间用","分割，标本之间用";"分割 8362586,8362588;8362590
        If strErr <> "" Then MsgBox "LIS部件获取医嘱ID失败：" & vbCrLf & strErr, vbInformation, Me.Caption
        If strTemp <> "" Then
            arrTmp = Split(strTemp, ";")
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrAdvice = Split(arrTmp(i), ",")
                strTemp = ""
                For j = LBound(arrAdvice) To UBound(arrAdvice)
                    If UBound(arrAdvice) = 0 Then Exit For
                    rsMedRec.Filter = "上级ID = 'R6' And EPRID='E' And ID LIKE '*," & arrAdvice(j) & ",*'"
                    If j = UBound(arrAdvice) Then
                        If Not rsMedRec.EOF Then rsMedRec!名称 = Mid(strTemp, 2) & "," & rsMedRec!名称
                        Exit For
                    Else
                        If Not rsMedRec.EOF Then strTemp = strTemp & "," & Split(rsMedRec!名称 & "", "【")(0)
                        rsMedRec.Delete
                    End If
                Next
            Next
            rsMedRec.Filter = ""
        End If
    End If
    
    Do While Not rsMedRec.EOF
        If InStr(",R1,R5,R9,", "," & rsMedRec!ID & ",") > 0 Then
            rsMedRec!是否选择 = 1  '缺省选择 病案首页,住院医嘱,临床路径
        Else
            If blnMorePati Then
                rsMedRec!是否选择 = 1  '批量时默认选择
            Else
                rsMedRec!是否选择 = 0  '缺省不选择
            End If
        End If
        rsMedRec!预览 = 0
        rsMedRec.MoveNext
    Loop
    
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gclsPackage.GetEmrLIST(lngPatiID, lngPatiMainID)
    If Not rsTmp Is Nothing Then
        If rsTmp.State = ADODB.adStateOpen Then
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Do Until rsTmp.EOF
                    rsMedRec.AddNew
                    rsMedRec!ID = rsTmp!ID
                    rsMedRec!上级ID = "R2"
                    rsMedRec!名称 = rsTmp!名称
                    rsMedRec!参数 = rsTmp!ID 'NVL(rsTmp!参数) '文档ID[|子文档ID]
                    If blnMorePati Then
                        rsMedRec!是否选择 = 1  '缺省选择
                    Else
                        rsMedRec!是否选择 = 0  '缺省不选择
                    End If
                    rsMedRec!预览 = 0
                    rsMedRec.Update
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    If rsMedRec.RecordCount > 0 Then
        rsMedRec.MoveFirst
    End If
End Sub
Private Sub LoadItemInfo()
'功能：加载明细打印内容
    Dim arrTmp As Variant
    Dim lngRow As Long
    Dim strSplit As String
    
    If mrsMedRec Is Nothing Then Exit Sub
    If tbcSub.Visible = False Then Exit Sub

    vsItemInfo.Rows = 1
    mrsMedRec.Filter = 0
    If mrsMedRec.RecordCount > 0 Then
        With tbcSub
            If .Selected.Caption = "住院病历" Then
                mrsMedRec.Filter = "上级ID='R2'"
                strSplit = "【"
            ElseIf .Selected.Caption = "护理病历" Then
                mrsMedRec.Filter = "上级ID='R3'"
                strSplit = "【"
            ElseIf .Selected.Caption = "护理记录" Then
                mrsMedRec.Filter = "上级ID='R4'"
                strSplit = "("
            ElseIf .Selected.Caption = "知情文件" Then
                mrsMedRec.Filter = "上级ID='R8'"
                strSplit = "【"
            ElseIf .Selected.Caption = "疾病证明" Then
                mrsMedRec.Filter = "上级ID='R7'"
                strSplit = "【"
            ElseIf .Selected.Caption = "检验报告" Then
                mrsMedRec.Filter = "上级ID='R6' And EPRID = 'E'"
                strSplit = "【"
            ElseIf .Selected.Caption = "检查报告" Then
                mrsMedRec.Filter = "上级ID='R6' And EPRID = 'D'"
                strSplit = "【"
            ElseIf .Selected.Caption = "住院证" Then
                mrsMedRec.Filter = "上级ID ='R10'"
                strSplit = "【"
            ElseIf .Selected.Caption = "其他报表" Then
                mrsMedRec.Filter = "上级ID ='R11'"
            End If
            If mrsMedRec.RecordCount > 0 Then
                mrsMedRec.MoveFirst
                With vsItemInfo
                    .Rows = 1
                    Do While Not mrsMedRec.EOF
                        .Rows = .Rows + 1
                        .Cell(flexcpData, .Rows - 1, 1) = mrsMedRec!ID & ""
                        arrTmp = Split(mrsMedRec!名称 & "", strSplit)
                        If UBound(arrTmp) = 1 Then
                            .TextMatrix(.Rows - 1, 1) = arrTmp(0)
                            .TextMatrix(.Rows - 1, 2) = strSplit & arrTmp(1)
                        Else
                            .TextMatrix(.Rows - 1, 1) = mrsMedRec!名称 & ""
                        End If
                        If NVL(mrsMedRec!是否选择, 1) = 1 Then
                            lngRow = lngRow + 1
                            .Cell(flexcpChecked, .Rows - 1, 0) = 1
                        End If
                        If .Rows - 1 = 1 Then
                            mrsMedRec!预览 = 1     '缺省预览第一个文件
                        End If
                        mrsMedRec.MoveNext
                    Loop
                    .Row = 1
                    If lngRow = .Rows - 1 Then
                        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                        .Cell(flexcpData, 0, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                        .Cell(flexcpData, 0, 0) = 0
                    End If
                End With
            Else
                Set vsItemInfo.Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                vsItemInfo.Cell(flexcpData, 0, 0) = 0
                vsItemInfo.Rows = 2
            End If
        End With
    End If
End Sub

Private Sub ReadPati(ByVal bytFunc As Byte)
'功能:读取病人信息
'参数:bytFunc =1 代表通过病人ID查询病人出院记录
'     bytFunc=2  代表通过范围查找病人出院记录
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngColor As Long
    
    On Error GoTo ErrH
    If bytFunc = 1 Then
        strSQL = "Select distinct a.病人id, a.主页id,a.编目日期,a.住院号, NVl(a.姓名,b.姓名) as 姓名, NVL(a.性别,b.性别) as 性别, a.年龄,a.家庭地址, b.身份证号,b.出生日期, a.入院日期, a.出院日期,a.住院医师,a.出院科室id,a.病人类型, c.名称 as 出院科室,Decode(D.病人ID,NULL,0,1) as 是否打印 " & vbNewLine & _
            "From 病案主页 A, 病人信息 B, 部门表 C,病案打印记录 D" & vbNewLine & _
            "Where a.病人id = b.病人id And a.出院科室id = c.Id And A.病人ID=D.病人ID(+) And A.主页ID=D.主页ID(+) And a.病人id = [1] and a.出院日期 is Not NULL "
    ElseIf bytFunc = 2 Then
        strSQL = "Select distinct a.病人id, a.主页id,a.编目日期,a.住院号,NVL(a.姓名,b.姓名) as 姓名, NVL(a.性别,a.性别) as 性别, a.年龄,a.家庭地址, b.身份证号,b.出生日期,a.入院日期, a.出院日期,a.住院医师,a.出院科室id,a.病人类型, c.名称 as 出院科室,Decode(D.病人ID,NULL,0,1) as 是否打印 " & vbNewLine & _
            "From 病案主页 A, 病人信息 B, 部门表 C,病案打印记录 D" & vbNewLine & _
            "Where a.病人id = b.病人id And a.出院科室id = c.Id And A.病人ID=D.病人ID(+) And A.主页ID=D.主页ID(+) And a.出院科室id = [2] And a.出院日期 between [3] And [4]"
    End If

    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, Val(cboDept.Tag), CDate(dtpBegin.Value), CDate(dtpEnd.Value))
    '加载病人列表
    Call ClearPatiInfo
    rptPati.Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")    '图标
        objItem.HasCheckbox = True
        Set objItem = objRecord.AddItem("")  '图标
        If InStr(rsPati!性别 & "", "男") > 0 Then
            objItem.Icon = imgPati.ListImages("Boy").Index - 1
        ElseIf InStr(rsPati!性别 & "", "女") > 0 Then
            objItem.Icon = imgPati.ListImages("Girl").Index - 1
        End If
        Set objItem = objRecord.AddItem("")  '图标
        If Val(rsPati!是否打印 & "") = 1 Then
            objItem.Icon = imgPati.ListImages("print").Index - 1
        End If
        
        objRecord.AddItem IIf(NVL(rsPati!编目日期) <> "", "已编目", "未编目")
        objRecord.AddItem Format(rsPati!编目日期 & "", "YYYY-MM-dd")
        objRecord.AddItem rsPati!住院号 & ""
        objRecord.AddItem rsPati!姓名 & ""
        objRecord.AddItem rsPati!性别 & ""
        objRecord.AddItem rsPati!身份证号 & ""
        objRecord.AddItem Format(rsPati!出生日期 & "", "YYYY-MM-DD")
        objRecord.AddItem Format(rsPati!入院日期 & "", "YYYY-MM-DD")
        objRecord.AddItem Format(rsPati!出院日期 & "", "YYYY-MM-DD")
        objRecord.AddItem rsPati!出院科室 & ""
        objRecord.AddItem rsPati!住院医师 & ""
        objRecord.AddItem rsPati!家庭地址 & ""
        objRecord.AddItem rsPati!年龄 & ""
        '隐藏列
        objRecord.AddItem rsPati!病人类型 & ""
        objRecord.AddItem CLng(rsPati!病人ID)
        objRecord.AddItem NVL(rsPati!主页ID)
        objRecord.AddItem rsPati!出院科室ID & ""
        
         '显示病人颜色
        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
        objRecord.Item(col_姓名).ForeColor = lngColor

        rsPati.MoveNext
    Next
    rptPati.Populate
    '加载病人列表结束
    If rptPati.Records.count > 0 Then
        rptPati.Rows(0).Selected = True
        rptPati.SetFocus
        Call rptPati_SelectionChanged
    Else
        '重新初始化界面
        Set mrsMedRec = Nothing
        Call InitFace
        Call SetItemInfoTab
        vsItemInfo.Rows = 1
        vsItemInfo.Rows = 2
        Call picMain_Resize
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        Set objCol = .Columns.Add(col_选择, "", 20, False)
            objCol.Icon = imgPati.ListImages("UnCheck").Index - 1
            objCol.EditOptions.AllowEdit = True
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_图标, "", 20, False)  '图标
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_打印图标, "", 20, False)  'col_打印图标
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_是否编目, "是否编目", 60, True)
        Set objCol = .Columns.Add(col_编目日期, "编目日期", 80, True)
        Set objCol = .Columns.Add(col_住院号, "住院号", 80, True)
        Set objCol = .Columns.Add(col_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(col_性别, "性别", 45, True)
        Set objCol = .Columns.Add(col_身份证号, "身份证号", 150, True)
        Set objCol = .Columns.Add(Col_出生日期, "出生日期", 80, True)
        Set objCol = .Columns.Add(col_入院日期, "入院日期", 80, True)
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 80, True)
        Set objCol = .Columns.Add(col_出院科室, "出院科室", 90, True)
        Set objCol = .Columns.Add(coL_住院医师, "住院医师", 80, True)
        Set objCol = .Columns.Add(col_家庭地址, "地址", 150, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 45, True)
        
        '隐藏列
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_出院科室ID, "出院科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_打印记录, "打印记录", 0, False): objCol.Visible = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_入院日期)
        .SortOrder(0).SortAscending = True
    End With
    
End Sub

Private Sub ClearPatiInfo()
'功能:清除病人信息显示框
    Dim i As Long
    mlngPatiID = 0
    mlngPatiMainID = 0
    mlngDeptId = 0
    rptPati.Tag = ""
    If lblPlugIn.Visible Then lblPlugIn.Enabled = False
    
    For i = lblShow.LBound To lblShow.UBound
        lblShow(i).Caption = ""
    Next
    Set imgPatient.Picture = imgList.ListImages("Patient").Picture
    
    If Me.Visible Then
        mintPatiCount = 0
        rptPati.Columns(col_选择).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
End Sub

Private Sub InitOutTime()
'功能：初始化出院日期
    cboOutTime.Clear
    With cboOutTime
        .AddItem "今天"
        .ItemData(.NewIndex) = 0
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "15天内"
        .ItemData(.NewIndex) = 15
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
        
        .ListIndex = 0
    End With
End Sub

Private Function InitDepts(Optional ByVal strIn As String) As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strDeptIDs As String, lngPreDept As Long
    
    cboDept.Clear
    On Error GoTo ErrH
    

    Set rsTmp = GetDataToDepts
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call cbo.SetIndex(cboDept.hWnd, 0)
        cboDept.Tag = cboDept.ItemData(0)
    End If
    
    InitDepts = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDataToDepts(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能：获取科室病区列表数据记录集
'参数：strIn 过滤条件
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strLike As String
    
    If strIn <> "" Then blnYN = True
    strSQL = "Select Distinct a.Id, a.编码, a.名称" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where b.部门id = a.Id And b.工作性质 = '临床' And" & vbNewLine & _
            "      b.服务对象 In (2, 3) And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)" & vbNewLine & _
            IIf(blnYN, " And (A.编码 Like [1] Or A.简码 Like [2] Or A.名称 Like [2])", "") & _
            "Order By a.编码"

       
    On Error GoTo ErrH
    If blnYN Then
        strLike = IIf(gstrMatchMethod = "0", "%", "")
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strIn) & "%", strLike & UCase(strIn) & "%")
    Else
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
End Function

Private Sub FuncPrint()
'功能:根据选择项目打印电子病案
'   首页正面,首页反面,首页附页一,首页附页二,住院医嘱,检验报告,检查报告,住院病历,护理病历,护理记录,知情文件,疾病证明,临床路径
    Dim i As Long
    Dim strRegRange As String
    Dim strRange As String
    Dim strPrinterName As String
    
    '统计并记录打印类别
    For i = imgCHK.LBound To imgCHK.UBound
        If imgCHK(i).Tag = "T" Then
            strRegRange = strRegRange & "," & lblItem(i).Tag
            If InStr(",5,52,53,54,", "," & lblItem(i).Tag & ",") > 0 Then '首页的正反面，类型都是5
                If InStr(strRange, "R5") = 0 Then '没加
                    strRange = strRange & ",R5"
                End If
            ElseIf lblItem(i).Tag = 6 Then
                If InStr(strRange, "R6") = 0 Then '没加
                    strRange = strRange & ",R6"
                End If
            Else
                strRange = strRange & ",R" & lblItem(i).Tag
            End If
        End If
    Next
    
    If strRange <> "" Then
        strRange = strRange & ","
        strRegRange = Mid(strRegRange, 2)
    Else
        MsgBox "请选择需要输出的档案！", vbInformation, gstrSysName
        Exit Sub
    End If
    strPrinterName = cboPrinterName.Text
    
    If strPrinterName = "" Then
        MsgBox "请先选择输出设备，再进行打印操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    Call SetRegister(私有模块, "打印档案", "打印内容", strRegRange)
    Call SetRegister(私有模块, "打印档案", "打印机", strPrinterName)
    Call PrintDocument(strRegRange, strRange, strPrinterName)
    
End Sub

Private Sub PrintDocument(ByVal strRegRange As String, ByVal strRange As String, ByVal strPrinterName As String)
    Dim i As Integer, lngNo As Long
    Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
    Dim varParam As Variant, strReportNO As String, blnNewTends As Boolean, intSel As Integer, strEprName As String
    Dim lngInNo As Long, blnDataMove As Boolean, strName As String
    Dim lngSel As Long
    Dim lngPage As Long
    Dim strMsg As String
    Dim strReport As String
    Dim rsRet As New ADODB.Recordset
    
    On Error GoTo ErrHand

    '输出对象
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '调用打印
    strReport = mstr检验对应报表 & ";" & mstr检查对应报表 & ";" & IIf(mblnLIS, mstr检验报告打印, "0")
    If strReport = ";" Then strReport = ""
    
    With rptPati
        If mintPatiCount > 1 Then
            For i = 0 To .Records.count - 1
                If .Records(i).Item(col_选择).Checked = True Then
                    With .Records(i)
                        Call GetRsMedRec(CLng(.Item(col_病人Id).Value), CLng(.Item(col_主页ID).Value), CLng(.Item(col_出院科室ID).Value), rsRet, True)
                        Call gclsPackage.FuncPrintBatch(CLng(.Item(col_病人Id).Value), CLng(.Item(col_主页ID).Value), CLng(.Item(col_出院科室ID).Value), _
                            strRange, strRegRange, mclsDockAduits, clsPath, clsTendsNew, False, "", CStr(.Item(col_姓名).Value), CLng(.Item(col_住院号).Value), Me, _
                            lblInfo.Caption, False, strPrinterName, True, mstrPrintDocIDs, rsRet, lngPage, strReport, mobjRichEMR)
                        
                    End With
                End If
            Next
        ElseIf mlngPatiID <> 0 Then
            lngPage = 0: lngSel = FuncShowTipInfo()
            If lngSel = 0 Then
                MsgBox "打印失败：您未勾选任何文件。", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
            Call gclsPackage.FuncPrintBatch(mlngPatiID, mlngPatiMainID, mlngDeptId, strRange, strRegRange, mclsDockAduits, _
                    clsPath, clsTendsNew, False, "", mstrPatiName, mlngInNO, Me, lblInfo.Caption, False, strPrinterName, True, mstrPrintDocIDs, mrsMedRec, lngPage, strReport, mobjRichEMR)
            strMsg = "您选择了" & lngSel & "份文件。" & vbCrLf & " 一共打印了：" & lngPage & "份。"
            If strMsg <> "" Then MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
            
        End If
    End With
    lblInfo.Caption = ""
    Exit Sub
ErrHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    lblInfo.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select 打印次数 As 打印次, 打印内容, 打印人, 打印时间 From 病案打印记录 Where 病人id = [1] And 主页id = [2] Order By 打印时间, 打印序号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!打印人, 10) & Rpad(Format(rs!打印时间, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!打印内容, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("打印人", 10) & Rpad("打印时间", 20) & Rpad("打印内容", 40) & GetPrintLog
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncPrintPreview()
'功能:预览
'   首页正面,首页反面,首页附页一,首页附页二,住院医嘱,检验报告,检查报告,住院病历,护理病历,护理记录,知情文件,疾病证明,临床路径，住院证
    Dim i As Long
    Dim strRegRange As String
    Dim strRange As String
    Dim strPrinterName As String
    Dim lngTabIX As Long
    Dim strTabCaption As String
    
    '定位预览类别
    lngTabIX = CLng(picItem(mbytType).Tag)
    If lngTabIX >= 0 Then
        If Not (tbcSub.Visible And tbcSub.Item(lngTabIX).Selected) And InStr(",检验报告,检查报告,住院病历,护理病历,护理记录,知情文件,疾病证明,住院证,", lblItem.Item(mbytType).Caption) > 0 Then
            MsgBox "您未勾选【" & lblItem(mbytType).Caption & "】，不能预览。", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    strRegRange = strRegRange & "," & lblItem(mbytType).Tag
    If InStr(",5,52,53,54,", "," & lblItem(mbytType).Tag & ",") > 0 Then '首页的正反面，类型都是5
        If InStr(strRange, "R5") = 0 Then '没加
            strRange = strRange & ",R5"
        End If
    Else
        strRange = strRange & ",R" & lblItem(mbytType).Tag
        strTabCaption = lblItem(mbytType).Caption
    End If

    
    If strRange <> "" Then
        strRange = Replace(strRange, ",", "")

        strRegRange = Replace(strRegRange, ",", "")
    Else
        MsgBox "请选择需要输出的档案！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strPrinterName = cboPrinterName.Text
    If strPrinterName = "" Then
        MsgBox "请先选择输出设备，再进行打印操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call PrintDocumentView(strRegRange, strRange, strPrinterName, strTabCaption)
End Sub

Private Sub PrintDocumentView(ByVal strRegRange As String, ByVal strRange As String, ByVal strPrinterName As String, ByVal strTabCaption As String)
    Dim i As Integer, rs As New ADODB.Recordset, lngNo As Long
    Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
    Dim varParam As Variant, strReportNO As String, blnNewTends As Boolean, intSel As Integer, strEprName As String
    Dim strName As String, strMsg As String
    Dim blnMod As Boolean
    
    On Error GoTo ErrHand

    '输出对象
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '读取记录
    Set rs = mrsMedRec
    If InStr(",R5,R9,R1,", "," & strRange & ",") > 0 Then
        rs.Filter = "ID = '" & strRange & "'"
        If rs.RecordCount > 0 Then
            Select Case rs("ID").Value
            Case "R5"               '首页
                Select Case mintMecStandard
                Case 0 '卫生部标准
                    If Have部门性质(mlngDeptId, "中医科") Then
                        strReportNO = "ZL1_INSIDE_1261_4"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_1"
                    End If
                Case 1    '四川省标准
                    If Have部门性质(mlngDeptId, "中医科") Then
                        strReportNO = "ZL1_INSIDE_1261_6"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_5"
                    End If
                Case 2    '云南省标准
                    If Have部门性质(mlngDeptId, "中医科") Then
                        strReportNO = "ZL1_INSIDE_1261_8"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_7"
                    End If
                Case 3     '湖南省标准
                    If Have部门性质(mlngDeptId, "中医科") Then
                        strReportNO = "ZL1_INSIDE_1261_10"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_9"
                    End If
                Case Else '当期修改时未定义
                    If Have部门性质(mlngDeptId, "中医科") Then
                        strReportNO = "ZL1_INSIDE_1261_4"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_1"
                    End If
                End Select
          
                Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & strReportNO, "Printer", strPrinterName)
                If InStr("," & strRegRange & ",", ",5,") > 0 Then '正面
                    Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "ReportFormat=1", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",52,") > 0 Then '反面
                    Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "ReportFormat=2", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",53,") > 0 Then '附一
                    Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "ReportFormat=3", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",54,") > 0 Then '附二
                    Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "ReportFormat=4", 1)
                End If
    
            Case "R1"               '医嘱
                '先打印长嘱
                Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_1", "Printer", strPrinterName)
                Call gobjKernel.zlPrintAdvice(Me, mlngPatiID, mlngPatiMainID, 0, 0, strPrinterName, 1)
                '再打印临嘱
                Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_2", "Printer", strPrinterName)
                Call gobjKernel.zlPrintAdvice(Me, mlngPatiID, mlngPatiMainID, 0, 1, strPrinterName, 1)
                
            Case "R9"               '临床路径
                If Not CheckPatiPath() Then
                    MsgBox "当前病人没有临床路径表,不能预览。", vbOKOnly, Me.Caption
                    Exit Sub
                End If
                Call clsPath.zlRefreshReadOnly(mlngPatiID, mlngPatiMainID)
                Call clsPath.zlFuncPathTableOutPut(2, True, "", 0, 0, strPrinterName)  '预览
            End Select
        End If
    Else
        '子项目
        rs.Filter = "上级id = '" & strRange & "'" & " And 预览=1 "
        If rs.RecordCount > 0 Then
            If InStr(rs!ID, "R") = 0 And Len(rs!ID) >= 32 Then
                'EMR病历预览
                If Not mobjRichEMR Is Nothing Then
                    If InStr(rs!参数, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(rs!参数, "|")(0), Split(rs!参数, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(rs!参数, "")
                    End If
                    Call mobjRichEMR.zlPrintDoc(True)
                End If
            Else
                varParam = Split(rs("参数").Value, ";")
                Select Case rs("上级id").Value
                Case "R2"               '住院病历
                    rs.Filter = "ID = '" & strRange & "'"
                    strEprName = Split(rs("名称").Value, "【")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R3"               '护理病历
                    If InStr("," & mstrPrintDocIDs, "," & Val(varParam(0)) & ",") = 0 Then '本次没打过
                        strEprName = Split(rs("名称").Value, "【")(0)
                        Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                    End If
                Case "R4"               '护理记录
                    blnNewTends = Get新版护理(mlngPatiID, mlngPatiMainID)
                    If blnNewTends = False Then
                        varParam = Split(rs("参数").Value, ";")
                        If UBound(varParam) >= 1 Then
                            If Val(varParam(1)) = -1 Then '体温单
                                Call mclsDockAduits.zlRefreshTendBody(mlngPatiID, mlngPatiMainID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)))
                                Call mclsDockAduits.zlPrintDocument(1, 1, , strPrinterName)
                            Else '护理记录
                                Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), mlngPatiID, mlngPatiMainID, Val(Split(varParam(0), "_")(0)), CStr(varParam(2)), , Val(varParam(4)))
                                Call mclsDockAduits.zlPrintDocument(2, 1, , strPrinterName)
                            End If
                        End If
                    Else
                        varParam = Split(rs("参数").Value, ";")
                        If UBound(varParam) >= 1 Then
                            Select Case Val(varParam(1))
                                Case -1 '体温单
                                    intSel = 1
                                Case 1  '产程图
                                    intSel = 3
                                Case Else '记录单
                                    intSel = 2
                            End Select
                            Call clsTendsNew.zlPrintDocument(mlngPatiID, mlngPatiMainID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), intSel, strPrinterName, False)
                        End If
                    End If
                Case "R6"               '检验检查报告
                    If NVL(rs!Eprid, "") = "E" And mstr检验对应报表 <> "" Then
                        strReportNO = Split(mstr检验对应报表, ",")(2)
                        varParam = Split(rs("参数").Value, ";")  '第二个参数是医嘱ID
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "医嘱ID=" & varParam(1), 1)
                    ElseIf NVL(rs!Eprid, "") = "D" And mstr检查对应报表 <> "" Then
                        strReportNO = Split(mstr检查对应报表, ",")(2)
                        varParam = Split(rs("参数").Value, ";")  '第二个参数是医嘱ID
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, "医嘱ID=" & varParam(1), 1)
                    Else
                        strEprName = Split(rs("名称").Value, "【")(0)
                        If UBound(Split(strEprName, ">")) > 0 Then
                            strEprName = Split(strEprName, ">")(1)
                        End If
                        blnMod = False
                        If NVL(rs!Eprid, "") = "E" And mblnLIS And Val(mstr检验报告打印) = 1 Then
                            If InitObjLis(False) Then
                                blnMod = gobjLIS.PrintReport(Me, Val(varParam(1)), mlngPatiID, 1, strMsg)
                            End If
                        End If
                        If Not blnMod Then
                            If Val(varParam(3)) <> 0 Then
                                'RIS
                                If Not gobjXWHIS Is Nothing Then
                                    Call gobjXWHIS.ShowViewReport(Me.hWnd, Val(varParam(1)), True, Val(varParam(3)))
                                End If
                            ElseIf Val(varParam(0)) <> 0 Then
                                Call mclsDockAduits.zlPrintDocument(4, 1, Val(varParam(0)), strPrinterName)
                            Else
                                If objPacsDoc Is Nothing Then
                                    Set objPacsDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "新版PACS编辑器", False)
                                    Call objPacsDoc.InitInterface(gcnOracle, gstrDBUser)
                                End If
                                Call objPacsDoc.PrintReport(varParam(2), strPrinterName, True) 'True预览
                            End If
                        End If
                    End If
                Case "R7"               '疾病证明
                    strEprName = Split(rs("名称").Value, "【")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R8"               '知情文件
                    strEprName = Split(rs("名称").Value, "【")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R10"     '电子病案打印 住院证
                    strEprName = "ZLCISBILL" & Format(rs!Eprid, "00000") & "-1"
                    If UBound(varParam) >= 1 Then
                        Call ReportOpen(gcnOracle, glngSys, strEprName, Me, "NO=" & varParam(0), "性质=" & varParam(1), "医嘱ID=0", 1)
                    End If
                Case "R11"  '其他报表
                    If UBound(varParam) >= 1 Then
                        strReportNO = varParam(2)
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & mlngPatiID, "主页id=" & mlngPatiMainID, 1)
                    End If
                End Select
            End If
        Else
            Select Case strRange
            
            Case "R2"
                strMsg = "当前病人没有住院病历，不能预览。"
            Case "R3"
                strMsg = "当前病人没有护理病历，不能预览。"
            Case "R4"
                strMsg = "当前病人没有护理记录，不能预览。"
            Case "R6"
                strMsg = "当前病人没有" & strTabCaption & "，不能预览。"
            Case "R7"
                strMsg = "当前病人没有疾病证明，不能预览。"
            Case "R8"
                strMsg = "当前病人没有知情文件，不能预览。"
            Case "R10"
                strMsg = "当前病人没有住院证，不能预览。"
            Case "R11"
                strMsg = "当前病人没有其他报表，不能预览。"
            End Select
            If strMsg <> "" Then
                MsgBox strMsg, vbOKOnly, Me.Caption
            End If
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ReturnItemTag(ByVal strName As String) As Integer
'功能:返回特定标记
'1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-检验检查报告;7-疾病证明;8-知情文件;9-临床路径;10-住院证;11-其他报表
    Dim intRet As Integer
    
    Select Case strName
    
    Case "住院医嘱"
        intRet = 1
    Case "住院病历"
        intRet = 2
    Case "护理病历"
        intRet = 3
    Case "护理记录"
        intRet = 4
    Case "首页正面"
        intRet = 5
    Case "首页反面"
        intRet = 52
    Case "首页附页一"
        intRet = 53
    Case "首页附页二"
        intRet = 54
    Case "检验报告", "检查报告"
        intRet = 6
    Case "疾病证明"
        intRet = 7
    Case "知情文件"
        intRet = 8
    Case "临床路径"
        intRet = 9
    Case "住院证"
        intRet = 10
    Case "其他报表"
        intRet = 11
    End Select
    ReturnItemTag = intRet
End Function

Private Function GetTabIndex(ByVal strName As String) As Integer
    Dim intRet As Integer
    
    Select Case strName
    
    Case "住院病历"
        intRet = tab_住院病历
    Case "护理病历"
        intRet = tab_护理病历
    Case "护理记录"
        intRet = tab_护理记录
    Case "检验报告"
        intRet = tab_检验报告
    Case "检查报告"
        intRet = tab_检查报告
    Case "疾病证明"
        intRet = tab_疾病证明
    Case "知情文件"
        intRet = tab_知情文件
    Case "住院证"
        intRet = tab_住院证
    Case "其他报表"
        intRet = tab_其他报表
    Case Else
        intRet = -1
    End Select
    GetTabIndex = intRet
End Function
Private Function CheckPatiPath() As Boolean
'功能:检查当前病人是否存在临床路径表单(状态=2-正常结束;3-变异结束)
    Dim strSQL As String
    Dim rsPath As ADODB.Recordset
    
    On Error GoTo ErrH
    strSQL = "Select Count(1) as 记录数 From 病人临床路径 A Where a.病人id = [1] And a.主页id = [2] And a.状态 In (2, 3)"
    Set rsPath = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mlngPatiMainID)
    CheckPatiPath = Val(rsPath!记录数 & "") > 0
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectItems(ByVal bytFunc As Byte)
'参数:
'   bytFunc=1 全选,=2取消全选
    Dim i As Long
    
    With rptPati
        For i = 0 To .Records.count - 1
            If bytFunc = 1 Then
                .Records(i).Item(col_选择).Checked = True
            Else
                .Records.Record(i).Item(0).Checked = False
            End If
        Next
        mintPatiCount = IIf(bytFunc = 1, .Records.count, 0)
    End With
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call LoadItemInfo
    If tbcSub.Visible Then Call FuncSetFocus(Val(tbcSub.Selected.Tag))
End Sub

Private Sub vsc_Change()
    fraIn.Top = 0 - (fraIn.Height - fraCent.Height) * (vsc.Value / vsc.Max)
    '转移焦点
    picCenter.SetFocus
End Sub

Private Sub vsItemInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItemInfo
        If Col = 0 Then
            If mrsMedRec Is Nothing Then Exit Sub
            mrsMedRec.Filter = "ID='" & .Cell(flexcpData, Row, 1) & "'"
            If mrsMedRec.RecordCount > 0 Then
                mrsMedRec.MoveFirst
                mrsMedRec!是否选择 = IIf(.Cell(flexcpChecked, Row, Col) = 1, 1, 0)
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, Col) = flexUnchecked Then Exit For
                Next
                If i = .Rows Then
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                    .Cell(flexcpData, 0, 0) = 1
                Else
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                    .Cell(flexcpData, 0, 0) = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItemInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub tmrTime_Timer()
    Dim vPos As PointAPI
    Dim vRect As RECT
    
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "勾选了" & mintPatiCount & "个病人！")
    cmdPrint.Enabled = mlngPatiID <> 0
    cmdPreView.Enabled = mlngPatiID <> 0
    
    If mblnTag = False Then Exit Sub
    
    GetCursorPos vPos
    GetWindowRect picCenter.hWnd, vRect
    If Not (Between(vPos.x, vRect.Left, vRect.Right) And Between(vPos.y, vRect.Top, vRect.Bottom)) Then
        Call SetPicItemBG(3)
    End If
End Sub

Private Sub vsItemInfo_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    If mbytType <> CByte(tbcSub.Selected.Tag) Then Call FuncSetFocus(CLng(tbcSub.Selected.Tag))
    
    If mrsMedRec Is Nothing Then Exit Sub
    With vsItemInfo
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow = 0 And lngCol = 0 Then
            If .TextMatrix(.Rows - 1, 1) <> "" Then
                If Val(.Cell(flexcpData, 0, 0) & "") = 0 Then
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                    .Cell(flexcpData, 0, 0) = 1
                Else
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                    .Cell(flexcpData, 0, 0) = 0
                End If
                
                For i = .FixedRows To .Rows - 1
                    .Cell(flexcpChecked, i, 0) = IIf(Val(.Cell(flexcpData, 0, 0) & "") = 0, flexUnchecked, flexChecked)
                    mrsMedRec.Filter = "ID='" & .Cell(flexcpData, i, 1) & "'"
                    If mrsMedRec.RecordCount > 0 Then
                        mrsMedRec!是否选择 = IIf(.Cell(flexcpChecked, i, lngCol) = 1, 1, 0)
                    End If
                Next
            End If
        ElseIf lngRow >= 1 And lngRow <= .Rows - 1 Then
            mrsMedRec.Filter = ""
            Do While Not mrsMedRec.EOF
                mrsMedRec!预览 = 0 '清空所有预览项
                mrsMedRec.MoveNext
            Loop
            mrsMedRec.Filter = "ID='" & .Cell(flexcpData, lngRow, 1) & "'"
            If mrsMedRec.RecordCount > 0 Then mrsMedRec!预览 = 1
        End If
    End With
End Sub

Private Sub InitVSTable()

    With vsItemInfo
        .Cols = 3: .ColWidth(0) = 300
        .ColWidth(1) = 3000: .ColWidth(2) = 7500
        .FixedAlignment(1) = flexAlignCenterCenter
        .RowHeightMin = 300
        .Editable = flexEDKbd
        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, 0) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(0, 1) = "请勾选需要打印的内容。": .ColDataType(0) = flexDTBoolean
    End With
End Sub

Private Sub FuncSetFocus(ByVal bytIndex As Byte)
    Dim i As Long
    
    If Not mblnLoad Then Exit Sub
    
    For i = Lin.LBound To Lin.UBound
        Set Lin(i).Container = picItem(bytIndex)
        Lin(i).BorderColor = &HFF0000
        Lin(i).Visible = True
    Next
     
    With Lin(0)
        .X1 = 0: .X2 = picItem(i).Width
        .Y1 = 0: .Y2 = 0
    End With
    With Lin(1)
        .X1 = picItem(i).Width - 15: .X2 = picItem(i).Width - 15
        .Y1 = 0: .Y2 = picItem(i).Height
    End With
    With Lin(2)
        .X1 = 0: .X2 = picItem(i).Width - 15
        .Y1 = picItem(i).Height - 15: .Y2 = picItem(i).Height - 15
    End With
    With Lin(3)
        .X1 = 0: .X2 = 0
        .Y1 = 0: .Y2 = picItem(i).Height
    End With
    mbytType = bytIndex
End Sub

Private Function FuncShowTipInfo() As Long
    Dim lngCount As Long
    Dim i As Long
    Dim strRange As String
    Dim lngAll As Long
    Dim lngSub As Long
    Dim strMsg As String
    
    If mrsMedRec Is Nothing Then Exit Function
    If mrsMedRec.RecordCount = 0 Then Exit Function

    '统计勾选文件份数 首页正反面或附件算一份,医嘱清单（长期\临时）算一份,临床路径表单算一份
    For i = imgCHK.LBound To imgCHK.UBound
        If imgCHK(i).Tag = "T" Then
            If InStr(",5,52,53,54,", "," & lblItem(i).Tag & ",") > 0 Then '首页的正反面，类型都是5
                If InStr(strRange, "R5") = 0 Then '没加
                    strRange = strRange & ",R5"
                    lngCount = lngCount + 1
                End If
            ElseIf InStr(",1,9,", "," & lblItem(i).Tag & ",") > 0 Then
                lngCount = lngCount + 1
            Else
                strRange = strRange & ",R" & lblItem(i).Tag
            End If
        End If
    Next
    '统计单个文件数目
    mrsMedRec.Filter = "是否选择 =1"
    lngAll = mrsMedRec.RecordCount
    mrsMedRec.Filter = "上级ID=Null And 是否选择 =1"
    lngSub = mrsMedRec.RecordCount
    lngCount = lngCount + (lngAll - lngSub)
    
    FuncShowTipInfo = lngCount
    
End Function

Private Sub FuncLoadReport()
    Dim objControl As CommandBarControl
    Dim objPop As Object
    Dim strHide As String
    Dim i As Long
    
    strHide = ",ZL1_INSIDE_1254_1,ZL1_INSIDE_1254_2,ZL1_INSIDE_1261_1,ZL1_INSIDE_1261_4,ZL1_INSIDE_1261_5,ZL1_INSIDE_1261_6,ZL1_INSIDE_1261_7,ZL1_INSIDE_1261_8,ZL1_INSIDE_1261_9,ZL1_INSIDE_1261_10,"
    mstr检查对应报表 = zlDatabase.GetPara("检查对应报表", glngSys, p电子病案打印)
    mstr检验对应报表 = zlDatabase.GetPara("检验对应报表", glngSys, p电子病案打印)
    If mstr检查对应报表 <> "" Then strHide = strHide & "," & Split(mstr检查对应报表, ",")(2) & ","
    If mstr检验对应报表 <> "" Then strHide = strHide & "," & Split(mstr检验对应报表, ",")(2) & ","
    mstr检验报告打印 = zlDatabase.GetPara("检验报告打印", glngSys, p电子病案打印)
    '清空缓存
    Set mcolReport = New Collection
    For i = 1 To cbsMain.ActiveMenuBar.Controls.count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
                cbsMain.ActiveMenuBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    Call zlDatabase.ShowReportMenu(cbsMain, glngSys, p电子病案打印, mstrPrivs, strHide)
    
    For i = 1 To cbsMain.ActiveMenuBar.Controls.count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
            Set objControl = cbsMain.ActiveMenuBar.Controls.Item(i)
            Exit For
        End If
    Next
    
    If Not objControl Is Nothing Then
        With objControl.CommandBar.Controls
            For i = 1 To .count
                Set objPop = .Item(i)
                mcolReport.Add Split(objPop.Caption, "(&")(0) & "," & objPop.Parameter, "_" & i     '报表名称,系统号,报表编号
            Next
        End With
    End If
End Sub

Public Sub DefCommandPlugInPopup(ByRef objControls As CommandBarControls)
'功能：扩展功能弹出菜单
    Dim strFunc As String, strTmp As String
    Dim arrTmp As Variant
    Dim objControl As CommandBarControl
    
    Dim i As Long
    
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, p电子病案打印)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    If strFunc <> "" Then
        arrTmp = Split(strFunc, ",")
        strTmp = Replace(strFunc, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        If objControls.count = 0 Then
            For i = 0 To UBound(arrTmp)
                Set objControl = objControls.Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                objControl.IconId = conMenu_Tool_PlugIn_Item
                objControl.Parameter = arrTmp(i)
            Next
        End If
        lblPlugIn.Visible = True
        lblPlugIn.Enabled = False
    End If
End Sub


