VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmBaseItemList 
   BorderStyle     =   0  'None
   Caption         =   "检验仪器项目"
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5190
      Index           =   3
      Left            =   4050
      ScaleHeight     =   5160
      ScaleWidth      =   8160
      TabIndex        =   27
      Top             =   5070
      Width           =   8190
      Begin VB.PictureBox picSub3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   315
         ScaleHeight     =   2400
         ScaleWidth      =   7515
         TabIndex        =   76
         Top             =   2250
         Width           =   7515
         Begin VB.TextBox txtReference 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   3900
            MaxLength       =   13
            TabIndex        =   56
            Top             =   495
            Width           =   900
         End
         Begin VB.ComboBox cboSampleType 
            Height          =   300
            Left            =   915
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   105
            Width           =   2100
         End
         Begin VB.TextBox txtAge 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   5355
            MaxLength       =   3
            TabIndex        =   52
            Top             =   105
            Width           =   495
         End
         Begin VB.ComboBox cboSex 
            Height          =   300
            Left            =   3915
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   105
            Width           =   915
         End
         Begin VB.TextBox txtAge 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   6075
            MaxLength       =   3
            TabIndex        =   53
            Top             =   105
            Width           =   495
         End
         Begin VB.ComboBox cboAgeUnitl 
            Height          =   300
            Left            =   6600
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   105
            Width           =   750
         End
         Begin VB.TextBox txtReference 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   5145
            MaxLength       =   13
            TabIndex        =   57
            Top             =   495
            Width           =   900
         End
         Begin VB.ComboBox cboFeatures 
            Height          =   300
            Left            =   915
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   465
            Width           =   2100
         End
         Begin VB.ComboBox cboMachine3 
            Height          =   300
            Left            =   915
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   840
            Width           =   2100
         End
         Begin VB.TextBox txtReferenceShow 
            Height          =   300
            Left            =   3900
            MaxLength       =   50
            TabIndex        =   60
            Top             =   840
            Width           =   3450
         End
         Begin VB.CheckBox chkDefault 
            Alignment       =   1  'Right Justify
            Caption         =   "默认参考"
            Height          =   180
            Left            =   6210
            TabIndex        =   58
            Top             =   555
            Width           =   1140
         End
         Begin VB.TextBox txtAbnorma 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   3930
            MaxLength       =   12
            TabIndex        =   62
            Top             =   1590
            Width           =   1020
         End
         Begin VB.TextBox txtAbnorma 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1815
            MaxLength       =   12
            TabIndex        =   61
            Top             =   1590
            Width           =   1020
         End
         Begin VB.TextBox txtReview 
            Height          =   300
            Index           =   1
            Left            =   1815
            MaxLength       =   12
            TabIndex        =   63
            Top             =   1965
            Width           =   1020
         End
         Begin VB.TextBox txtReview 
            Height          =   300
            Index           =   0
            Left            =   3930
            MaxLength       =   12
            TabIndex        =   64
            Top             =   1980
            Width           =   1020
         End
         Begin VB.Label Label11 
            Caption         =   "医学审核条件"
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   90
            TabIndex        =   86
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label lblSampleType 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "标本种类"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   90
            TabIndex        =   85
            Top             =   165
            Width           =   720
         End
         Begin VB.Label lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "年龄      ～"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4980
            TabIndex        =   84
            Top             =   165
            Width           =   1080
         End
         Begin VB.Label lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3105
            TabIndex        =   83
            Top             =   165
            Width           =   360
         End
         Begin VB.Label lblReference 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "参考范围  　　      ～"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3105
            TabIndex        =   82
            Top             =   525
            Width           =   1980
         End
         Begin VB.Label lblFeatures 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "临床特征"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   90
            TabIndex        =   81
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lblMachine3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "对应仪器"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   90
            TabIndex        =   80
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lblReferenceShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "参考显示"
            Height          =   180
            Left            =   3105
            TabIndex        =   79
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lblReview 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "如项目结果大于              或结果小于             提示复查"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   435
            TabIndex        =   78
            Top             =   2025
            Width           =   5310
         End
         Begin VB.Label lblAbnorma 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "危急值下限              危急值上限"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   750
            TabIndex        =   77
            Top             =   1650
            Width           =   3060
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgReference 
         Height          =   1260
         Left            =   345
         TabIndex        =   75
         Top             =   840
         Width           =   5820
         _cx             =   10266
         _cy             =   2222
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
      Begin XtremeCommandBars.CommandBars cbsSub3 
         Left            =   480
         Top             =   285
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9060
      Index           =   0
      Left            =   2130
      ScaleHeight     =   9030
      ScaleWidth      =   9495
      TabIndex        =   24
      Top             =   105
      Width           =   9525
      Begin VB.PictureBox picSub0 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8355
         Left            =   240
         ScaleHeight     =   8355
         ScaleWidth      =   9405
         TabIndex        =   30
         Top             =   375
         Width           =   9405
         Begin VB.TextBox txtDecimal 
            Height          =   300
            Left            =   3275
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1950
            Width           =   1300
         End
         Begin VB.TextBox txtManual 
            Height          =   1185
            Left            =   990
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   5040
            Width           =   8325
         End
         Begin VB.CheckBox chkGLU 
            Caption         =   "糖耐量项目"
            Height          =   240
            Left            =   7710
            TabIndex        =   16
            Top             =   3090
            Width           =   1260
         End
         Begin VB.ComboBox cboDataType 
            Height          =   300
            ItemData        =   "frmBaseItemList.frx":0000
            Left            =   5730
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   345
            Width           =   3585
         End
         Begin VB.TextBox txtDefault 
            Height          =   300
            Left            =   990
            MaxLength       =   200
            TabIndex        =   7
            Top             =   1950
            Width           =   1300
         End
         Begin VB.CheckBox chkMorInformation 
            Caption         =   "多参考"
            Height          =   210
            Left            =   6682
            TabIndex        =   15
            Top             =   3090
            Width           =   900
         End
         Begin VB.CheckBox chkPrivacy 
            Caption         =   "隐私项目"
            Height          =   195
            Left            =   5460
            TabIndex        =   14
            Top             =   3090
            Width           =   1020
         End
         Begin VB.CommandButton cmdFormula 
            Caption         =   "∑"
            Height          =   300
            Left            =   9015
            TabIndex        =   33
            Top             =   2340
            Width           =   300
         End
         Begin VB.TextBox txtFormula 
            Height          =   300
            Left            =   990
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   10
            Top             =   2355
            Width           =   8025
         End
         Begin VB.TextBox txtTestMethods 
            Height          =   300
            Left            =   5730
            MaxLength       =   40
            TabIndex        =   9
            Top             =   1950
            Width           =   3585
         End
         Begin VB.TextBox txtOutOfControlRate 
            Height          =   300
            Left            =   2865
            MaxLength       =   12
            TabIndex        =   18
            Top             =   3705
            Width           =   615
         End
         Begin VB.TextBox txtAlertsRate 
            Height          =   300
            Left            =   1170
            MaxLength       =   12
            TabIndex        =   17
            Top             =   3705
            Width           =   615
         End
         Begin VB.TextBox txtVariationAlerts 
            Height          =   300
            Left            =   2865
            MaxLength       =   12
            TabIndex        =   20
            Top             =   4080
            Width           =   615
         End
         Begin VB.TextBox txtVariationAlarm 
            Height          =   300
            Left            =   1170
            MaxLength       =   12
            TabIndex        =   19
            Top             =   4080
            Width           =   615
         End
         Begin VB.TextBox txtWBCode 
            Height          =   300
            Left            =   5730
            MaxLength       =   60
            TabIndex        =   32
            Top             =   1545
            Width           =   3585
         End
         Begin VB.TextBox txtPYCode 
            Height          =   300
            Left            =   5730
            MaxLength       =   60
            TabIndex        =   31
            Top             =   1155
            Width           =   3585
         End
         Begin VB.TextBox txtUnits 
            Height          =   300
            Left            =   990
            MaxLength       =   10
            TabIndex        =   6
            Top             =   1545
            Width           =   3585
         End
         Begin VB.TextBox txtEnglish 
            Height          =   300
            Left            =   5730
            MaxLength       =   40
            TabIndex        =   4
            Top             =   750
            Width           =   3585
         End
         Begin VB.TextBox txtChineseName 
            Height          =   300
            Left            =   990
            MaxLength       =   60
            TabIndex        =   5
            Top             =   1155
            Width           =   3585
         End
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   990
            MaxLength       =   20
            TabIndex        =   3
            Top             =   750
            Width           =   3585
         End
         Begin VB.TextBox txtType 
            Height          =   300
            Left            =   990
            MaxLength       =   40
            TabIndex        =   1
            Top             =   345
            Width           =   3585
         End
         Begin VB.OptionButton optType 
            Caption         =   "普通项目"
            Height          =   180
            Index           =   0
            Left            =   285
            TabIndex        =   11
            Top             =   3090
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton optType 
            Caption         =   "计算项目"
            Height          =   180
            Index           =   1
            Left            =   1425
            TabIndex        =   12
            Top             =   3090
            Width           =   1050
         End
         Begin VB.OptionButton optType 
            Caption         =   "酶标项目"
            Height          =   180
            Index           =   2
            Left            =   2580
            TabIndex        =   13
            Top             =   3090
            Width           =   1050
         End
         Begin VB.Image imgNote 
            Height          =   240
            Index           =   1
            Left            =   105
            Picture         =   "frmBaseItemList.frx":0014
            Top             =   6225
            Width           =   240
         End
         Begin VB.Label Label13 
            Caption         =   "临床意义"
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   60
            TabIndex        =   74
            Top             =   4740
            Width           =   5745
         End
         Begin VB.Label Label5 
            Caption         =   "默认结果                  小数位数"
            Height          =   180
            Left            =   135
            TabIndex        =   65
            Top             =   1995
            Width           =   3240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计算公式"
            Height          =   180
            Index           =   10
            Left            =   135
            TabIndex        =   49
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "比对警示率         比对失控率        (相同样本在不同仪器上检测结果的对比)"
            Height          =   180
            Index           =   9
            Left            =   210
            TabIndex        =   48
            Top             =   3765
            Width           =   6570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "变异报警           变异警示        (本次检测结果与上次检测结果的差异)"
            Height          =   180
            Left            =   390
            TabIndex        =   47
            Top             =   4140
            Width           =   6210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计算单位"
            Height          =   180
            Index           =   6
            Left            =   135
            TabIndex        =   46
            Top             =   1590
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "项目名称*"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   5
            Left            =   135
            TabIndex        =   45
            Top             =   1185
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "项目代码*"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   4
            Left            =   135
            TabIndex        =   44
            Top             =   795
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "项目分类"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   43
            Top             =   390
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "项目类别"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4770
            TabIndex        =   42
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "报告单代号"
            Height          =   195
            Left            =   4770
            TabIndex        =   41
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label Label4 
            Caption         =   "测试方法"
            Height          =   225
            Left            =   4770
            TabIndex        =   40
            Top             =   1995
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "五笔简码"
            Height          =   180
            Left            =   4770
            TabIndex        =   39
            Top             =   1590
            Width           =   810
         End
         Begin VB.Label Label7 
            Caption         =   "拼音简码"
            Height          =   195
            Left            =   4770
            TabIndex        =   38
            Top             =   1185
            Width           =   750
         End
         Begin VB.Label Label8 
            Caption         =   "基本信息"
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   60
            TabIndex        =   37
            Top             =   60
            Width           =   1260
         End
         Begin VB.Label Label9 
            Caption         =   "其他信息"
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   60
            TabIndex        =   36
            Top             =   2790
            Width           =   1035
         End
         Begin VB.Label Label10 
            Caption         =   "警示与比对"
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   60
            TabIndex        =   35
            Top             =   3420
            Width           =   5745
         End
         Begin VB.Label lblHelp 
            ForeColor       =   &H00C00000&
            Height          =   825
            Left            =   45
            TabIndex        =   34
            Top             =   6270
            Width           =   6870
         End
      End
      Begin XtremeCommandBars.CommandBars cbsSub0 
         Left            =   435
         Top             =   60
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   1140
      ScaleHeight     =   3465
      ScaleWidth      =   6450
      TabIndex        =   25
      Top             =   6570
      Width           =   6480
      Begin VB.Frame fraBase 
         BorderStyle     =   0  'None
         Height          =   1470
         Left            =   705
         TabIndex        =   66
         Top             =   285
         Width           =   5820
         Begin VB.ComboBox cboMachine1 
            Height          =   300
            ItemData        =   "frmBaseItemList.frx":059E
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1110
            Width           =   4590
         End
         Begin VB.Image imgNote 
            Height          =   240
            Index           =   0
            Left            =   210
            Picture         =   "frmBaseItemList.frx":05B2
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label12 
            Caption         =   "检验仪器"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   1170
            Width           =   795
         End
         Begin VB.Label lblSequence 
            Caption         =   $"frmBaseItemList.frx":0B3C
            ForeColor       =   &H00C00000&
            Height          =   750
            Left            =   60
            TabIndex        =   68
            Top             =   210
            Width           =   4620
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgSequence 
         Height          =   1125
         Left            =   495
         TabIndex        =   29
         Top             =   2190
         Width           =   5325
         _cx             =   9393
         _cy             =   1984
         Appearance      =   3
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
      Begin XtremeCommandBars.CommandBars cbsSub1 
         Left            =   210
         Top             =   165
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   2
      Left            =   180
      ScaleHeight     =   3465
      ScaleWidth      =   6450
      TabIndex        =   26
      Top             =   6270
      Width           =   6480
      Begin VSFlex8Ctl.VSFlexGrid vfgChannel 
         Bindings        =   "frmBaseItemList.frx":0BB8
         Height          =   1530
         Left            =   360
         TabIndex        =   73
         Top             =   1680
         Width           =   5385
         _cx             =   9499
         _cy             =   2699
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
      Begin VB.Frame fraSub2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   195
         TabIndex        =   70
         Top             =   435
         Width           =   5820
         Begin VB.ComboBox cboMachine2 
            Height          =   300
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   390
            Width           =   4590
         End
         Begin VB.Label lblMachine 
            Caption         =   "检验仪器"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   450
            Width           =   795
         End
      End
      Begin XtremeCommandBars.CommandBars cbsSub2 
         Left            =   105
         Top             =   120
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Frame FraWE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   2250
      MousePointer    =   9  'Size W E
      TabIndex        =   28
      Top             =   -750
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tabBase 
      Bindings        =   "frmBaseItemList.frx":0BCC
      Height          =   1890
      Left            =   5460
      TabIndex        =   23
      Top             =   285
      Width           =   6405
      _Version        =   589884
      _ExtentX        =   11298
      _ExtentY        =   3334
      _StockProps     =   64
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8865
      Left            =   540
      ScaleHeight     =   8835
      ScaleWidth      =   5160
      TabIndex        =   0
      Top             =   585
      Width           =   5190
      Begin XtremeReportControl.ReportControl RptItem 
         Height          =   7725
         Left            =   270
         TabIndex        =   22
         Top             =   315
         Width           =   4140
         _Version        =   589884
         _ExtentX        =   7302
         _ExtentY        =   13626
         _StockProps     =   0
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   180
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBaseItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsItems As ADODB.Recordset '业务数据集-检验指标
Private mrsMachineItems As ADODB.Recordset '业务数据集-检验仪器指标

Private mPages As Collection        '属性页数据集
Private mminHeight As Single, mminWidth As Single '各区域最小高度与宽度
Private Enum mRpt
     ID = 0: 指标代码: 中文名: 英文名: 拼音码: 五笔码: 项目分类: 项目类别: 结果类型: 单位: 计算公式: 检验方法: 默认值: 变异警示率: 变异报警率: 比对警示率: 比对失控率: 打印序号: 排列序号: 临床意义: 隐私项目: 多参考: 小数位数: 糖耐量项目
End Enum
Private mintCurrRow As Integer '保存当前选中行，当刷新时重新定位到原来选择行。
Private mlngItemID As Long      '当前项目ID
Private mintEditIndex As Integer '当前开始编辑的页面。0-未开始编辑 1-项目属性页 2-常用取值页 3-参考页 4-临床意义页
Private mintType As Integer         '1-新增 2-修改 3-删除
Private WithEvents mfrmFind As frmPubRptFind
Attribute mfrmFind.VB_VarHelpID = -1

Private mstrPrivs As String                 '模块权限

Private Enum mTabIndex
    Tab_基本信息 = 0: Tab_常用取值: Tab_通道码: Tab_项目参考
End Enum
'----------------------------------------------------------------------------------
'--- 以下为本窗体公共过程
'----------------------------------------------------------------------------------


Private Sub cbsMainInit()
    Dim Menus As New Collection
    Menus.Add conMenu_Edit_Add & ",仪器管理(&D)  ,False"
    Menus.Add conMenu_Edit_BillSet & ",检验单据(&B)  ,True"
'    Menus.Add conMenu_Edit_ItemImport & ",导入项目(&I)  ,True"
'    Menus.Add conMenu_Edit_ItemExport & ",导出项目(&E)  ,False"
    Menus.Add conMenu_Edit_Refresh & ",刷新(&E)  ,False"
    Menus.Add conMenu_Edit_Find & ",查找(&F)  ,False"
    Menus.Add conMenu_Edit_ItemSort & ",调整顺序(&S)  ,True"
    Menus.Add conMenu_Edit_Exit & ",退出(&E)      ,True"
    Call CbsButtonInit(cbsMain, Menus, True, xtpBarBottom)
    Set Menus = Nothing
End Sub

Private Sub cbsSubInit(ByRef cbsSub As CommandBars)
    Dim Menus As New Collection
    Menus.Add conMenu_Edit_ItemAdd & ",增加,False"
    Menus.Add conMenu_Edit_ItemEdit & ",修改,False"
    Menus.Add conMenu_Edit_ItemDele & ",删除,False"
    Menus.Add conMenu_Edit_ItemUndo & ",取消,False"
    Menus.Add conMenu_Edit_ItemSave & ",保存,False"
    Call CbsButtonInit(cbsSub, Menus, False, xtpBarTop)
    Set Menus = Nothing
End Sub

Private Sub cbsSubInit1(ByRef cbsSub As CommandBars)
    Dim Menus As New Collection
    Menus.Add conMenu_Edit_ItemEdit & ",编辑,False"
    Menus.Add conMenu_Edit_ItemAdd & ",增加行,True"
    Menus.Add conMenu_Edit_ItemDele & ",删除行,False"
    Menus.Add conMenu_Edit_ItemUndo & ",取消,True"
    Menus.Add conMenu_Edit_ItemSave & ",保存,False"
    Call CbsButtonInit(cbsSub, Menus, False, xtpBarTop)
    Set Menus = Nothing
End Sub

Private Sub LoadBaseData()
    '装入本窗体中CBO控件要用的数据
    Dim rsMachine As ADODB.Recordset
    Dim strData As String
    Dim strErr As String
    
    '取所有仪器的项目
    If Not GetMachineItems(0, mrsMachineItems, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    
    'picInfo(0) 结果类型
    If Not GetResultTypeToCbo(Me.cboDataType, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If

    'picInfo(1) 常用取值的检验仪器
    If Not GetMachineToCbo(cboMachine1, True, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    'picInfo(2) 通道码的检验仪器
    If Not GetMachineToCbo(cboMachine2, True, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    'picInfo(3) 参考的检验仪器
    If Not GetMachineToCbo(cboMachine3, True, strErr, "-") Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    '           标本种类
    If Not GetSampleTypeToCbo(cboSampleType, strErr, True, "-") Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    '          临床特征
    If Not GetFeatureToCbo(cboFeatures, True, strErr, "-") Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    '          年龄单位
    If Not GetAgeUnitToCbo(cboAgeUnitl, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    '           性别
    If Not GetSexToCbo(cboSex, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    
End Sub

Private Sub Find()
    Dim strFindSQL As String, strFindFiled As String
    If mfrmFind Is Nothing Then Set mfrmFind = New frmPubRptFind
    strFindSQL = "Select ID " & vbNewLine & _
            "From 检验指标 Where 指标代码 Like [1] Or 中文名 like [1] Or 英文名 Like [1] Or 拼音码 Like [1] Or 五笔码 Like [1]"
    Call mfrmFind.ShowFind(strFindSQL)
    
End Sub
Private Sub RefRptItemData()
    '装载项目数据
    Dim strErr As String, i As Integer
    
    If Not GetItems(strErr, mrsItems) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    Else
        '将数据装入RPT控件
        Call RptLoadFromRecord(RptItem, mrsItems, strErr)
        '对装入的数据做特殊处理
        For i = 0 To RptItem.Records.count - 1
            '普通项目 = 1    计算项目 = 3    酶标项目 = 4
            If RptItem.Records(i).Item(mRpt.项目类别).Value = "4" Then
                RptItem.Records(i).Item(mRpt.项目类别).Value = "酶标项目"
                RptItem.Records(i).Item(mRpt.项目类别).Caption = "酶标项目"
            ElseIf RptItem.Records(i).Item(mRpt.项目类别).Value = "3" Then
                RptItem.Records(i).Item(mRpt.项目类别).Value = "计算项目"
                RptItem.Records(i).Item(mRpt.项目类别).Caption = "计算项目"
            Else
                RptItem.Records(i).Item(mRpt.项目类别).Value = "普通项目"
                RptItem.Records(i).Item(mRpt.项目类别).Caption = "普通项目"
            End If
            '定量 = 1    定性 = 2    半定量 = 3
            If RptItem.Records(i).Item(mRpt.结果类型).Value = "1" Then
                RptItem.Records(i).Item(mRpt.结果类型).Value = "定量"
                RptItem.Records(i).Item(mRpt.结果类型).Caption = "定量"
            ElseIf RptItem.Records(i).Item(mRpt.结果类型).Value = "2" Then
                RptItem.Records(i).Item(mRpt.结果类型).Value = "定性"
                RptItem.Records(i).Item(mRpt.结果类型).Caption = "定性"
            Else
                RptItem.Records(i).Item(mRpt.结果类型).Value = "半定量"
                RptItem.Records(i).Item(mRpt.结果类型).Caption = "半定量"
            End If
            
            If RptItem.Records(i).Item(mRpt.项目分类).Value = "" Then
                RptItem.Records(i).Item(mRpt.项目分类).Value = "未分类"
                RptItem.Records(i).Item(mRpt.项目分类).Caption = "未分类"
            End If
            RptItem.Records(i).Item(mRpt.排列序号).Value = Val("" & RptItem.Records(i).Item(mRpt.排列序号).Value)
        Next
        '加入分类
        RptItem.GroupsOrder.DeleteAll
        Call RptItem.GroupsOrder.Add(RptItem.Columns(mRpt.项目分类))
        
        RptItem.Columns(mRpt.指标代码).Visible = True
        RptItem.Columns(mRpt.指标代码).Width = 50
        RptItem.Columns(mRpt.指标代码).Alignment = xtpAlignmentLeft
        
        RptItem.Columns(mRpt.中文名).Visible = True
        RptItem.Columns(mRpt.中文名).Width = 120
        RptItem.Columns(mRpt.中文名).Alignment = xtpAlignmentLeft
        
        RptItem.Columns(mRpt.项目类别).Visible = True
        RptItem.Columns(mRpt.项目类别).Width = 60
        
        RptItem.Columns(mRpt.结果类型).Visible = True
        RptItem.Columns(mRpt.结果类型).Width = 50
        
        RptItem.SortOrder.DeleteAll
        Call RptItem.SortOrder.Add(RptItem.Columns(mRpt.排列序号))
        Call RptItem.SortOrder.Add(RptItem.Columns(mRpt.指标代码))
        
        RptItem.Populate
    End If
    
    '提取检验仪器指标数据，分页中要用
    
    Call RptItem_SelectionChanged
   
End Sub

Private Sub cboAgeUnitl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cboDataType_Click()
    '禁止或启用公式
    
End Sub

Private Sub cboDataType_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cboFeatures_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cboMachine3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cboSampleType_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub chkDefault_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub chkGLU_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub chkMorInformation_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub chkPrivacy_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub cmdFormula_Click()
    If optType(1).Value And Val(cboDataType.List(cboDataType.ListIndex)) = 1 Then
        txtFormula = FrmBaseItemFormula.DefFormula(mlngItemID, txtFormula, Me)
    Else
        MsgBox "只有定量的计算项目才能设置公式！", vbQuestion, Me.Caption
    End If
End Sub

Private Sub cmdFormula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmFind Is Nothing Then Unload mfrmFind
    Set mfrmFind = Nothing
End Sub

Private Sub mfrmFind_Finded(ByVal blnFind As Boolean, ByVal strVale As String)
    '定位:
    Dim varTmp As Variant, strID As String
    If blnFind Then
        varTmp = Split(strVale, ",")
        strID = varTmp(0)
        Call RptFindRowToFocuse(Me.RptItem, mRpt.ID, strID)
        Call RptItem_SelectionChanged
    End If
End Sub

Private Sub optType_Click(Index As Integer)
    
    '非计算项目不能设置计算公式
    If Me.optType(1).Value = True Then
        Me.txtFormula.Text = Me.txtFormula.Tag: Me.txtFormula.Enabled = True
        Me.cmdFormula.Enabled = True
    Else
        Me.txtFormula.Tag = Me.txtFormula.Text: Me.txtFormula.Enabled = False
        Me.cmdFormula.Enabled = False
    End If
    
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub picInfo_Resize(Index As Integer)
    If Index = 0 Then Call cbsSub0_Resize
    If Index = 1 Then Call cbsSub1_Resize
    If Index = 2 Then Call cbsSub2_Resize
    If Index = 3 Then Call cbsSub3_Resize
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With RptItem
        .Left = picLeft.Left + 45
        .Top = picLeft.Top + 45
        .Width = picLeft.Width - 90
        .Height = picLeft.Height - 90
    End With
End Sub

Private Sub RptItem_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Row.GroupRow = False Then
         
        If tabBase.Selected.Index = Tab_基本信息 Then
            If cbsSub0.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit).Enabled = True Then
                Call cbsSub0_Execute(cbsSub0.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit))
                On Error Resume Next
                If picSub0.Enabled Then txtNo.SetFocus
            End If
        ElseIf tabBase.Selected.Index = Tab_常用取值 Then
            If cbsSub1.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit).Enabled = True Then
                Call cbsSub1_Execute(cbsSub1.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit))
            End If
        ElseIf tabBase.Selected.Index = Tab_通道码 Then
            If cbsSub2.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit).Enabled = True Then
                Call cbsSub2_Execute(cbsSub2.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit))
            End If
        ElseIf tabBase.Selected.Index = Tab_项目参考 Then
            If cbsSub3.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit).Enabled = True Then
                Call cbsSub3_Execute(cbsSub3.ActiveMenuBar.FindControl(, conMenu_Edit_ItemEdit))
                On Error Resume Next
                If picSub3.Enabled Then txtReference(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub RptItem_SelectionChanged()
    '设置选中行
    Dim strTmp As String
    Call RptSelectRecord(RptItem, mintCurrRow)
    ClearPicInfo0
    If RptItem.SelectedRows.count = 0 Then Exit Sub '没有选择的行,则退出
    With RptItem
        If Not .SelectedRows(0).GroupRow Then
            mintCurrRow = .SelectedRows(0).Index
            mlngItemID = Val("" & .SelectedRows(0).Record(mRpt.ID).Value)
        Else
            mlngItemID = 0
        End If
    End With
    '基本属性页
    Call RefPicInfo0
    '常用取值页
    Call RefPicinfo1
    '通道码页
    Call RefPicinfo2
    '参考页
    Call RefPicInfo3
    
End Sub

Private Sub RefPicInfo0()
    '显示基本信息页数据
    Dim strTmp As String
     
    With RptItem
        If Not .SelectedRows(0).GroupRow Then
            '显示选中项目的基本信息
            txtType = "" & .SelectedRows(0).Record(mRpt.项目分类).Value
            strTmp = "" & .SelectedRows(0).Record(mRpt.结果类型).Value
            Call cboSelect(cboDataType, strTmp, 1)
            txtNo = Trim("" & .SelectedRows(0).Record(mRpt.指标代码).Value)
            txtChineseName = Trim("" & .SelectedRows(0).Record(mRpt.中文名).Value)
            txtEnglish = Trim("" & .SelectedRows(0).Record(mRpt.英文名).Value)
            txtUnits = Trim("" & .SelectedRows(0).Record(mRpt.单位).Value)
            txtDefault = Trim("" & .SelectedRows(0).Record(mRpt.默认值).Value)
            txtTestMethods = Trim("" & .SelectedRows(0).Record(mRpt.检验方法).Value)
            txtPYCode = Trim("" & .SelectedRows(0).Record(mRpt.拼音码).Value)
            txtWBCode = Trim("" & .SelectedRows(0).Record(mRpt.五笔码).Value)
            txtFormula = Trim("" & .SelectedRows(0).Record(mRpt.计算公式).Value)
            txtFormula.Tag = Trim("" & .SelectedRows(0).Record(mRpt.计算公式).Value)
            '1-普通;3-计算项目;4-酶标项目
            strTmp = "" & .SelectedRows(0).Record(mRpt.项目类别).Value
             
            If strTmp = "普通项目" Then
                optType(0).Value = True
            ElseIf strTmp = "计算项目" Then
                optType(1).Value = True
            ElseIf strTmp = "酶标项目" Then
                optType(2).Value = True
            End If
            cmdFormula.Enabled = optType(1).Value
            strTmp = "" & .SelectedRows(0).Record(mRpt.隐私项目).Value
            chkPrivacy.Value = Val(strTmp)
            strTmp = "" & .SelectedRows(0).Record(mRpt.多参考).Value
            chkMorInformation.Value = Val(strTmp)
            txtAlertsRate = Trim("" & .SelectedRows(0).Record(mRpt.比对警示率).Value)
            txtOutOfControlRate = Trim("" & .SelectedRows(0).Record(mRpt.比对失控率).Value)
            txtVariationAlarm = Trim("" & .SelectedRows(0).Record(mRpt.变异报警率).Value)
            txtVariationAlerts = Trim("" & .SelectedRows(0).Record(mRpt.变异警示率).Value)
            txtManual = Trim("" & .SelectedRows(0).Record(mRpt.临床意义).Value)
            txtDecimal = Trim("" & .SelectedRows(0).Record(mRpt.小数位数).Value)
        End If
    End With
End Sub

Private Sub RefPicinfo1()
    '从vfgItem中直接得到项目属性，填到界面上。
    '常用取值页
    Dim i As Integer, strTmp As String, lngMachineID As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim varTmp As Variant, varTmp1 As Variant, varTmp2 As Variant
    On Error GoTo errH
    strTmp = "显示取值,1800," & flexAlignLeftCenter
    strTmp = strTmp & ";质控取值,1800," & flexAlignLeftCenter
    strTmp = strTmp & ";结果标志,2200," & flexAlignCenterCenter
    With vfgSequence
        .Clear
        Call vfgSetting(0, vfgSequence, strTmp)
        .ColComboList(.ColIndex("结果标志")) = "|#0;|#1;1-正常|#2;2-偏低|#3;3-偏高|#4;4-异常|#5;5-警戒下限|#6;6-警戒上限|#7;7-复查下限|#8;8-复查上限"
        .Editable = flexEDKbdMouse
        
        If mlngItemID <= 0 Then Exit Sub
        If cboMachine1.ListCount <= 0 Then Exit Sub
        '取数据
        lngMachineID = cboMachine1.ItemData(cboMachine1.ListIndex)
        strSQL = "Select 取值序列, 序列取值, 结果标志 From 检验仪器指标 Where 仪器ID=[1] And 项目ID=[2]"
        Set rsTmp = ComOpenSQL(strSQL, "取常用取值", lngMachineID, mlngItemID)
        If rsTmp.EOF Then Exit Sub
        
        strTmp = Trim("" & rsTmp!取值序列)
        varTmp = Split(strTmp, ",")
        strTmp = Trim("" & rsTmp!序列取值)
        varTmp1 = Split(strTmp, ",")
        strTmp = Trim("" & rsTmp!结果标志)
        varTmp2 = Split(strTmp, ",")
        
        '显示
        .Rows = .FixedRows
        For i = LBound(varTmp) To UBound(varTmp)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("显示取值")) = Trim("" & varTmp(i))
            If LBound(varTmp1) <= i And UBound(varTmp1) >= i And UBound(varTmp1) >= 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("质控取值")) = Trim("" & varTmp1(i))
            Else
                .TextMatrix(.Rows - 1, .ColIndex("质控取值")) = ""
            End If
            If LBound(varTmp2) <= i And UBound(varTmp2) >= i And UBound(varTmp2) >= 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("结果标志")) = IIf(Val("" & varTmp2(i)) = 0, "", Val("" & varTmp2(i)))
            Else
                .TextMatrix(.Rows - 1, .ColIndex("结果标志")) = ""
            End If
        Next
    End With
    Exit Sub
errH:
    If ComErrCenter() = 1 Then
        Resume
    End If
    Call SaveLog(Me.Caption & "-显示基本属性", LOG_ERR, Err.Number, Err.Description)
End Sub

Private Sub RefPicinfo2()
    '从vfgItem中直接得到项目属性，填到界面上。
    '常用取值页
    Dim i As Integer, strTmp As String, lngMachineID As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim varTmp As Variant, varTmp1 As Variant, varTmp2 As Variant
    On Error GoTo errH
    strTmp = "通道编码,1800," & flexAlignLeftCenter
    
    With vfgChannel
        .Clear
        Call vfgSetting(0, vfgChannel, strTmp)
        .Editable = flexEDKbdMouse
        
        If mlngItemID <= 0 Then Exit Sub
        If cboMachine2.ListCount <= 0 Then Exit Sub
        '取数据
        lngMachineID = cboMachine2.ItemData(cboMachine2.ListIndex)
        strSQL = "Select 通道编码 From 检验仪器指标 Where 仪器ID=[1] And 项目ID=[2]"
        Set rsTmp = ComOpenSQL(strSQL, "取常用取值", lngMachineID, mlngItemID)
        If rsTmp.EOF Then Exit Sub
        
        strTmp = Trim("" & rsTmp!通道编码)
        varTmp = Split(strTmp, ",")
        
        '显示
        .Rows = .FixedRows
        For i = LBound(varTmp) To UBound(varTmp)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("通道编码")) = Trim("" & varTmp(i))
        Next
    End With
    Exit Sub
errH:
    If ComErrCenter() = 1 Then
        Resume
    End If
    Call SaveLog(Me.Caption & "-显示常用取值", LOG_ERR, Err.Number, Err.Description)

End Sub
Private Sub RefPicInfo3()
    '刷新项目参考页数据
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intRow As Integer, strErr As String
    On Error GoTo errH
    vfgReference.Tag = ""
    strSQL = "Select a.id,b.名称 as 仪器,a.默认, a.标本类型 as 标本, a.性别域 as 性别, a.年龄上限, a.年龄下限, a.年龄单位,'' as 年龄, a.临床特征, a.参考低值, a.参考高值, a.参考描述 as 参考值," & _
             "a.警示下限 as 危急低值, a.警示上限 as 危急高值, a.复查下限 ,  a.复查上限" & vbNewLine & _
            "From 检验指标参考 A,检验仪器记录 B " & vbNewLine & _
            "Where a.仪器id=B.id(+) And a.项目id = [1] order by a.id"
    Set rsTmp = ComOpenSQL(strSQL, Me.Caption, mlngItemID)
    If Not vfgLoadFromRecord(vfgReference, rsTmp, strErr) Then
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    
    With vfgReference
        .Tag = "Ok"
        .Redraw = flexRDNone
        
        For intRow = .FixedRows To .Rows - 1
            '小数点处理
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("年龄下限"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("年龄下限")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("年龄下限")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("年龄上限"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("年龄上限")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("年龄上限")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("参考高值"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("参考高值")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("参考高值")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("参考低值"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("参考低值")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("参考低值")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("复查上限"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("复查上限")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("复查上限")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("复查下限"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("复查下限")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("复查下限")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("危急高值"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("危急高值")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("危急高值")))
             If Left(Trim("" & .TextMatrix(intRow, .ColIndex("危急低值"))), 1) = "." Then .TextMatrix(intRow, .ColIndex("危急低值")) = "0" & Trim("" & .TextMatrix(intRow, .ColIndex("危急低值")))
             
            '年龄显示处理
            If Val(.TextMatrix(intRow, .ColIndex("年龄下限"))) <> 0 And Val(.TextMatrix(intRow, .ColIndex("年龄上限"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("年龄")) = .TextMatrix(intRow, .ColIndex("年龄下限")) & "～" & .TextMatrix(intRow, .ColIndex("年龄上限")) & .TextMatrix(intRow, .ColIndex("年龄单位"))
            ElseIf Val(.TextMatrix(intRow, .ColIndex("年龄下限"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("年龄")) = ">=" & .TextMatrix(intRow, .ColIndex("年龄下限")) & .TextMatrix(intRow, .ColIndex("年龄单位"))
            ElseIf Val(.TextMatrix(intRow, .ColIndex("年龄上限"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("年龄")) = "<=" & .TextMatrix(intRow, .ColIndex("年龄上限")) & .TextMatrix(intRow, .ColIndex("年龄单位"))
            Else
                .TextMatrix(intRow, .ColIndex("年龄")) = ""
            End If
        Next
        .RowHeight(.FixedRows - 1) = 500
        
        .ColWidth(.ColIndex("默认")) = 500
        .ColDataType(.ColIndex("默认")) = flexDTBoolean
        
        .ColWidth(.ColIndex("标本")) = 900
        
        .ColComboList(.ColIndex("性别")) = "|#0;不限|#1;男|#2;女|#9;未说明"
        .ColWidth(.ColIndex("性别")) = 600
        
        .ColWidth(.ColIndex("年龄")) = 600
        .ColWidth(.ColIndex("临床特征")) = 1000
        .ColWidth(.ColIndex("参考低值")) = 600
        .ColWidth(.ColIndex("参考高值")) = 600
        .ColWidth(.ColIndex("参考值")) = 1200
        .ColWidth(.ColIndex("危急低值")) = 600
        .ColWidth(.ColIndex("危急高值")) = 600
        .ColWidth(.ColIndex("复查下限")) = 600
        .ColWidth(.ColIndex("复查上限")) = 600
        .Redraw = flexRDDirect
    End With
    Call vfgReference_RowColChange
    Exit Sub
errH:
    If ComErrCenter() = 1 Then
        Resume
    End If
    Call SaveLog(Me.Caption & "-显示参考值", LOG_ERR, Err.Number, Err.Description)
    
End Sub
Private Sub LockEdit(ByVal Index As Integer, ByVal blnLock As Boolean)
    'index:
    '编辑前锁定界面，保存后解除锁定
    Dim i As Integer
    If blnLock Then
        '锁定除编辑页面之外的控件
        
        '左边列表
        picLeft.Enabled = False
        '项目属性页
        picSub0.Enabled = Index = 0
         
        '常用取值页
        vfgSequence.Enabled = (Index = 1)
        '通道码页
        vfgChannel.Enabled = (Index = 2)
        '参考
        picSub3.Enabled = (Index = 3)
        vfgReference.Enabled = Not (Index = 3)
        mintEditIndex = Index + 1
    Else
        '退出编辑，重置控件状态
        picLeft.Enabled = True
        '项目属性页
        picSub0.Enabled = False
        '常用取值页
        vfgSequence.Enabled = False
        '通道码
        vfgChannel.Enabled = False
        mintEditIndex = 0
        '参考
        picSub3.Enabled = False
        vfgReference.Enabled = True
    End If
End Sub
Private Sub ClearPicInfo0()
    '清空属性页中控件的值
    If Not mintEditIndex = 1 Then
        txtType = ""
    End If
    
    If cboDataType.ListCount > 0 Then cboDataType.ListIndex = 0
    txtNo = ""
    txtChineseName = ""
    txtEnglish = ""
    txtUnits = ""
    txtDefault = ""
    txtFormula = ""
    txtFormula.Tag = ""
    txtTestMethods = ""
    txtVariationAlarm = ""
    txtVariationAlerts = ""
    txtAlertsRate = ""
    txtOutOfControlRate = ""
    txtPYCode = ""
    txtWBCode = ""
    optType(0).Value = True
    chkMorInformation.Value = 0
    chkPrivacy.Value = 0
    txtManual = ""
    txtDecimal = "2"
End Sub
Private Function SaveItem(ByVal iType As Integer) As Boolean
    '开始修改
    'iType: 1-增加 2-修改 3-删除
    Dim strSQL As String
    Dim strNO As String, strChina As String '项目代码,名称
    Dim strEnglish As String, strTYPE As String, intItemType As Integer '英文名，项目分类,项目类型
    Dim intResultType As Integer, strUntil As String '结果类型，单位
    Dim strFormu As String, strMethods As String '计算公式，实验方法
    Dim strDefault As String, intVariationAlarm As Integer '默认值,变异报警
    Dim intVariationAlerts As Integer, intAlertsRate As Integer '变异警示,比对警示
    Dim intOutOfControlRate As Integer, intPrivacy As Integer '比对失控,隐私项
    Dim intGLU As Integer, intMorInformation As Integer '糖耐量,多参考
    Dim strPY  As String, strWb  As String  '拼音,五笔
    Dim intDecimal As Integer, strManual As String  '小数,临床意义
    Dim strInvalidWord As String        '禁止输入的字符
    Dim strErr As String
    If iType = 3 Then
        If mlngItemID <> 0 Then
            If ExecProcEditItems(3, strErr, mlngItemID) Then
                SaveItem = True
                Call RefRptItemData
            Else
                MsgBox strErr, vbInformation, Me.Caption
                Exit Function
            End If
        Else
            MsgBox "请选中一行数据后再执行操作！", vbInformation, Me.Caption
            Exit Function
        End If
    Else
        strInvalidWord = Replace(gSysParameter.InvaidWord, "%", "")
        strInvalidWord = Replace(strInvalidWord, "#", "")
        strNO = UCase(StringDelInvalidWord(txtNo, strInvalidWord))
'        If Not IsLettersAndDigital(strNO) Then
'            MsgBox "指标代码只能为英文字母与数字！", vbInformation, Me.Caption
'            Exit Function
'        End If
        
        strChina = StringDelInvalidWord(txtChineseName, strInvalidWord)
        strEnglish = StringDelInvalidWord(txtEnglish, strInvalidWord)
        strTYPE = StringDelInvalidWord(txtType)
        '1-普通;3-计算项目;4-酶标项目
        If optType(0).Value = True Then
            intItemType = 1
        ElseIf optType(1).Value = True Then
            intItemType = 3
        ElseIf optType(2).Value = True Then
            intItemType = 4
        End If
        '"1-定量;2-定性;3-半定量"
        intResultType = Val(Split(cboDataType.List(cboDataType.ListIndex), "-")(0))
        strUntil = StringDelInvalidWord(txtUnits, strInvalidWord)
        
        
        strInvalidWord = Replace(gSysParameter.InvaidWord, "]", "")
        strInvalidWord = Replace(strInvalidWord, "[", "")
        
        strFormu = StringDelInvalidWord(txtFormula, strInvalidWord)
        strMethods = StringDelInvalidWord(txtTestMethods)
        strDefault = StringDelInvalidWord(txtDefault)
        intVariationAlarm = Val(txtVariationAlarm)
        intVariationAlerts = Val(txtVariationAlerts)
        intAlertsRate = Val(txtAlertsRate)
        intOutOfControlRate = Val(txtOutOfControlRate)
        intPrivacy = chkPrivacy.Value
        
        intMorInformation = chkMorInformation.Value
        strPY = StringDelInvalidWord(txtPYCode)
        strWb = StringDelInvalidWord(txtWBCode)
        intGLU = chkGLU.Value
        intDecimal = Val(txtDecimal)
        
        strManual = Trim(txtManual)
        
        If ExecProcEditItems(iType, strErr, mlngItemID, strNO, strChina, strEnglish, _
                             strTYPE, intItemType, intResultType, strUntil, strFormu, _
                             strMethods, strDefault, intVariationAlarm, intVariationAlerts, _
                             intAlertsRate, intOutOfControlRate, intPrivacy, intMorInformation, _
                             strPY, strWb, intDecimal, intGLU, strManual) Then
            SaveItem = True
            Call RefRptItemData
        Else
            MsgBox strErr, vbQuestion, Me.Caption
            Exit Function
        End If
    End If
End Function
Private Function SvaeSequence() As Boolean
    '保存常用取值页的内容
    Dim iRow As Integer, strTmp As String, strShow As String, strQC As String, strReturnFlg As String
    Dim lngMachineID As Long, intType As Integer, strErr As String
    Dim strChannle As String, intDecimal As Integer, intGLU As Integer
    With vfgSequence
        For iRow = .FixedRows To .Rows - 1
            strTmp = Trim(.TextMatrix(iRow, .ColIndex("显示取值")))
            If strTmp <> "" Then strShow = strShow & "," & strTmp
            
            strTmp = Trim(.TextMatrix(iRow, .ColIndex("质控取值")))
            If strTmp <> "" Then strQC = strQC & "," & strTmp
            
            strTmp = Trim(.TextMatrix(iRow, .ColIndex("结果标志")))
            If strTmp <> "" Then strReturnFlg = strReturnFlg & "," & Split(strTmp, "-")(0)
        Next
        
        If strShow <> "" Then strShow = Mid(strShow, 2)
        If strQC <> "" Then strQC = Mid(strQC, 2)
        If strReturnFlg <> "" Then strReturnFlg = Mid(strReturnFlg, 2)
        lngMachineID = cboMachine1.ItemData(cboMachine1.ListIndex)
        If lngMachineID <> 0 And mlngItemID <> 0 Then
            intType = 22
        Else
            MsgBox "请给选中的项目指定仪器后再使用此功能！", vbQuestion, Me.Caption
            Exit Function
        End If
        If Not ExecProcMachineItems(intType, lngMachineID, mlngItemID, strErr, strChannle, strShow, strQC, strReturnFlg) Then
            MsgBox strErr, vbQuestion, Me.Caption
            Exit Function
        End If
        SvaeSequence = True
    End With
End Function

Private Function SvaeChannel() As Boolean
    '保存通道码页的内容
    Dim iRow As Integer, strTmp As String
    Dim lngMachineID As Long, intType As Integer, strErr As String
    Dim strChannle As String, intDecimal As Integer, intGLU As Integer
    With vfgChannel
        For iRow = .FixedRows To .Rows - 1
            strTmp = Trim(.TextMatrix(iRow, .ColIndex("通道编码")))
            If strTmp <> "" Then strChannle = strChannle & "," & strTmp
        Next
        
        If strChannle <> "" Then strChannle = Mid(strChannle, 2)
        
        lngMachineID = cboMachine2.ItemData(cboMachine2.ListIndex)
        If lngMachineID <> 0 And mlngItemID <> 0 Then
            intType = 12
        Else
            MsgBox "请给选中的项目指定仪器后再使用此功能！", vbQuestion, Me.Caption
            Exit Function
        End If
        If Not ExecProcMachineItems(intType, lngMachineID, mlngItemID, strErr, strChannle) Then
            MsgBox strErr, vbQuestion, Me.Caption
            Exit Function
        End If
        SvaeChannel = True
    End With
End Function

Private Sub ReferenceAdd()
    '添加参考
    With vfgReference
        If Not Val(.TextMatrix(.Rows - 1, .ColIndex("id"))) = 0 Then
            .Rows = .Rows + 1
        End If
        Call .Select(.Rows - 1, .ColIndex("标本"))
    End With
    
End Sub

Private Function SvaeReference() As Boolean
    '保存参考
 
    Dim lngId As Long, strErr As String
    Dim lngMachineID As Long, strSampleType As String, intSex As Integer, lngAgeLow As Long, lngAgeHigh As Long
    Dim strAgeUnitl As String, dblReferenceLow As Double, dblReferenceHigh As Double, strFeatures As String
    Dim dblAbnormaLow As Double, dblAbnormaHigh As Double, dblReviewLow As Double, dblReviewHigh As Double
    Dim intDefault As Integer
    Dim strReferenceShow As String
    
    If mlngItemID = 0 Then
        MsgBox "请选中一个指标后再执行此操作!", vbQuestion, Me.Caption
        Exit Function
    End If

    With vfgReference
        If .Rows > .FixedRows Then
            If .Row >= .FixedRows And .Row <= .Rows - 1 Then
                lngId = Val(.TextMatrix(.Row, .ColIndex("ID")))
            End If
        End If
    End With

    
    lngMachineID = cboMachine3.ItemData(cboMachine3.ListIndex)
    If cboSampleType.ListIndex = 0 Then
        strSampleType = ""
    Else
        strSampleType = Split(cboSampleType.List(cboSampleType.ListIndex), "-")(1)
    End If
    intSex = Val(Split(cboSex.List(cboSex.ListIndex), "-")(0))
    lngAgeLow = Val(txtAge(0))
    lngAgeHigh = Val(txtAge(1))
    strAgeUnitl = cboAgeUnitl.List(cboAgeUnitl.ListIndex)
    dblReferenceLow = Val(txtReference(0))
    dblReferenceHigh = Val(txtReference(1))
    strReferenceShow = StringDelInvalidWord(txtReferenceShow)
    strFeatures = Split(cboFeatures.List(cboFeatures.ListIndex), "-")(1)
    dblAbnormaLow = Val(txtAbnorma(0))
    dblAbnormaHigh = Val(txtAbnorma(1))
    dblReviewLow = Val(txtReview(0))
    dblReviewHigh = Val(txtReview(1))
    
    intDefault = chkDefault.Value
    If Not ExecProcReference(2, lngId, strErr, mlngItemID, lngMachineID, strSampleType, intSex, lngAgeLow, lngAgeHigh, strAgeUnitl, _
           dblReferenceLow, dblReferenceHigh, strReferenceShow, strFeatures, dblAbnormaLow, dblAbnormaHigh, dblReviewLow, dblReviewHigh, intDefault) Then
        MsgBox strErr, vbQuestion, Me.Caption
        
        Exit Function
    Else
        SvaeReference = True
        
    End If
        
    lblReference.Tag = ""
End Function

Private Sub ReferenceDel()
    '删除一行参考
    Dim lngId As Long, strErr As String
    With vfgReference
        If .Rows > .FixedRows Then
            If .Row >= .FixedRows And .Row <= .Rows - 1 Then
                lngId = Val(.TextMatrix(.Row, .ColIndex("ID")))
                If lngId <> 0 Then
                    If Not ExecProcReference(3, lngId, strErr) Then
                        MsgBox strErr, vbQuestion, Me.Caption
                        Exit Sub
                    End If
                End If
                .RemoveItem .Row
            End If
        End If
    End With
End Sub

Private Sub FillReferenceShow()
    '当参考显示为空时,根据参考高值,低值填充参考显示,
    If Trim(txtReferenceShow) = "" Then
        If Trim(txtReference(0)) <> "" And Trim(txtReference(1)) <> "" Then
            txtReferenceShow = Trim(txtReference(0)) & "～" & Trim(txtReference(1))
        ElseIf Trim(txtReference(0)) <> "" Then
            txtReferenceShow = "≤" & Trim(txtReference(0))
        ElseIf Trim(txtReference(1)) <> "" Then
            txtReferenceShow = "≥" & Trim(txtReference(1))
        End If
    End If
End Sub
'----------------------------------------------------------------------------------
'--- 以下为本窗体控件过程
'----------------------------------------------------------------------------------

Private Sub cbsSub3_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        '参考页的编辑功能
        Case conMenu_Edit_ItemEdit  '修改
            Call LockEdit(3, True)
        Case conMenu_Edit_ItemAdd   '新增
            Call LockEdit(3, True)
            Call ReferenceAdd
        Case conMenu_Edit_ItemDele  '删除
            If MsgBox("是否删除参考！", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                Call ReferenceDel
            End If
        Case conMenu_Edit_ItemUndo '取消
            Call LockEdit(3, False)
            Call RptItem_SelectionChanged
        Case conMenu_Edit_ItemSave '保存
            If SvaeReference Then
                Call LockEdit(3, False)
                Call RptItem_SelectionChanged
            End If
    End Select
End Sub

Private Sub cbsSub3_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '参考
    Select Case Control.ID
        Case conMenu_Edit_ItemEdit
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemAdd
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemDele
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemUndo
            Control.Enabled = mintEditIndex = 4
        Case conMenu_Edit_ItemSave
            Control.Enabled = mintEditIndex = 4
    End Select
End Sub

Private Sub cboMachine1_Click()
    Call RefPicinfo1    '刷新数据
End Sub

Private Sub cboMachine2_Click()
    Call RefPicinfo2    '刷新通道码数据
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Add               '仪器设置
        frmBaseAddMachine.Show vbModal
        LoadBaseData
    Case conMenu_Edit_BillSet           '单据设置
        frmBaseSetBill.Show vbModal
        LoadBaseData
    Case conMenu_Edit_Refresh
        Call RefRptItemData
    Case conMenu_Edit_Find
        Call Find
    Case conMenu_Edit_ItemSort          '项目顺序调整
        frmBaseItemSort.Show vbModal
    Case conMenu_Edit_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call CbsResize2(Me.cbsMain, Me, Me.picLeft, Me.FraWE, Me.tabBase, mminWidth, True)
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Add               '仪器设置
        Control.Enabled = Not (mintEditIndex <> 0) And InStr(mstrPrivs, ";仪器管理;") > 0
    Case conMenu_Edit_BillSet           '单据设置
        Control.Enabled = Not (mintEditIndex <> 0) And InStr(mstrPrivs, ";单据管理;") > 0
    Case conMenu_Edit_Refresh
        Control.Enabled = Not (mintEditIndex <> 0)
    Case conMenu_Edit_Find
        Control.Enabled = Not (mintEditIndex <> 0)
    Case conMenu_Edit_ItemSort          '项目顺序调整
        Control.Enabled = Not (mintEditIndex <> 0)
    Case conMenu_Edit_Exit
        Control.Enabled = Not (mintEditIndex <> 0)
    End Select
End Sub

Private Sub cbsSub0_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        '项目基本属性页的编辑功能
        Case conMenu_Edit_ItemAdd   '新增
            mintType = 1
            Call LockEdit(0, True)
            ClearPicInfo0
        Case conMenu_Edit_ItemEdit  '修改
            mintType = 2
            Call LockEdit(0, True)
        Case conMenu_Edit_ItemDele  '删除
            If MsgBox("是否删除当前指标？", vbQuestion + vbOKCancel + vbDefaultButton2, Me.Caption) = vbOK Then
                Call SaveItem(3)
            End If
        Case conMenu_Edit_ItemUndo  '取消
            mintType = 0
            Call LockEdit(0, False)
            Call RptItem_SelectionChanged
        Case conMenu_Edit_ItemSave  '保存
            If SaveItem(mintType) Then
                mintType = 0
                Call LockEdit(0, False)
            End If
    End Select
End Sub

Private Sub cbsSub0_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call cbsSub0.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With picSub0
        .Left = lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
        .Width = lngRight - lngLeft
    End With
End Sub

Private Sub cbsSub0_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
        Case conMenu_Edit_ItemAdd
            Control.Enabled = mintEditIndex = 0
        Case conMenu_Edit_ItemEdit
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemDele
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemSave, conMenu_Edit_ItemUndo '取消，保存
            Control.Enabled = mintEditIndex = 1
    End Select
End Sub

Private Sub cbsSub1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        '常用取值页的编辑功能
        Case conMenu_Edit_ItemEdit
            Call LockEdit(1, True)
        Case conMenu_Edit_ItemAdd   '新增
            vfgSequence.Rows = vfgSequence.Rows + 1
        Case conMenu_Edit_ItemDele  '删除
            With vfgSequence
                If .Rows > .FixedRows Then
                    If .Row >= .FixedRows And .Row <= .Rows - 1 Then
                        .RemoveItem .Row
                    End If
                End If
            End With
        Case conMenu_Edit_ItemUndo '取消
            Call LockEdit(1, False)
            Call RptItem_SelectionChanged
        Case conMenu_Edit_ItemSave '保存
            If SvaeSequence Then
                Call LockEdit(1, False)
                Call RptItem_SelectionChanged
            End If
    End Select
End Sub

Private Sub cbsSub1_Resize()
    Call cbsSubResize(Me.vfgSequence, cbsSub1)
    On Error Resume Next
    
    With fraBase
        .Left = vfgSequence.Left + 45
        .Top = vfgSequence.Top + 45
    End With
 
    With vfgSequence
        .Height = .Height - fraBase.Height - 90
        .Top = fraBase.Top + fraBase.Height + 45
    End With

End Sub

Private Sub cbsSub1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim iDataType As Integer
    
    If cboDataType.ListCount > 0 Then iDataType = Val(cboDataType.List(cboDataType.ListIndex))
    
    Select Case Control.ID
        Case conMenu_Edit_ItemEdit
            Control.Enabled = mintEditIndex = 0 And iDataType <> 1 And mlngItemID <> 0
        Case conMenu_Edit_ItemAdd
            Control.Enabled = mintEditIndex = 2
        Case conMenu_Edit_ItemDele
            Control.Enabled = mintEditIndex = 2
        Case conMenu_Edit_ItemUndo
            Control.Enabled = mintEditIndex = 2
        Case conMenu_Edit_ItemSave
            Control.Enabled = mintEditIndex = 2
    End Select
End Sub

Private Sub cbsSub2_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        '通道码页的编辑功能
        Case conMenu_Edit_ItemEdit
            Call LockEdit(2, True)
        Case conMenu_Edit_ItemAdd   '新增
            vfgChannel.Rows = vfgChannel.Rows + 1
        Case conMenu_Edit_ItemDele  '删除
            With vfgChannel
                If .Rows > .FixedRows Then
                    If .Row >= .FixedRows And .Row <= .Rows - 1 Then
                        .RemoveItem .Row
                    End If
                End If
            End With
        Case conMenu_Edit_ItemUndo '取消
            Call LockEdit(2, False)
            Call RptItem_SelectionChanged
        Case conMenu_Edit_ItemSave '保存
            If SvaeChannel Then
                Call LockEdit(2, False)
                Call RptItem_SelectionChanged
            End If
    End Select

End Sub

Private Sub cbsSub2_Resize()
    Call cbsSubResize(Me.vfgChannel, cbsSub2)
    On Error Resume Next
    
    With fraSub2
        .Left = vfgChannel.Left + 45
        .Top = vfgChannel.Top + 45
    End With
 
    With vfgChannel
        .Height = .Height - fraSub2.Height - 90
        .Top = fraSub2.Top + fraSub2.Height + 45
    End With
End Sub

Private Sub cbsSub2_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
        Case conMenu_Edit_ItemEdit
            Control.Enabled = mintEditIndex = 0 And mlngItemID <> 0
        Case conMenu_Edit_ItemAdd
            Control.Enabled = mintEditIndex = 3 And mlngItemID <> 0
        Case conMenu_Edit_ItemDele
            Control.Enabled = mintEditIndex = 3 And mlngItemID <> 0
        Case conMenu_Edit_ItemUndo
            Control.Enabled = mintEditIndex = 3
        Case conMenu_Edit_ItemSave
            Control.Enabled = mintEditIndex = 3
    End Select
End Sub

Private Sub cbsSub3_Resize()
    
    Call cbsSubResize(Me.vfgReference, cbsSub3)
    On Error Resume Next
    
    With picSub3
        .Left = vfgReference.Left + 45
        .Top = vfgReference.Top + 45
    End With
 
    With vfgReference
        .Height = .Height - picSub3.Height - 90
        .Top = picSub3.Top + picSub3.Height + 45
    End With
End Sub

Private Sub Form_Load()
    Dim strTitle As String, strErr As String, i As Integer
    '初始化各区域宽高
    mstrPrivs = ComGetPrivs(gSysInfo.SysNo, 1002)
    If Left(mstrPrivs, 1) <> ";" And mstrPrivs <> "" Then mstrPrivs = ";" & mstrPrivs
    If Right(mstrPrivs, 1) <> ";" And mstrPrivs <> "" Then mstrPrivs = mstrPrivs & ";"
    
    mminHeight = 4200: mminWidth = Me.picInfo(0).Width
    
    '美化pic控件边框
    Call setBorderColor(picLeft.hwnd)
    For i = picInfo.LBound To picInfo.UBound
        Call setBorderColor(picInfo(i).hwnd)
    Next
    
    lblHelp.Caption = "     1.警示与比对信息填入百分比数字，例如5%填入5即可。" & vbNewLine & _
                      "     2.带*号的项目为必须输入的项目，其中指标代码的字母，统一保存为大写。" & vbNewLine & _
                      "     3.仅定量的计算项目可设置计算公式，计算公式中不包含本项目及其他计算项目。"
    '初始化主菜单
    Call cbsMainInit
    '初始化项目属性页
    picInfo(0).Tag = "基本信息"
    picInfo(1).Tag = "常用取值"
    picInfo(2).Tag = "通道码"
    picInfo(3).Tag = "项目参考"
    Set mPages = New Collection
    mPages.Add picInfo(0), "_基本信息"
    mPages.Add picInfo(1), "_常用取值"
    mPages.Add picInfo(2), "_通道码"
    mPages.Add picInfo(3), "_项目参考"
    Call TabSetting(mPages, Me.tabBase)
    
    '初始化分页中的菜单
    '基本信息页
    Call cbsSubInit(cbsSub0)
    '常用取值页
    Call cbsSubInit1(cbsSub1)
    Call vfgSetting(0, vfgSequence)
    
    '通道码页
    Call cbsSubInit1(cbsSub2)
    Call vfgSetting(0, vfgChannel)
    
    '参考值
    Call cbsSubInit(cbsSub3)
    vfgReference.Tag = ""
    Call vfgSetting(0, vfgReference)
    '加载数据
    Call LoadBaseData
    Call RefRptItemData         '装载检验指标数据
    
    Call LockEdit(-1, False)     '初始化界面锁定
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub picSub0_Resize()
    On Error Resume Next
    With txtManual
        .Height = picSub0.ScaleHeight - lblHelp.Height - txtManual.Top - 90
        
        lblHelp.Top = .Top + .Height + 45
        imgNote(1).Top = lblHelp.Top
    End With
    
End Sub

Private Sub txtAbnorma_GotFocus(Index As Integer)
    txtAbnorma(Index).SelStart = 0: txtAbnorma(Index).SelLength = Len(txtAbnorma(Index))
End Sub

Private Sub txtAbnorma_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtAge_GotFocus(Index As Integer)
    txtAge(Index).SelStart = 0: txtAge(Index).SelLength = Len(txtAge(Index))
End Sub

Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtAlertsRate_GotFocus()
    txtAlertsRate.SelStart = 0: txtAlertsRate.SelLength = Len(txtAlertsRate)
End Sub

Private Sub txtAlertsRate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtChineseName_GotFocus()
    txtChineseName.SelStart = 0: txtChineseName.SelLength = Len(txtChineseName)
End Sub

Private Sub txtChineseName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtChineseName_Validate(Cancel As Boolean)
    If txtChineseName <> "" Then
        If txtPYCode = "" Then txtPYCode = ComGetSymbol(txtChineseName)
        If txtWBCode = "" Then txtWBCode = ComGetSymbol(txtChineseName, 1)
    End If
End Sub

Private Sub txtDecimal_GotFocus()
    txtDecimal.SelStart = 0: txtDecimal.SelLength = Len(txtDecimal)
End Sub

Private Sub txtDecimal_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtDefault_GotFocus()
    txtDefault.SelStart = 0: txtDefault.SelLength = Len(txtDefault)
End Sub

Private Sub txtDefault_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtEnglish_GotFocus()
    If txtEnglish.Text = "" Then txtEnglish.Text = txtNo.Text
    txtEnglish.SelStart = 0: txtEnglish.SelLength = Len(txtEnglish)
End Sub

Private Sub txtEnglish_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtManual_GotFocus()
    txtManual.SelStart = 0: txtManual.SelLength = Len(txtManual)
End Sub

Private Sub txtManual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtNo_GotFocus()
    txtNo.SelStart = 0: txtNo.SelLength = Len(txtNo)
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtOutOfControlRate_GotFocus()
    txtOutOfControlRate.SelStart = 0: txtOutOfControlRate.SelLength = Len(txtOutOfControlRate)
End Sub

Private Sub txtOutOfControlRate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtPYCode_GotFocus()
    txtPYCode.SelStart = 0: txtPYCode.SelLength = Len(txtPYCode)
End Sub

Private Sub txtPYCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtReference_GotFocus(Index As Integer)
    txtReference(Index).SelStart = 0: txtReference(Index).SelLength = Len(txtReference(Index))
End Sub

Private Sub txtReference_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtReference_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then Call FillReferenceShow
End Sub

Private Sub txtReferenceShow_GotFocus()
    Call FillReferenceShow
    txtReferenceShow.SelStart = 0: txtReferenceShow.SelLength = Len(txtReferenceShow)
End Sub

Private Sub txtReferenceShow_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtReview_GotFocus(Index As Integer)
    txtReview(Index).SelStart = 0: txtReview(Index).SelLength = Len(txtReview(Index))
End Sub

Private Sub txtReview_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtTestMethods_GotFocus()
    txtTestMethods.SelStart = 0: txtTestMethods.SelLength = Len(txtTestMethods)
End Sub

Private Sub txtTestMethods_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtType_GotFocus()
txtType.SelStart = 0: txtType.SelLength = Len(txtType)
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtUnits_GotFocus()
    txtUnits.SelStart = 0: txtUnits.SelLength = Len(txtUnits)
End Sub

Private Sub txtUnits_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtVariationAlarm_GotFocus()
    txtVariationAlarm.SelStart = 0: txtVariationAlarm.SelLength = Len(txtVariationAlarm)
End Sub

Private Sub txtVariationAlarm_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtVariationAlerts_GotFocus()
    txtVariationAlerts.SelStart = 0: txtVariationAlerts.SelLength = Len(txtVariationAlerts)
End Sub

Private Sub txtVariationAlerts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub txtWBCode_GotFocus()
    txtWBCode.SelStart = 0: txtWBCode.SelLength = Len(txtWBCode)
End Sub

Private Sub txtWBCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab)
End Sub

Private Sub vfgReference_RowColChange()
    '更新显示
    Dim strTmp As String
    With vfgReference
        If .Row >= .FixedRows And .Row <= .Rows - 1 Then
            If Not .Tag = "Ok" Then Exit Sub    '已初始化了标题才执行后续操作
            
            strTmp = Trim("" & .TextMatrix(.Row, .ColIndex("标本")))
            cboSelect cboSampleType, strTmp, 1
            strTmp = Trim("" & .TextMatrix(.Row, .ColIndex("年龄单位")))
            cboSelect cboSampleType, strTmp
            strTmp = Trim("" & .TextMatrix(.Row, .ColIndex("性别")))
            cboSelect cboSex, strTmp
            strTmp = Trim("" & .TextMatrix(.Row, .ColIndex("临床特征")))
            cboSelect cboFeatures, strTmp, 1
            strTmp = Trim("" & .TextMatrix(.Row, .ColIndex("仪器")))
            cboSelect cboFeatures, strTmp, 1
            
            txtAge(0) = Trim("" & .TextMatrix(.Row, .ColIndex("年龄下限")))
            txtAge(1) = Trim("" & .TextMatrix(.Row, .ColIndex("年龄上限")))
            txtReference(0) = Trim("" & .TextMatrix(.Row, .ColIndex("参考低值")))
            txtReference(1) = Trim("" & .TextMatrix(.Row, .ColIndex("参考高值")))
            txtReferenceShow = Trim("" & .TextMatrix(.Row, .ColIndex("参考值")))
            chkDefault.Value = Val(Trim("" & .TextMatrix(.Row, .ColIndex("默认"))))
            
            txtAbnorma(0) = Trim("" & .TextMatrix(.Row, .ColIndex("危急低值")))
            txtAbnorma(1) = Trim("" & .TextMatrix(.Row, .ColIndex("危急高值")))
            txtReview(0) = Trim("" & .TextMatrix(.Row, .ColIndex("复查下限")))
            txtReview(1) = Trim("" & .TextMatrix(.Row, .ColIndex("复查上限")))
        End If
    End With
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picLeft
        Rightcoll.Add Me.tabBase
        Call SplitWE(LeftColl, Me.FraWE, Rightcoll, X, mminWidth)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub



