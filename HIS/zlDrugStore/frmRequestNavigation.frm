VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequestNavigation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品申领管理自动生成向导"
   ClientHeight    =   7605
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmRequestNavigation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8145
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7170
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":1582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicSetup 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   -15
      Width           =   1485
      Begin VB.Image imgSetup 
         Height          =   6645
         Left            =   60
         Picture         =   "frmRequestNavigation.frx":328C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5265
      TabIndex        =   1
      Top             =   7095
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   330
      TabIndex        =   2
      Top             =   7095
      Width           =   1230
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6630
      TabIndex        =   0
      Top             =   7095
      Width           =   1230
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
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
            Picture         =   "frmRequestNavigation.frx":8872
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":974C
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":9B9E
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":9FF0
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   -120
      Width           =   6555
      Begin VB.CheckBox chk检查库存 
         Caption         =   "出库库房无可用库存时不产生申领记录"
         Height          =   180
         Left            =   360
         TabIndex        =   59
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "申领方式"
         Height          =   4600
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   6135
         Begin VB.CheckBox chkLowerLimit 
            Caption         =   "固定申领上限数量"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   2880
            TabIndex        =   51
            Top             =   2085
            Width           =   1875
         End
         Begin VB.OptionButton optMode 
            Caption         =   "4、按药品的储备上限，下限综合考虑"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   45
            Top             =   2655
            Width           =   3285
         End
         Begin VB.CheckBox chk上限 
            Caption         =   "固定申领上限数量"
            Enabled         =   0   'False
            Height          =   225
            Left            =   690
            TabIndex        =   43
            Top             =   1305
            Width           =   1875
         End
         Begin VB.CheckBox chkLowerLimit 
            Caption         =   "固定申领下限数量"
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   690
            TabIndex        =   42
            Top             =   2085
            Width           =   1875
         End
         Begin VB.OptionButton optMode 
            Caption         =   "6、计算日销售量"
            Height          =   195
            Index           =   5
            Left            =   330
            TabIndex        =   15
            Top             =   3960
            Width           =   3255
         End
         Begin VB.CheckBox chk仅申领库存小于消耗量 
            Caption         =   "仅申领库存小于消耗量的药品"
            Height          =   225
            Left            =   690
            TabIndex        =   36
            Top             =   480
            Width           =   2715
         End
         Begin VB.OptionButton optMode 
            Caption         =   "5、根据指定时间范围内的申领单"
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   14
            Top             =   3390
            Width           =   3285
         End
         Begin VB.OptionButton optMode 
            Caption         =   "3、按药品的储备下限"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   13
            Top             =   1875
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "2、按药品的储备上限"
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   12
            Top             =   1080
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "1、根据指定时间范围内药品的消耗量"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   3405
         End
         Begin VB.OptionButton optMode 
            Caption         =   "7、根据指定时间范围内药品的销售量"
            Height          =   180
            Index           =   6
            Left            =   330
            TabIndex        =   52
            Top             =   240
            Width           =   3405
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "提取库存低于下限的药品，并使当前库房的储备量始终保持在上限标准，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   46
            Top             =   2880
            Width           =   5460
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据指定时间范围内的日销售量，以及设定的库存上限、库存下限天数，计算后产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   510
            Index           =   5
            Left            =   390
            TabIndex        =   20
            Top             =   4200
            Width           =   5400
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据指定时间范围内的申领单的未发数量，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   4
            Left            =   390
            TabIndex        =   19
            Top             =   3630
            Width           =   4680
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使当前库房的药品储备量始终保持在下限标准，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   18
            Top             =   2340
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使当前库房的药品储备量始终保持在上限标准，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   1560
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据您指定的时间范围，以药品的消耗量为依据，产生本次的申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   5400
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据您指定的时间范围，以药品的销售（发药）量为依据，产生本次的申领单。注意：本次申领单的起始时间为上个申领单的截止时间."
            ForeColor       =   &H00004000&
            Height          =   780
            Index           =   6
            Left            =   360
            TabIndex        =   53
            Top             =   480
            Width           =   5295
         End
      End
      Begin VB.CheckBox chk申领数量 
         Caption         =   "申领数量作为参考数量"
         Height          =   180
         Left            =   2880
         TabIndex        =   54
         Top             =   1750
         Width           =   2415
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "非常备药品"
         Height          =   180
         Index           =   2
         Left            =   3600
         TabIndex        =   49
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "常备药品"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   48
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "全部药品"
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   47
         Top             =   1440
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chk整数 
         Caption         =   "确保申领数量为整数"
         Height          =   180
         Left            =   360
         TabIndex        =   44
         Top             =   1750
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   6255
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   2235
      End
      Begin VB.Label lblDrugType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "选择药品"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   50
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "第一步：决定产生申领单的方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "你准备向哪个库房发生申领请求"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   780
         Width           =   2730
      End
      Begin VB.Label lbl库房 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   2
      Left            =   1470
      TabIndex        =   37
      Top             =   -120
      Width           =   6555
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Width           =   6255
      End
      Begin MSComctlLib.TreeView tvw用途 
         Height          =   5370
         Left            =   150
         TabIndex        =   39
         Top             =   1035
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   9472
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "第三步：指定药品分类以缩小范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   41
         Top             =   240
         Width           =   6300
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品分类选择选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   40
         Top             =   810
         Width           =   1560
      End
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   1
      Left            =   1470
      TabIndex        =   21
      Top             =   -120
      Width           =   6555
      Begin VB.CheckBox chk材质中草药 
         Caption         =   "中草药"
         Height          =   180
         Left            =   3600
         TabIndex        =   57
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk材质中成药 
         Caption         =   "中成药"
         Height          =   180
         Left            =   2460
         TabIndex        =   56
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk材质西成药 
         Caption         =   "西成药"
         Height          =   180
         Left            =   1320
         TabIndex        =   55
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt下限天数 
         Height          =   300
         Left            =   5370
         TabIndex        =   33
         Top             =   3090
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt上限天数 
         Height          =   300
         Left            =   5370
         TabIndex        =   31
         Top             =   2700
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   285
         Left            =   4170
         TabIndex        =   26
         Top             =   1350
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   158400515
         CurrentDate     =   38096
      End
      Begin MSComctlLib.ListView lvwSelect 
         Height          =   4845
         Left            =   150
         TabIndex        =   24
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8546
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   285
         Left            =   4170
         TabIndex        =   28
         Top             =   1980
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   158400515
         CurrentDate     =   38096
      End
      Begin VB.Label lbl材质分类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "材质分类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   58
         Top             =   6120
         Width           =   780
      End
      Begin VB.Label lbl下限天数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "下限天数(&T)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   32
         Top             =   3150
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl上限天数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "上限天数(&X)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl库存限额条件 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "库存限额条件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   29
         Top             =   2460
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbl其它条件设置 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其它条件设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   35
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lbl剂型选择 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "剂型选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lbl结束时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   27
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lbl开始时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   25
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "第二步：指定剂型以缩小范围（对中草药无效）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   23
         Top             =   240
         Width           =   6300
      End
   End
End
Attribute VB_Name = "frmRequestNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 模式
    根据消耗量 = 0
    根据上限 = 1
    根据下限 = 2
    根据上下限 = 3
    根据申领单未发数 = 4
    根据日销售量
    根据销售总量
End Enum
Private mstr剂型 As String
Private mblnOK As Boolean
Private mlngStockID As Long                 '申领库房ID
Private mintAutoType As Integer             '自动申领时，申领方式：1-根据消耗量;2-根据上限;3-根据下限;4-根据上下限;5-根据申领单未发数;6-根据日销售量;7-根据销售总量
Private mintCheck As Integer                '库存检查参数
Private mblnFirst  As Boolean
Private mblnStart As Boolean
Private mfrmMain As Object
Private mintStep As Integer
Private mIntCol申领数量 As Integer
Private mIntCol填写数量 As Integer
Private mstr库房id As String
Private mstr材质分类 As String          '用来记录选择的材质分类
Private mstr分类 As String      '5-西药，6-成药 7-草药

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数
Private mintUnit As Integer                 '单位
Private mint按批次出库 As Integer           '0-不按批次出库,1-按批次出库
Private Sub CheckAll(ByVal myNodes As Nodes, ByVal blnCheck As Boolean)
    Dim tmpNode As Node
    
    For Each tmpNode In myNodes
        tmpNode.Checked = True
        If tmpNode.Child > 0 Then
            Call CheckAll(tmpNode, blnCheck)
        End If
    Next
End Sub

Private Function Ini药品分类() As Boolean
    '药品用途分类
    Dim lng分类id As Long
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node
    
    mstr材质分类 = ""
    If chk材质西成药.Value = 1 Then
        mstr材质分类 = "1"
        mstr分类 = "5"
    End If
    
    If mstr材质分类 <> "" Then
        If chk材质中成药.Value = 1 And chk材质中成药.Visible = True Then
            mstr材质分类 = mstr材质分类 & ",2"
            mstr分类 = mstr分类 & ",6"
        End If
    Else
        If chk材质中成药.Value = 1 And chk材质中成药.Visible = True Then
            mstr材质分类 = "2"
            mstr分类 = "6"
        End If
    End If
    
    If mstr材质分类 <> "" Then
        If chk材质中草药.Value = 1 And chk材质中草药.Visible = True Then
            mstr材质分类 = mstr材质分类 & ",3"
            mstr分类 = mstr分类 & ",7"
        End If
    Else
        If chk材质中草药.Value = 1 And chk材质中草药.Visible = True Then
            mstr材质分类 = "3"
            mstr分类 = "7"
        End If
    End If
    
    If mstr材质分类 = "" Then
        MsgBox "请选择材质分类！", vbInformation, gstrSysName
        Ini药品分类 = False
        Exit Function
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select Level As 层, a.ID, a.上级id, a.名称, Decode(a.类型, 1, '西成药', 2, '中成药', '中草药') As 材质" & _
                " From 诊疗分类目录 a" & _
                " Where a.类型 in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList)))" & _
                " Start With a.上级id Is Null" & _
                " Connect By Prior a.ID = a.上级id" & _
                " Order By Level"

    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "药品用途分类", mstr材质分类)

    tvw用途.Nodes.Clear
    Set objNode = tvw用途.Nodes.Add(, , "Root", "所有分类", "Item")
    
    If InStr(mstr材质分类, "1") > 0 Then
        Set objNode = tvw用途.Nodes.Add("Root", 4, "_西成药", "西成药", "Item")
    End If
    If InStr(mstr材质分类, "2") > 0 Then
        Set objNode = tvw用途.Nodes.Add("Root", 4, "_中成药", "中成药", "Item")
    End If
    If InStr(mstr材质分类, "3") > 0 Then
        Set objNode = tvw用途.Nodes.Add("Root", 4, "_中草药", "中草药", "Item")
    End If
    
    Do While Not rsTmp.EOF
        If rsTmp!层 = 1 Then
            Set objNode = tvw用途.Nodes.Add("_" & rsTmp!材质, 4, "_" & rsTmp!Id, rsTmp!名称, "Item")
        Else
            Set objNode = tvw用途.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!Id, rsTmp!名称, "Item")
        End If
        rsTmp.MoveNext
    Loop
    tvw用途.Nodes("Root").Selected = True
    tvw用途.Nodes("Root").Expanded = True
    
    Ini药品分类 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RunByStep(ByVal intStep As Integer)
    Dim bln材料分类 As Boolean
    
    Select Case intStep
        Case 0  '第一步
            fraStep(0).Visible = True
            fraStep(0).ZOrder
            
            fraStep(1).Visible = False
            fraStep(2).Visible = False
            
            cmdPrevious.Enabled = False
            cmdNext.Caption = "下一步(&N)"
        Case 1  '第二步
            fraStep(1).Visible = True
            fraStep(1).ZOrder
            
            fraStep(0).Visible = False
            fraStep(2).Visible = False
            
            cmdPrevious.Enabled = True
            cmdNext.Caption = "下一步(&N)"
            
            '调整剂型的位置
            Call ResizeDrug
            '调整材料分类位置
            If chk材质西成药.Visible = False And chk材质中成药.Visible = True Then
                chk材质中成药.Left = chk材质西成药.Left
            End If
            If chk材质西成药.Visible = True And chk材质中成药.Visible = False And chk材质中草药.Visible = True Then
                chk材质中草药.Left = chk材质中成药.Left
            End If
            If chk材质西成药.Visible = False And chk材质中成药.Visible = False And chk材质中草药.Visible = True Then
                chk材质中草药.Left = chk材质西成药.Left
            End If
            
        Case 2  '第三步
            bln材料分类 = Ini药品分类
            If bln材料分类 = True Then
                fraStep(2).Visible = True
                fraStep(2).ZOrder
                
                fraStep(0).Visible = False
                fraStep(1).Visible = False
                
                cmdPrevious.Enabled = True
                cmdNext.Caption = "完成(&F)"
            Else
                Exit Sub
            End If
        Case 3  '完成
            If optMode(根据日销售量) Then
                '库存上限天数不能小于库存下限天数
                '库存上限天数与库存下限天数不能为零
                If dtp开始时间.Value > dtp结束时间.Value Then
                     MsgBox "开始时间不能大于结束时间", vbInformation, gstrSysName
                     Call RunByStep(1)
                     dtp开始时间.SetFocus
                     Exit Sub
                End If
                
                If Trim(txt上限天数.Text) = "" Then
                    MsgBox "请输入库存上限天数！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt上限天数.SetFocus
                    Exit Sub
                End If
                If Trim(txt下限天数.Text) = "" Then
                    MsgBox "请输入库存下限天数！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt下限天数.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(txt上限天数.Text) Then
                    MsgBox "库存上限天数中含有非法字符！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt上限天数.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(txt下限天数.Text) Then
                    MsgBox "库存下限天数中含有非法字符！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt下限天数.SetFocus
                    Exit Sub
                End If
                If Val(txt上限天数.Text) <= 0 Then
                    MsgBox "库存上限天数不能小于零！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt上限天数.SetFocus
                    Exit Sub
                End If
                If Val(txt下限天数.Text) <= 0 Then
                    MsgBox "库存下限天数不能小于零！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt下限天数.SetFocus
                    Exit Sub
                End If
                If Val(txt上限天数.Text) < Val(txt下限天数.Text) Then
                    MsgBox "库存上限天数不能小于库存下限天数！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt上限天数.SetFocus
                    Exit Sub
                End If
                If Val(txt上限天数.Text) > 300 Then
                    MsgBox "库存上限天数不能大于300天！", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt上限天数.SetFocus
                    Exit Sub
                End If
            ElseIf optMode(根据销售总量) Then
                If dtp开始时间.Value > dtp结束时间.Value Then
                    MsgBox "开始时间不能大于结束时间", vbInformation, gstrSysName
                    Call RunByStep(1)
                    If dtp开始时间.Enabled = True Then
                        dtp开始时间.SetFocus
                    Else
                        dtp结束时间.SetFocus
                    End If
                    Exit Sub
                End If
            End If
            
            '产生数据
            Call Get剂型串
            If Not CheckData Then Exit Sub
            
            mblnOK = True
            Unload Me
    End Select
End Sub

'Private Sub cbo材质分类_Click()
'    If cbo材质分类.ListIndex >= 0 Then
'        Call Ini药品分类(cbo材质分类.ItemData(cbo材质分类.ListIndex))
'    End If
'End Sub


Private Sub chkLowerLimit_Click(index As Integer)
    If chkLowerLimit(index).Value = 1 Then
        chkLowerLimit(Abs(index - 1)).Value = 0
    End If
    
    If chkLowerLimit(0).Value = 1 Then
        lblTip(2).Caption = "使当前库房的药品储备量至少保持在下限标准，产生本次申领单"
    ElseIf chkLowerLimit(1).Value = 1 Then
        lblTip(2).Caption = "使当前库房的药品储备量至少保持在上限标准，产生本次申领单"
    Else
        lblTip(2).Caption = "使当前库房的药品储备量始终保持在下限标准，产生本次申领单"
    End If
End Sub

Private Sub chk上限_Click()
    If chk上限.Value = 1 Then
        lblTip(1).Caption = "使当前库房的药品储备量至少保持在上限标准，产生本次申领单"
    Else
        lblTip(1).Caption = "使当前库房的药品储备量始终保持在上限标准，产生本次申领单"
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    mintStep = IIf(mintStep = 3, 3, mintStep + 1)
    Call RunByStep(mintStep)
End Sub

Private Sub cmdPrevious_Click()
    mintStep = IIf(mintStep = 0, 0, mintStep - 1)
    Call RunByStep(mintStep)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    '----缺省选中所有剂型----
    
    If Not mblnFirst Then Exit Sub
    
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
    
    If mintAutoType <> 7 Then
        optMode(0).Visible = True
        chk仅申领库存小于消耗量.Visible = True
        lblTip(0).Visible = True
        
        optMode(1).Visible = True
        chk上限.Visible = True
        lblTip(1).Visible = True
        
        optMode(2).Visible = True
        chkLowerLimit(0).Visible = True
        chkLowerLimit(1).Visible = True
        lblTip(2).Visible = True
        
        optMode(3).Visible = True
        lblTip(3).Visible = True
        
        optMode(4).Visible = True
        lblTip(4).Visible = True
        
        optMode(5).Visible = True
        lblTip(5).Visible = True
        
        optMode(6).Visible = False
        lblTip(6).Visible = False
    Else
        optMode(0).Visible = False
        chk仅申领库存小于消耗量.Visible = False
        lblTip(0).Visible = False
        
        optMode(1).Visible = False
        chk上限.Visible = False
        lblTip(1).Visible = False
        
        optMode(2).Visible = False
        chkLowerLimit(0).Visible = False
        chkLowerLimit(1).Visible = False
        lblTip(2).Visible = False
        
        optMode(3).Visible = False
        lblTip(3).Visible = False
        
        optMode(4).Visible = False
        lblTip(4).Visible = False
        
        optMode(5).Visible = False
        lblTip(5).Visible = False
        
        optMode(6).Visible = True
        lblTip(6).Visible = True
        optMode(6).Value = True
    End If
    
    Call RunByStep(0)
    lvwSelect.ListItems(1).Checked = True
    Call lvwSelect_ItemCheck(lvwSelect.ListItems(1))
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim dateCurDate As Date
    Dim intStock As Integer
    Dim str站点权限 As String
    Dim int申领数量 As Integer
    Dim int整数 As Integer
    Dim int检查可用数量 As Integer
    
    mblnStart = False
    mblnFirst = True
    mintStep = 0
    
    int申领数量 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "申领数量作为参考数量", 0)))
    int整数 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "确保申领数量为整数", 0)))
    int检查可用数量 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "检查可用数量", 0)))
    mint按批次出库 = Val(zlDataBase.GetPara("药品按批次出库", glngSys, 1343, 0))
    
    Me.chk申领数量.Value = IIf(int申领数量 = 1, 1, 0)
    Me.chk整数.Value = IIf(int整数 = 1, 1, 0)
    Me.chk检查库存.Value = IIf(int检查可用数量 = 1, 1, 0)
    
    On Error GoTo errHandle

    '----提取药品剂型----
    '没有药品剂型时仍可以继续，只是不能限定药品的剂型
    gstrSQL = "Select 编码,名称 From 药品剂型"
    Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "提取剂型数据")
    
    Me.lvwSelect.ListItems.Clear
    Me.lvwSelect.ListItems.Add , "R", "所有剂型", , 1
    With rsTemp
        Do While Not .EOF
            Me.lvwSelect.ListItems.Add , "K" & !编码, !名称, , 1
            .MoveNext
        Loop
    End With
    
    '仅提取该部门拥有的材质
'    cbo材质分类.Clear
    gstrSQL = " Select distinct substr(工作性质,1,2) 工作性质" & _
              " From 部门性质说明" & _
              " Where 部门ID = [1] ANd 工作性质 IN ('西药房','西药库','成药房','成药库','中药房','中药库')"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[仅提取该部门拥有的材质]", mlngStockID)
    
    With rsTemp
        Do While Not .EOF
'            cbo材质分类.AddItem IIf(!工作性质 Like "西药*", "西成药", IIf(!工作性质 Like "成药*", "中成药", "中草药"))
'            cbo材质分类.ItemData(cbo材质分类.NewIndex) = IIf(!工作性质 Like "西药*", 5, IIf(!工作性质 Like "成药*", 6, 7))
            
            If !工作性质 Like "西药*" Then
                chk材质西成药.Visible = True
                chk材质西成药.Value = 1
            ElseIf !工作性质 Like "成药*" Then
                chk材质中成药.Visible = True
                chk材质中成药.Value = 1
            Else
                chk材质中草药.Visible = True
                chk材质中草药.Value = 1
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then
'            cbo材质分类.ListIndex = 0
'        Else
            Exit Sub
        End If
    End With
    
    '----提取药品库房----
    'gstrSQL = ReturnSQL(mlngStockID, False)
    'Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取可以申领的库房]", mlngStockID)
    Set rsTemp = ReturnSQL(mlngStockID, Me.Caption & "[提取可以申领的库房]", False, 1343)
    
    If rsTemp.EOF Then
        MsgBox "没有任何库房允许申领，请在[基础参数设置]的药品流向中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    With cbo库房
        .Clear
        mstr库房id = ""
        Do While Not rsTemp.EOF
            If InStr(1, mstr库房id, "|" & rsTemp!Id & "|") = 0 Then
                .AddItem rsTemp!名称
                .ItemData(.NewIndex) = rsTemp!Id
                mstr库房id = mstr库房id & "|" & rsTemp!Id & "|"
                If rsTemp!药库性质 = 1 And intStock = 0 Then
                    intStock = .NewIndex
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        .ListIndex = intStock
    End With
    
    dateCurDate = zlDataBase.Currentdate()
    '----设置缺省的时间范围（一个月）----
    Me.dtp开始时间.Value = Format(DateAdd("m", -1, dateCurDate), "yyyy-MM-dd") & " 00:00:00"
    Me.dtp结束时间.Value = Format(dateCurDate, "yyyy-MM-dd HH:mm:ss")
        
    If mintAutoType = 7 Then
        '特殊的申领方式，取上次审核的申领单的日期作为开始时间
        gstrSQL = " Select a.频次 As 结束时间 " & _
            " From 药品收发记录 A, " & _
            " (Select Nvl(Max(审核日期), Sysdate) As 审核日期 " & _
            " From 药品收发记录 " & _
            " Where 单据 = 6 And 单量 = 7 And 入出系数 = 1 And 库房id + 0 = [1] And 审核日期 Between Sysdate - 60 And Sysdate) B " & _
            " Where a.单据 = 6 And a.单量 = 7 And a.入出系数 = 1 And a.库房id + 0 = [1] And a.审核日期 = b.审核日期 And Rownum = 1 "

'        gstrSQL = "Select Max(审核日期) As 审核日期 From 药品收发记录 " & _
'            " Where 单据 = 6 And 单量 = 7 And 入出系数 = 1 And 库房id = [1] And 审核日期 Between Sysdate - 60 And Sysdate "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "mnuEditAddAutoBySale_Click", mlngStockID)
        
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp!结束时间) Then
                Me.dtp开始时间.Value = Format(DateAdd("s", 1, rsTemp!结束时间), "yyyy-mm-dd hh:mm:ss")
                Me.dtp开始时间.Enabled = False
            End If
        End If
    End If
    mblnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品申领管理", "申领数量作为参考数量", Me.chk申领数量.Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品申领管理", "确保申领数量为整数", Me.chk整数.Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品申领管理", "检查可用数量", Me.chk检查库存.Value)
End Sub

Private Sub lvwSelect_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim intItem As Integer, intItems As Integer, intSelectItems As Integer
    Dim BlnSelect As Boolean
    
    intItems = lvwSelect.ListItems.count
    If Item.Key = "R" Then
        '全清或全选
        If lvwSelect.Tag = "" Then
            BlnSelect = Item.Checked
            For intItem = 2 To intItems
                lvwSelect.ListItems(intItem).Checked = BlnSelect
            Next
        End If
    Else
        lvwSelect.Tag = "1"     '表示不需要激活事件
        intSelectItems = 0
        For intItem = 2 To intItems
            If lvwSelect.ListItems(intItem).Checked Then intSelectItems = intSelectItems + 1
        Next
        If intSelectItems = intItems - 1 Then
            '全选
            lvwSelect.ListItems(1).Checked = True
        Else
            '没选任何一个
            lvwSelect.ListItems(1).Checked = False
        End If
        lvwSelect.Tag = ""
    End If
End Sub

Private Sub Get剂型串()
    Dim intItem As Integer, intItems As Integer
    mstr剂型 = ""
    intItems = lvwSelect.ListItems.count
    
    If lvwSelect.ListItems(1).Checked Then
        '全选
        mstr剂型 = "1"
    Else
        For intItem = 2 To intItems
            If lvwSelect.ListItems(intItem).Checked Then
                mstr剂型 = mstr剂型 & ",'" & Mid(lvwSelect.ListItems(intItem).Key, 2) & "'"
            End If
        Next
        If mstr剂型 <> "" Then
            mstr剂型 = "(" & Mid(mstr剂型, 2) & ")"
        Else
            '已选择的剂型为空
            mstr剂型 = "-1"
        End If
    End If
End Sub

Private Sub ResizeDrug()
    Dim blnEnable As Boolean
    '判断是否允许用户输入其它条件
    blnEnable = (optMode(根据申领单未发数) Or optMode(根据消耗量) Or optMode(根据日销售量) Or optMode(根据销售总量))
    lbl其它条件设置.Visible = blnEnable
    lbl开始时间.Visible = blnEnable
    lbl结束时间.Visible = blnEnable
    dtp开始时间.Visible = blnEnable
    dtp结束时间.Visible = blnEnable
    
    If blnEnable Then
        lvwSelect.Width = lbl开始时间.Left - 200 - lvwSelect.Left
    Else
        lvwSelect.Width = fraStep(1).Width - 200 - lvwSelect.Left
    End If
    
    blnEnable = optMode(根据日销售量)
    lbl库存限额条件.Visible = blnEnable
    lbl上限天数.Visible = blnEnable
    lbl下限天数.Visible = blnEnable
    txt上限天数.Visible = blnEnable
    txt下限天数.Visible = blnEnable
End Sub

Public Function ShowNavigation(ByVal frmParent As Object, ByVal lngStockid As Long, ByRef intAutoType As Integer, ByRef strEndTime As String, ByRef bln申领状态 As Boolean) As Boolean
    On Error Resume Next
    mlngStockID = lngStockid
    mintAutoType = intAutoType  '1-普通自动申领;7-按销量自动申领
    mblnOK = False
    Set mfrmMain = frmParent
    Me.Show 1, frmParent
    ShowNavigation = mblnOK
    intAutoType = mintAutoType
    If Me.chk申领数量.Value = 0 Then
        bln申领状态 = True
    End If
    If mintAutoType = 7 Then
        strEndTime = Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")
    End If
End Function

Private Function CheckData() As Boolean
    Dim lng天数 As Long
    Dim lngTargetID As Long             '目标库房的ID
    Dim str剂型 As String
    Dim rsCheck As New ADODB.Recordset
    Dim str用途ID As String
    Dim n As Integer
    
    '检查是否存在符合条件的记录（始终只按总数量进行比较，具体分的时候再按是否明确批次来分配各批次）
    On Error GoTo ErrHand
    CheckData = False

    For n = 1 To tvw用途.Nodes.count
        If tvw用途.Nodes(n).Key <> "Root" And _
            tvw用途.Nodes(n).Key <> "_中成药" And _
            tvw用途.Nodes(n).Key <> "_中草药" And _
            tvw用途.Nodes(n).Key <> "_西成药" And _
            tvw用途.Nodes(n).Checked Then
            str用途ID = str用途ID & "," & Mid(tvw用途.Nodes(n).Key, 2)
        End If
    Next

    If str用途ID <> "" Then
        str用途ID = Mid(str用途ID, 2)
    End If
    
    gstrSQL = ""
    str剂型 = IIf(mstr剂型 = "1", "", IIf(mstr剂型 = "-1", " And 1=2", " And (C.剂型 IN " & mstr剂型 & " Or C.剂型 Is NULL)"))
    lngTargetID = cbo库房.ItemData(cbo库房.ListIndex)
    
    If optMode(根据消耗量) Then
        '如果明确批次，则药品库存中没有记录的药品数据，不提取出来
        gstrSQL = "" & _
                 " Select Distinct Nvl(A.申领数量,0) 申领数量,Nvl(B.可用数量,0) 可用数量,Nvl(B.实际数量,0) 实际数量,Nvl(B.实际金额,0) 实际金额,Nvl(B.实际差价,0) 实际差价, " & _
                 "        D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率, " & _
                 "        D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备 " & _
                 " From (Select 库房id, 药品id, Sum(Nvl(实际数量, 0) * Nvl(付数, 1)) 申领数量 " & _
                 " From 药品收发记录 Where 库房id = [2] And 单据 In (7,8,9,10,11) And 入出系数 = -1 And " & _
                 " 审核日期 Between [3] And [4] Group By 库房id, 药品id Having Sum(Nvl(实际数量, 0) * Nvl(付数, 1)) > 0) A,药品信息 C,药品规格 D,收费项目目录 F,收费项目别名 E,收费价目 P, 诊疗项目目录 M,诊疗分类目录 L, " & _
                 "      (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,sum(Nvl(实际数量,0)) 实际数量,Sum(Nvl(实际金额,0)) 实际金额 ,Sum(Nvl(实际差价,0)) 实际差价" & _
                 "      From 药品库存 Where 库房ID=[1] And 性质=1 Group By 药品ID) B "
        If chk仅申领库存小于消耗量.Value = 1 Then
            gstrSQL = gstrSQL & ",(Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,sum(Nvl(实际数量,0)) 实际数量,Sum(Nvl(实际金额,0)) 实际金额 ,Sum(Nvl(实际差价,0)) 实际差价" & _
                " From 药品库存 Where 库房ID=[2] And 性质=1 Group By 药品ID) K "
        End If
        gstrSQL = gstrSQL & "" & _
                 " Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.药品ID+0=B.药品ID(+) And A.药品ID+0=D.药品ID And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist))) " & str剂型 & _
                 " And D.药品ID=F.ID And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select 执行科室ID From 收费执行科室 X Where 执行科室ID=[2] And X.收费细目id = D.药品id) " & _
                 " And Exists (Select 执行科室ID From 收费执行科室 Y Where 执行科室ID=[1] And Y.收费细目id = D.药品id) " & _
                 " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & IIf(chk检查库存.Value = 1, " And Nvl(b.可用数量, 0) <> 0", "") & _
                 IIf(chk仅申领库存小于消耗量.Value = 1, " And A.药品ID=K.药品ID(+) And Nvl(A.申领数量,0)>Nvl(K.可用数量,0)", "") & _
                 " Order By F.编码 "
       Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, CDate(Format(dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")), mstr分类, str用途ID)
    ElseIf optMode(根据上限) Then
       gstrSQL = "Select Distinct " & IIf(chk上限.Value = 1, "Nvl(A.上限,0)", "Nvl(A.上限,0)-Sum(Nvl(B.可用数量,0))") & " 申领数量,Sum(Nvl(K.可用数量,0)) 可用数量,Sum(Nvl(K.实际数量,0)) 实际数量,Sum(Nvl(K.实际金额,0)) 实际金额,Sum(Nvl(K.实际差价,0)) 实际差价,  " & _
                "         D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                "         D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备 " & _
                "  From (Select 药品id,上限 From 药品储备限额 Where 库房ID=[2] And Nvl(上限,0)>0" & _
                " ) A, " & _
                "       药品信息 C,药品规格 D,收费项目目录 F,收费项目别名 E,收费价目 P,诊疗项目目录 M,诊疗分类目录 L , " & _
                "       (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K," & _
                "       (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[1] Group By 执行科室ID,收费细目ID) I, " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "       From 药品库存 Where 库房ID=[2] And 性质=1" & _
                "       Group by 药品ID) B,  " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "      From 药品库存 Where 库房ID=[1] And 性质=1" & _
                "      Group by 药品ID) K " & _
                " Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.药品ID+0=D.药品ID And A.药品ID+0=B.药品ID(+) And A.药品ID+0=K.药品ID(+) And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str剂型 & _
                " And D.药品ID=F.ID And D.药品ID=K.收费细目ID And D.药品ID=I.收费细目ID " & _
                " And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
                " Having Nvl(A.上限,0)-Sum(Nvl(B.可用数量,0))>0 " & IIf(chk检查库存.Value = 1, " And Sum(Nvl(k.可用数量, 0))<>0 ", "") & _
                " Group By Nvl(A.上限,0),D.药品ID,F.编码,F.名称,E.名称,F.是否变价,D.药库分批,D.药房分批,P.现价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                " D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位, D.是否常备 "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, mstr分类, str用途ID)
    ElseIf optMode(根据下限) Then
        gstrSQL = "Select Distinct " & IIf(chkLowerLimit(0).Value = 1, "Nvl(A.下限,0)", IIf(chkLowerLimit(1).Value = 1, "Nvl(A.上限,0)", "Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0))")) & " 申领数量,Sum(Nvl(K.可用数量,0)) 可用数量,Sum(Nvl(K.实际数量,0)) 实际数量,Sum(Nvl(K.实际金额,0)) 实际金额,Sum(Nvl(K.实际差价,0)) 实际差价,  " & _
                "         D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                "         D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备  " & _
                "  From (Select 药品id,上限,下限 From 药品储备限额 Where 库房ID=[2] And Nvl(下限,0)>0" & IIf(chkLowerLimit(1).Value = 1, " And Nvl(上限,0)>0", "") & _
                " ) A, " & _
                "       药品信息 C,药品规格 D,收费项目目录 F, 收费项目别名 E,收费价目 P,诊疗项目目录 M,诊疗分类目录 L, " & _
                "      (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K," & _
                "      (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[1] Group By 执行科室ID,收费细目ID) I, " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "       From 药品库存 Where 库房ID=[2] And 性质=1" & _
                "       Group by 药品ID) B,  " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "      From 药品库存 Where 库房ID=[1] And 性质=1" & _
                "      Group by 药品ID) K " & _
                "  Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.药品ID+0=D.药品ID And A.药品ID+0=B.药品ID(+) And A.药品ID+0=K.药品ID(+) And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in(select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str剂型 & _
                "  And D.药品ID=F.ID And D.药品ID=K.收费细目ID And D.药品ID=I.收费细目ID " & _
                "  And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
                "  Having Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0))>0 " & IIf(chk检查库存.Value = 1, " And Sum(Nvl(k.可用数量, 0))<>0 ", "") & _
                "  Group By Nvl(A.上限,0),Nvl(A.下限,0),D.药品ID,F.编码,F.名称,E.名称,F.是否变价,D.药库分批,D.药房分批,P.现价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                "        D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位, D.是否常备 "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, mstr分类, str用途ID)
    ElseIf optMode(根据上下限) Then
        '提取库存低于下限的药品，按补足上限来填写申领数量
        gstrSQL = "Select Distinct Nvl(A.上限,0)-Sum(Nvl(B.可用数量,0)) As 申领数量, Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0)) As 补足下限数量,Sum(Nvl(K.可用数量,0)) 可用数量,Sum(Nvl(K.实际数量,0)) 实际数量,Sum(Nvl(K.实际金额,0)) 实际金额,Sum(Nvl(K.实际差价,0)) 实际差价,  " & _
                "         D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                "         D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备  " & _
                "  From (Select 药品id,上限,下限 From 药品储备限额 Where 库房ID=[2] And Nvl(上限,0)>0 And Nvl(下限,0)>0" & _
                " ) A, " & _
                "       药品信息 C,药品规格 D,收费项目目录 F, 收费项目别名 E,收费价目 P,诊疗项目目录 M,诊疗分类目录 L," & _
                "      (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K," & _
                "      (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[1] Group By 执行科室ID,收费细目ID) I, " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "       From 药品库存 Where 库房ID=[2] And 性质=1" & _
                "       Group by 药品ID) B,  " & _
                "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                "      From 药品库存 Where 库房ID=[1] And 性质=1" & _
                "      Group by 药品ID) K " & _
                "  Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.药品ID+0=D.药品ID And A.药品ID+0=B.药品ID(+) And A.药品ID+0=K.药品ID(+) And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in(select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str剂型 & _
                "  And D.药品ID=F.ID And D.药品ID=K.收费细目ID And D.药品ID=I.收费细目ID " & _
                "  And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
                "  Having Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0))>0 " & IIf(chk检查库存.Value = 1, " And Sum(Nvl(k.可用数量, 0))<>0 ", "") & _
                "  Group By Nvl(A.上限, 0),Nvl(A.下限,0),D.药品ID,F.编码,F.名称,E.名称,F.是否变价,D.药库分批,D.药房分批,P.现价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                "        D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位,D.是否常备 "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, mstr分类, str用途ID)
    ElseIf optMode(根据申领单未发数) Then
        '根据申领单未发数（不加条件And Nvl(A.发药方式,0)=1 ，是因为审核时，是删除申领单，产生移库单再审核的，标志已经没有了）
        gstrSQL = "select Distinct A.申领数量,Nvl(B.可用数量,0) 可用数量,Nvl(B.实际数量,0) 实际数量,Nvl(B.实际金额,0) 实际金额,Nvl(B.实际差价,0) 实际差价, " & _
                 "        D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率, " & _
                 "        D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备 " & _
                 " from (Select 库房id, 药品id, Sum(Nvl(填写数量,0) - Nvl(实际数量,0)) 申领数量 " & _
                 " From 药品收发记录 Where 库房id = [1] And 对方部门id = [2] And 单据 = 6 And " & _
                 " 审核日期 Between [3] And [4] Group By 库房id, 药品id Having Sum(填写数量 - 实际数量) > 0) A,药品信息 C,药品规格 D,收费项目目录 F,收费项目别名 E,收费价目 P,诊疗项目目录 M,诊疗分类目录 L , " & _
                 "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                 "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                 "      From 药品库存 Where 库房ID=[1] And 性质=1" & _
                 "      Group by 药品ID) B " & _
                 " Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.药品ID+0=B.药品ID(+) And A.药品ID+0=D.药品ID And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in(select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))" & str剂型 & _
                 " And D.药品ID=F.ID And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select 执行科室ID From 收费执行科室 X Where 执行科室ID=[2] And X.收费细目id = D.药品id) " & _
                 " And Exists (Select 执行科室ID From 收费执行科室 Y Where 执行科室ID=[1] And Y.收费细目id = D.药品id) " & _
                 " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & IIf(chk检查库存.Value = 1, " And Nvl(B.可用数量,0)<>0 ", "") & _
                 " Order By F.编码 "
         Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, CDate(Format(dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")), mstr分类, str用途ID)
    ElseIf optMode(根据日销售量) Then
        '药房库存上下限设置，先计算近期某药房各品种日平均销售量。
        '1、近期某药房各品种日平均销售量=近期(时间长短的自定义，如果比较麻烦，可以设置成上月全月时间)某药房各品种药品销售量总和/上月天数
        '2、药房库存上限 = 近期某药房各品种药品日平均销售量 * 需要设定库存上限天数(时间可自定义)
        '3、药房库存下限 = 近期某药房各品种药品日平均销售量 * 需要设定库存下限天数(时间可自定义)
        '4、药房自动申领计划为:
        '   (1)、某药房单品种现库存量>= 药房库存下限，不产生申领计划
        '   (2)、某药房单品种现库存量< 药房库存下限，产生申领计划
        '   (3)、药房申领计划=本药房库存上限-现有库存量
        lng天数 = CDate(Format(dtp结束时间.Value, "yyyy-MM-dd")) - CDate(Format(dtp开始时间.Value, "yyyy-MM-dd")) + 1
        If lng天数 <= 0 Then lng天数 = 1
            gstrSQL = "Select Distinct Nvl(A.库存上限,0)-Nvl(B.可用数量,0) 申领数量,Nvl(K.可用数量,0) 可用数量,Nvl(K.实际数量,0) 实际数量,Nvl(K.实际金额,0) 实际金额,Nvl(K.实际差价,0) 实际差价,  " & _
                 "         D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率,  " & _
                 "         D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备  " & _
                 "  From (SELECT A.药品ID,A.销售量,A.销售量*" & Val(txt上限天数.Text) & " AS 库存上限,A.销售量*" & Val(txt下限天数.Text) & " AS 库存下限" & _
                 "        FROM " & _
                 "           (SELECT 药品ID,SUM(NVL(实际数量,0)*NVL(付数,1))/" & lng天数 & " AS 销售量" & _
                 "           FROM 药品收发记录 WHERE 库房ID+0=[2] AND 单据 IN (8,9,10)" & _
                 "           AND 审核日期 BETWEEN [3] AND [4] " & _
                 "           GROUP BY 药品ID) A ) A," & _
                 "       药品信息 C,药品规格 D,收费项目目录 F,收费项目别名 E,收费价目 P,诊疗项目目录 M,诊疗分类目录 L,"
            gstrSQL = gstrSQL & _
                 "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                 "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                 "       From 药品库存 Where 库房ID=[2] And 性质=1" & _
                 "       Group by 药品ID) B,  " & _
                 "       (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,Sum(Nvl(实际数量,0)) 实际数量, " & _
                 "       Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价   " & _
                  "      From 药品库存 Where 库房ID=[1] And 性质=1" & _
                  "      Group by 药品ID) K " & _
                 "  Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.药品ID+0=D.药品ID And A.药品ID+0=B.药品ID(+) And A.药品ID+0=K.药品ID(+) And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in(select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist))) " & str剂型 & _
                 "  And D.药品ID=F.ID " & _
                 "  AND Nvl(B.可用数量,0)<A.库存下限 " & _
                 "  And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select 执行科室ID From 收费执行科室 X Where 执行科室ID=[2] And X.收费细目id = D.药品id) " & _
                 " And Exists (Select 执行科室ID From 收费执行科室 Y Where 执行科室ID=[1] And Y.收费细目id = D.药品id) " & _
                 " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & IIf(chk检查库存.Value = 1, " And Nvl(K.可用数量,0)<>0 ", "") & _
                 " Order By F.编码 "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, CDate(Format(dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")), mstr分类, str用途ID)
    ElseIf optMode(根据销售总量) Then
        '根据指定时间范围内的销售（发药）量，作为本次的申领数量
        gstrSQL = "" & _
                 " Select Distinct Nvl(A.申领数量,0) 申领数量,Nvl(B.可用数量,0) 可用数量,Nvl(B.实际数量,0) 实际数量,Nvl(B.实际金额,0) 实际金额,Nvl(B.实际差价,0) 实际差价, " & _
                 "        D.药品ID,F.编码,F.名称 As 通用名,E.名称 As 商品名,F.是否变价,D.药库分批,D.药房分批,P.现价 售价,F.规格,F.产地,D.原产地,D.最大效期,D.加成率, " & _
                 "        D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装,F.计算单位 售价单位, D.是否常备 " & _
                 " From  (Select  库房id, 药品id, Sum(Nvl(实际数量, 0) * Nvl(付数, 1)) 申领数量 " & _
                 " From 药品收发记录 Where 库房id = [2] And 单据 In (8, 9, 10) And 入出系数 = -1 And " & _
                 " 审核日期 Between [3] And [4] Group By 库房id, 药品id Having Sum(Nvl(实际数量, 0) * Nvl(付数, 1)) > 0) A,药品信息 C,药品规格 D,收费项目目录 F,收费项目别名 E,收费价目 P, 诊疗项目目录 M,诊疗分类目录 L," & _
                 "      (Select 药品ID,Sum(Nvl(可用数量,0)) 可用数量,sum(Nvl(实际数量,0)) 实际数量,Sum(Nvl(实际金额,0)) 实际金额 ,Sum(Nvl(实际差价,0)) 实际差价" & _
                 "      From 药品库存 Where 库房ID=[1] And 性质=1 Group By 药品ID) B " & _
                 " Where D.药名ID=M.ID And M.分类ID=L.ID And B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And 码类(+) = 1 AND L.类型 In (1,2,3) " & IIf(str用途ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.药品ID+0=B.药品ID(+) And A.药品ID+0=D.药品ID And D.药名ID=C.药名ID And D.药品ID=P.收费细目ID And F.类别 in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))" & str剂型 & _
                 " And D.药品ID=F.ID And SysDate Between P.执行日期 And Nvl(P.终止日期,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select 执行科室ID From 收费执行科室 X Where 执行科室ID=[2] And X.收费细目id = D.药品id) " & _
                 " And Exists (Select 执行科室ID From 收费执行科室 Y Where 执行科室ID=[1] And Y.收费细目id = D.药品id) " & _
                 " And (F.撤档时间 Is Null Or To_char(F.撤档时间,'yyyy-MM-dd')='3000-01-01') " & IIf(chk检查库存.Value = 1, " And Nvl(B.可用数量,0)<>0 ", "") & _
                 " Order By F.编码 "
       Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在符合条件的记录]", lngTargetID, mlngStockID, CDate(Format(dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")), mstr分类, str用途ID)
    End If
    
    If rsCheck.RecordCount = 0 Then
        MsgBox "没找到符合条件的记录！", vbInformation, gstrSysName
        mintStep = mintStep - 1
        Exit Function
    End If
    
    On Error GoTo 0
    Call WriteResult(rsCheck)
    
    Dim intCount As Integer
    With frmRequestDrugCard
        For intCount = 0 To .cboStock.ListCount - 1
            If .cboStock.ItemData(intCount) = lngTargetID Then
                .cboStock.ListIndex = intCount: Exit For
            End If
        Next
    End With
    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub WriteResult(ByVal rsCheck As ADODB.Recordset)
    Dim strUnit As String
    Dim lngTargetID As Long
    Dim blnAdd As Boolean
    Dim bln提示 As Boolean, bln库房 As Boolean
    Dim bln允许 As Boolean, bln特药 As Boolean       'bln允许-根据系统参数“库存检查”和用户操作来决定是否产生无库存的药品；bln特药-当前药品是否是时价或批次药品
    Dim dbl申领数量 As Double, dbl填写数量 As Double, dbl比例系数 As Double
    Dim rsStock As New ADODB.Recordset  '药品库存
    Dim rsTemp  As New ADODB.Recordset
    Dim blnStock As Boolean             '是否常备药品
    Dim blnShowMsg As Boolean
        
    On Error GoTo errHandle
    lngTargetID = cbo库房.ItemData(cbo库房.ListIndex)
    Call GetPara(lngTargetID)
    bln库房 = CheckStock(lngTargetID)
    strUnit = GetDrugUnit(mlngStockID, "药品申领管理")
    
    Call GetDrugDigit(lngTargetID, "药品申领管理", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    '准备产生数据（全部以零售单位为准，最终在SetColValue函数中转换，传入的系数为当前单位的系数）
    With rsCheck
        Do While Not .EOF
            dbl申领数量 = Calc_Clique(!药品ID, !申领数量)
            '确定常备药品
            blnStock = IIf(IsNull(!是否常备), False, !是否常备)
            blnStock = Not blnStock
            If optDrugType(1).Value Then
                If blnStock = False Then GoTo Continue
            ElseIf optDrugType(2).Value Then
                If blnStock = True Then GoTo Continue
            End If
            
            If mint按批次出库 = 1 Then
                gstrSQL = " Select Nvl(可用数量,0) 可用数量,Nvl(实际数量,0) 实际数量,Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价," & _
                          "     Nvl(批次,0) 批次,效期,上次批号 批号,上次产地 产地,原产地,NVL(上次供应商ID,0) 上次供应商ID,批准文号 " & _
                          " From 药品库存 Where 库房ID=[1] And 药品ID=[2] And 性质=1"
                If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
                    gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
                Else
                    gstrSQL = gstrSQL & " Order by 效期,Nvl(批次,0)"
                End If
                
                Set rsStock = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取该药品的库存]", lngTargetID, CLng(!药品ID))
                
                blnAdd = False
                If rsStock.RecordCount <> 0 Then
                    '有库存的药品、分批或时价药品按此产生
                    Do While Not rsStock.EOF
                        If dbl申领数量 >= rsStock!可用数量 Then
                            dbl填写数量 = IIf(rsStock!可用数量 > 0, rsStock!可用数量, 0)
                            '校正填写数量
                            dbl填写数量 = Calc_Clique(!药品ID, dbl填写数量, True)
                        Else
                            '不需要校正，因为最外层已校正，经过其中多个分支后，其剩余数仍然是符合要求的，所以不需校正
                            dbl填写数量 = dbl申领数量
                        End If
                        
                        '明确批次，所以需要按可用数量填写申领单
                        dbl比例系数 = IIf(strUnit = "住院单位", !住院包装, IIf(strUnit = "门诊单位", !门诊包装, IIf(strUnit = "药库单位", !药库包装, 1)))
                        If dbl填写数量 <> 0 Then
                            
                            If SetColValue(!药品ID, "[" & !编码 & "]", !通用名, IIf(IsNull(!商品名), "", !商品名), IIf(IsNull(!规格), "", !规格), IIf(IsNull(rsStock!产地), "", rsStock!产地), _
                                IIf(strUnit = "住院单位", !住院单位, IIf(strUnit = "门诊单位", !门诊单位, IIf(strUnit = "药库单位", !药库单位, !售价单位))), _
                                !售价, IIf(IsNull(rsStock!批号), "", rsStock!批号), IIf(IsNull(rsStock!效期), "", rsStock!效期), zlStr.Nvl(!最大效期, 0), !药库分批, IIf(IsNull(!可用数量), 0, !可用数量), _
                                IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !加成率 / 100, _
                                IIf(strUnit = "住院单位", !住院包装, IIf(strUnit = "门诊单位", !门诊包装, IIf(strUnit = "药库单位", !药库包装, 1))), _
                                rsStock!批次, dbl填写数量, !药房分批, !是否变价, zlStr.Nvl(rsStock!上次供应商ID, 0), _
                                IIf(IsNull(rsStock!批准文号), "", rsStock!批准文号), blnStock, IIf(IsNull(rsStock!原产地), "", rsStock!原产地)) Then blnAdd = True
                        End If
                        
                        dbl申领数量 = dbl申领数量 - dbl填写数量
                        If dbl申领数量 = 0 Then Exit Do
                        rsStock.MoveNext
                    Loop
                    If dbl申领数量 > 0 And blnAdd Then
                        '未申领完的数量全部放在最后一行的药品上
                        If Me.chk申领数量.Value = 1 Then
                            frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol申领数量) = zlStr.FormatEx(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol申领数量)) + dbl申领数量 / dbl比例系数, mintNumberDigit, , True)
                        Else
                            frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量) = zlStr.FormatEx(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量)) + dbl申领数量 / dbl比例系数, mintNumberDigit, , True)
                        End If
                        
                        If chk整数.Value = 1 Then
                            If Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量)) <> Int(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量))) Then
                                frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量) = zlStr.FormatEx(Int(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol填写数量))) + 1, mintNumberDigit, , True)
                            End If
                        End If
                    End If
                Else
                    '不分批且无时价属性的药品按此产生
                    '如果参数为不足禁止，根本不执行以下语句
                    If mintCheck <> 2 Then
                        gstrSQL = " Select Nvl(A.药库分批,0) 药库分批,Nvl(A.药房分批,0) 药房分批,Nvl(B.是否变价,0) 时价, a.上次供应商ID " & _
                                  " From 药品规格 A,收费项目目录 B" & _
                                  " Where A.药品ID = B.ID And A.药品ID = [1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该药品对于出库库房是否分批、时价的属性]", CLng(!药品ID))
                        
                        bln特药 = (rsTemp!时价 = 1) Or IIf(bln库房, (rsTemp!药库分批 = 1), (rsTemp!药房分批 = 1))
                        If Not bln特药 Then
                            If Not bln提示 Then
                                If mintCheck = 1 Then
                                    bln允许 = (MsgBox("药品库存不足,是否继续申领？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
                                Else
                                    bln允许 = True
                                End If
                                bln提示 = True
                            End If
                            If bln允许 Then
                                '为无库存药品产生申领记录
                                If dbl申领数量 <> 0 Then
                                    Call SetColValue(!药品ID, "[" & !编码 & "]", !通用名, IIf(IsNull(!商品名), "", !商品名), IIf(IsNull(!规格), "", !规格), "", _
                                        IIf(strUnit = "住院单位", !住院单位, IIf(strUnit = "门诊单位", !门诊单位, IIf(strUnit = "药库单位", !药库单位, !售价单位))), _
                                        !售价, "", "", zlStr.Nvl(!最大效期, 0), !药库分批, IIf(IsNull(!可用数量), 0, !可用数量), _
                                        IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !加成率 / 100, _
                                        IIf(strUnit = "住院单位", !住院包装, IIf(strUnit = "门诊单位", !门诊包装, IIf(strUnit = "药库单位", !药库包装, 1))), _
                                        0, dbl申领数量, !药房分批, !是否变价, IIf(rsTemp Is Nothing, 0, zlStr.Nvl(rsTemp!上次供应商ID, 0)), "", blnStock, "")
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                '依据传入记录集产生数据
                If dbl申领数量 <> 0 Then
'                    If mintCheck = 2 And dbl申领数量 > Val(IIf(IsNull(!可用数量), 0, !可用数量)) Then
'                        '库存不足禁止
'                        If blnShowMsg = False Then
'                            MsgBox "出库库房已设置了出库检查，库存不足时将不产生申领数据。", vbInformation, gstrSysName
'                            blnShowMsg = True
'                        End If
'                    Else
                        '上次供应商ID
                        gstrSQL = "select 上次供应商ID, 上次产地 from 药品规格 where 药品ID = [1] "
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-上次供应商", CLng(!药品ID))
                    
                        '产生数据
                        Call SetColValue(!药品ID, "[" & !编码 & "]", !通用名, IIf(IsNull(!商品名), "", !商品名), IIf(IsNull(!规格), "", !规格), _
                            IIf(rsTemp Is Nothing, zlStr.Nvl(!产地), zlStr.Nvl(rsTemp!上次产地)), _
                            IIf(strUnit = "住院单位", !住院单位, IIf(strUnit = "门诊单位", !门诊单位, IIf(strUnit = "药库单位", !药库单位, !售价单位))), _
                            !售价, "", "", zlStr.Nvl(!最大效期, 0), !药库分批, IIf(IsNull(!可用数量), 0, !可用数量), _
                            IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !加成率 / 100, _
                            IIf(strUnit = "住院单位", !住院包装, IIf(strUnit = "门诊单位", !门诊包装, IIf(strUnit = "药库单位", !药库包装, 1))), _
                            0, dbl申领数量, !药房分批, !是否变价, IIf(rsTemp Is Nothing, 0, zlStr.Nvl(rsTemp!上次供应商ID, 0)), "", blnStock, zlStr.Nvl(!原产地))
'                    End If
                End If
            End If
Continue:
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal lng药品id As Long, ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str规格 As String, _
    ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal int分批核算 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal dbl加成率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal Dbl数量 As Double, ByVal int药房分批 As Integer, ByVal int是否变价 As Integer, _
    ByVal lng上次供应商ID As Long, ByVal str批准文号 As String, ByVal bln是否常备 As Boolean, ByVal str原产地 As String) As Boolean
    
    Dim intDrugNameShow As Integer
    Dim str药名 As String
    Dim intCount As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim str药品来源 As String, str基本药物 As String
    
    Dim mconIntCol药名  As Integer    ' = 2
    Dim mconIntCol商品名 As Integer  '= 3
    Dim mconIntCol来源 As Integer       '=4
    Dim mconIntCol序号 As Integer   '=5
    Dim mconIntCol规格   As Integer  '= 6
    Dim mconIntCol分批核算  As Integer ' = 7
    Dim mconIntCol最大效期  As Integer   ' = 8
    Dim mconIntCol可用数量  As Integer   ' = 9
    Dim mconIntcol加成率 As Integer     '= 10
    Dim mconIntCol实际金额 As Integer    '= 11
    Dim mconIntCol实际差价 As Integer   ' = 12
    Dim mconIntCol比例系数 As Integer    '= 13

    Dim mconIntCol批次 As Integer    '= 14
    Dim mconIntCol产地 As Integer    '= 15
    Dim mconIntCol原产地 As Integer    '= 16
    Dim mconIntCol单位 As Integer  ' = 17
    Dim mconIntCol批号 As Integer    '= 18
    Dim mconIntCol效期 As Integer     '= 19
    Dim mconIntCol批准文号 As Integer '= 20
    Dim mconintcol当前库存 As Integer
    Dim mconintcol对方库存 As Integer
    Dim mconIntCol填写数量 As Integer  ' = 21
    Dim mconIntCol申领数量 As Integer
    Dim mconIntCol实际数量 As Integer  ' = 22
    Dim mconIntCol采购价 As Integer     '= 23
    Dim mconIntCol采购金额 As Integer   '= 24
    Dim mconIntCol售价 As Integer   '= 25
    Dim mconIntCol售价金额 As Integer   ' = 26
    Dim mconintCol差价 As Integer   '= 27
    Dim mconIntCol上次供应商ID As Long '=28
    Dim mconIntCol药品编码和名称 As Integer
    Dim mconIntCol药品编码 As Integer
    Dim mconIntCol药品名称 As Integer
    Dim mconIntCol基本药物 As Integer
    Dim intCol常备药品 As Integer

    Dim num实际数量 As Double
    Dim rsTemp As New ADODB.Recordset
    mconIntCol药名 = 2
    mconIntCol商品名 = 3
    mconIntCol来源 = 4
    mconIntCol基本药物 = 5
    mconIntCol序号 = 6
    mconIntCol规格 = 7
    mconIntCol分批核算 = 8
    mconIntCol最大效期 = 9
    mconIntCol可用数量 = 10
    mconIntcol加成率 = 11
    mconIntCol实际金额 = 12
    mconIntCol实际差价 = 13
    mconIntCol比例系数 = 14

    mconIntCol批次 = 15
    mconIntCol产地 = 16
    mconIntCol原产地 = 17
    mconIntCol单位 = 18
    mconIntCol批号 = 19
    mconIntCol效期 = 20
    mconIntCol批准文号 = 21
    mconintcol当前库存 = 22
    mconintcol对方库存 = 23
    mconIntCol申领数量 = 24: mIntCol申领数量 = mconIntCol申领数量
    mconIntCol填写数量 = 25:  mIntCol填写数量 = mconIntCol填写数量
    mconIntCol实际数量 = 26
    mconIntCol采购价 = 27
    mconIntCol采购金额 = 28
    mconIntCol售价 = 29
    mconIntCol售价金额 = 30
    mconintCol差价 = 31
    mconIntCol上次供应商ID = 32
    
    mconIntCol药品编码和名称 = 34
    mconIntCol药品编码 = 35
    mconIntCol药品名称 = 36
    intCol常备药品 = 37
    
    SetColValue = False
    On Error GoTo errHandle
    intDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "药品名称显示方式", 0)))
    
    '如果申领数量为零则退出
    If IIf(Dbl数量 >= num可用数量, num可用数量, Dbl数量) = 0 And mint按批次出库 = 1 And (int是否变价 = 1 Or lng批次 <> 0) Then Exit Function

    gstrSQL = "Select 药品来源,基本药物 From 药品规格 Where 药品ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取药品来源]", lng药品id)
    
    str药品来源 = zlStr.Nvl(rsTemp!药品来源)
    str基本药物 = zlStr.Nvl(rsTemp!基本药物)
    
    '如果明确批次时，时价药品根据批次提取售价;
    If mint按批次出库 = 1 And int是否变价 = 1 Then
        num售价 = Get零售价(int是否变价 = 1, lng药品id, Val(cbo库房.ItemData(cbo库房.ListIndex)), lng批次)
    End If
    
    With frmRequestDrugCard.mshBill
        intRow = .rows - 1
        .TextMatrix(intRow, 0) = lng药品id
        .TextMatrix(intRow, 1) = intRow
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = str药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If intDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf intDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = str商品名
        
        .TextMatrix(intRow, mconIntCol来源) = str药品来源
        .TextMatrix(intRow, mconIntCol基本药物) = str基本药物
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol原产地) = str原产地
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        
        If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
            '换算为有效期
            .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
        End If
        
        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(num售价 * num比例系数, mintPriceDigit, , True)
        
        If mint按批次出库 <> 1 And int是否变价 = 1 Then
            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Get时价零售价(lng药品id, cbo库房.ItemData(cbo库房.ListIndex), lng批次, num比例系数), mintPriceDigit, , True)
        End If
        
        .TextMatrix(intRow, mconIntCol分批核算) = int分批核算
        .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(num可用数量 / num比例系数, mintNumberDigit, , True)
        .TextMatrix(intRow, mconIntCol最大效期) = int最大效期 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol实际差价) = zlStr.FormatEx(num实际差价, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol实际金额) = zlStr.FormatEx(num实际金额, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntcol加成率) = dbl加成率
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        '如果是时价药品或分批药品,不能超过当前库存数量
        If Me.chk申领数量.Value = 1 Then
            frmRequestDrugCard.mshBill.ColWidth(mconIntCol申领数量) = 1100
            frmRequestDrugCard.cmd全部复制.Visible = True
            frmRequestDrugCard.cmd全清.Visible = True
            
            If (int是否变价 = 1 Or lng批次 <> 0) And mint按批次出库 = 1 Then
                .TextMatrix(intRow, mconIntCol申领数量) = zlStr.FormatEx(IIf(Dbl数量 >= num可用数量, num可用数量, Dbl数量) / num比例系数, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol填写数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol申领数量) = zlStr.FormatEx(Dbl数量 / num比例系数, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol填写数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
            End If
        Else
            frmRequestDrugCard.cmd全部复制.Visible = False
            frmRequestDrugCard.cmd全清.Visible = False
            
            If (int是否变价 = 1 Or lng批次 <> 0) And mint按批次出库 = 1 Then
                .TextMatrix(intRow, mconIntCol填写数量) = zlStr.FormatEx(IIf(Dbl数量 >= num可用数量, num可用数量, Dbl数量) / num比例系数, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(IIf(Dbl数量 >= num可用数量, num可用数量, Dbl数量) / num比例系数, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol填写数量) = zlStr.FormatEx(Dbl数量 / num比例系数, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(Dbl数量 / num比例系数, mintNumberDigit, , True)
            End If
        End If
        
        If chk整数.Value = 1 Then
            If Val(.TextMatrix(intRow, mconIntCol填写数量)) <> Int(Val(.TextMatrix(intRow, mconIntCol填写数量))) Then
                .TextMatrix(intRow, mconIntCol填写数量) = zlStr.FormatEx(Int(Val(.TextMatrix(intRow, mconIntCol填写数量))) + 1, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol填写数量), mintNumberDigit, , True)
            End If
        End If
        
        If .TextMatrix(intRow, mconIntCol售价) <> "" Then
            .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) * .TextMatrix(intRow, mconIntCol实际数量), mintMoneyDigit, , True)
        End If
        
        .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(Get成本价(lng药品id, Val(cbo库房.ItemData(cbo库房.ListIndex)), lng批次) * num比例系数, mintCostDigit, , True)
        .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol采购价)) * Val(.TextMatrix(intRow, mconIntCol填写数量)), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价金额)) - Val(.TextMatrix(intRow, mconIntCol采购金额)), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol上次供应商ID) = lng上次供应商ID
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, intCol常备药品) = bln是否常备
                             
        .rows = .rows + 1
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckStock(ByVal lng库房ID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '检查指定库房是药库、药房还是制剂室(传入的库房肯定是药库、药房或制剂室中的一个)
    On Error GoTo errHandle
    gstrSQL = " Select 部门ID From 部门性质说明 " & _
              " Where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断是不是药房或制剂室]", lng库房ID)
              
    If rsCheck.EOF Then
        CheckStock = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPara(ByVal lng库房ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '获取出库检查的参数设置值（0-不检查;1-检查，不足提醒;2-不足禁止）
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(检查方式,0) Value From 药品出库检查 Where 库房ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取检查库存的参数]", lng库房ID)
    
    If Not rsTemp.EOF Then
        mintCheck = rsTemp!Value
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optMode_Click(index As Integer)
    mintAutoType = index + 1
    
    chk仅申领库存小于消耗量.Enabled = False
    chk上限.Enabled = False
    chkLowerLimit(0).Enabled = False
    chkLowerLimit(1).Enabled = False
    
    Select Case index
    Case 0
        chk仅申领库存小于消耗量.Enabled = True
    Case 1
        chk上限.Enabled = True
    Case 2
        chkLowerLimit(0).Enabled = True
        chkLowerLimit(1).Enabled = True
    End Select
    
End Sub

Private Sub tvw用途_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If tvw用途.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw用途.Nodes(intIdx).Next.index
            Loop
            If intIdx = Node.LastSibling.index Then
                If tvw用途.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck
        End If
    End If
End Sub


