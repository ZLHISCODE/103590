VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫生材料规格编辑"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmStuffSpec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "保存后新增品种(&A)"
      Height          =   350
      Left            =   2280
      TabIndex        =   59
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "保存后新增规格(&B)"
      Height          =   350
      Left            =   4275
      TabIndex        =   58
      Top             =   7680
      Width           =   1695
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3720
      Left            =   1200
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   8760
      Visible         =   0   'False
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   6562
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   109
      Top             =   7560
      Width           =   8880
   End
   Begin VB.PictureBox picFound 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2835
      ScaleHeight     =   210
      ScaleWidth      =   5505
      TabIndex        =   104
      Top             =   885
      Width           =   5505
      Begin VB.Label lblFound 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "注：该规格建立于2002年12月20日，于2003年8月10日停用。"
         Height          =   225
         Left            =   105
         TabIndex        =   65
         Top             =   0
         Width           =   4860
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf产地 
      Height          =   1845
      Left            =   360
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   8640
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   3254
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   107
      Top             =   570
      Width           =   8775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存退出(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   56
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7785
      TabIndex        =   57
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      Picture         =   "frmStuffSpec.frx":030A
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1100
   End
   Begin TabDlg.SSTab stbSpec 
      Height          =   6705
      Left            =   120
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   720
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   11827
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "卫材信息(&1)"
      TabPicture(0)   =   "frmStuffSpec.frx":0454
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "价格信息(&2)"
      TabPicture(1)   =   "frmStuffSpec.frx":0470
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk分零使用"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmd病案"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra分批核算"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chk屏蔽费别"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cbo服务对象"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cbo费用类型"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cbo收入项目"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fra(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt病案费目"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl病案费目"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl(20)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl(18)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl(19)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CheckBox chk分零使用 
         Caption         =   "分零使用(&G)"
         Height          =   285
         Left            =   -68280
         TabIndex        =   125
         Top             =   2400
         Width           =   1290
      End
      Begin VB.CommandButton cmd病案 
         Caption         =   "…"
         Height          =   240
         Left            =   -67200
         TabIndex        =   122
         TabStop         =   0   'False
         Tag             =   "分类"
         ToolTipText     =   "按*打开选择器"
         Top             =   1950
         Width           =   255
      End
      Begin VB.Frame fra分批核算 
         Caption         =   "分批管理"
         Height          =   1230
         Left            =   -70200
         TabIndex        =   98
         Top             =   2880
         Width           =   3780
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   17
            Left            =   2820
            MaxLength       =   5
            TabIndex        =   102
            Tag             =   "保存期"
            Top             =   375
            Width           =   630
         End
         Begin VB.CheckBox chk库房 
            Caption         =   "卫材库房分批(&W)"
            Height          =   420
            Left            =   105
            TabIndex        =   99
            Top             =   315
            Width           =   1665
         End
         Begin VB.CheckBox chk在用 
            Caption         =   "发料部门分批(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   105
            TabIndex        =   100
            Top             =   750
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "效期(&7)       月"
            Height          =   180
            Left            =   2190
            TabIndex        =   101
            Top             =   435
            Width           =   1440
         End
      End
      Begin VB.Frame fra 
         Caption         =   "分类属性"
         Height          =   2370
         Index           =   1
         Left            =   5370
         TabIndex        =   35
         Top             =   456
         Width           =   3195
         Begin VB.ComboBox cbo存储条件 
            Height          =   300
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Tag             =   "材料来源"
            Top             =   1965
            Width           =   2970
         End
         Begin VB.ComboBox cbo材质分类 
            Height          =   300
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "材料来源"
            Top             =   1290
            Width           =   2970
         End
         Begin VB.ComboBox cbo材料来源 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Tag             =   "材料来源"
            Top             =   690
            Width           =   1950
         End
         Begin VB.ComboBox cbo货源 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Tag             =   "货源情况"
            Top             =   330
            Width           =   1950
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "存储条件(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   25
            Left            =   150
            TabIndex        =   42
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "材质分类(&J)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   150
            TabIndex        =   40
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "来源分类(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   150
            TabIndex        =   38
            Top             =   750
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "货源情况(&Q)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   150
            TabIndex        =   36
            Top             =   390
            Width           =   990
         End
      End
      Begin VB.Frame Fra2 
         Caption         =   "附加属性"
         Height          =   3640
         Left            =   5355
         TabIndex        =   44
         Top             =   2925
         Width           =   3195
         Begin VB.CheckBox chk植入耗材 
            Caption         =   "植入性耗材"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   1260
         End
         Begin VB.CheckBox chkInstrument 
            Caption         =   "器械包卫材单件"
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   285
            Width           =   1575
         End
         Begin VB.CheckBox chkCode 
            Caption         =   "条码管理(&7)"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1507
            Width           =   1365
         End
         Begin VB.CheckBox chkCostly 
            Caption         =   "高值材料(&6)"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1214
            Width           =   1485
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   3030
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txt备选码 
            Height          =   300
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   54
            Top             =   2565
            Width           =   2085
         End
         Begin VB.CheckBox chk原料 
            Caption         =   "原料(&3)"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   1500
         End
         Begin VB.CheckBox chk无菌性材料 
            Caption         =   "无菌卫材(&4)"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   603
            Width           =   1500
         End
         Begin VB.CheckBox Chk一次性材料 
            Caption         =   "一次性卫材(&5)"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   922
            Width           =   1485
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   7
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   53
            Tag             =   "灭菌效期"
            Top             =   2175
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "院区"
            Height          =   180
            Left            =   120
            TabIndex        =   63
            Top             =   3105
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "备选码(&V)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "月"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1920
            TabIndex        =   61
            Top             =   2235
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "灭菌效期(&7)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   52
            Top             =   2235
            Width           =   990
         End
      End
      Begin VB.CheckBox chk屏蔽费别 
         Caption         =   "屏蔽费别(&M)"
         Height          =   285
         Left            =   -70080
         TabIndex        =   97
         Top             =   2400
         Width           =   1290
      End
      Begin VB.ComboBox cbo服务对象 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Tag             =   "应用对象"
         Top             =   1500
         Width           =   2115
      End
      Begin VB.ComboBox cbo费用类型 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Tag             =   "医保类型"
         Top             =   1125
         Width           =   2115
      End
      Begin VB.ComboBox cbo收入项目 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Tag             =   "收入项目"
         Top             =   750
         Width           =   2115
      End
      Begin VB.Frame fra 
         Height          =   6120
         Index           =   0
         Left            =   90
         TabIndex        =   108
         Top             =   450
         Width           =   5190
         Begin VB.CommandButton cmd产地 
            Caption         =   "…"
            Height          =   285
            Left            =   4750
            TabIndex        =   124
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   2978
            Width           =   285
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   22
            Left            =   1125
            MaxLength       =   250
            TabIndex        =   34
            Tag             =   "说明"
            Top             =   5700
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   20
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "五笔简码"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   19
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "拼音简码"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   18
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "商品名"
            Top             =   1050
            Width           =   3945
         End
         Begin VB.CheckBox chk核算材料 
            Caption         =   "核算材料(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   3795
            TabIndex        =   14
            Top             =   2205
            Width           =   1335
         End
         Begin VB.TextBox txt注册证号 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   26
            Top             =   4120
            Width           =   3915
         End
         Begin VB.CheckBox chk跟踪病人 
            Caption         =   "跟踪病人(&S)"
            Height          =   210
            Left            =   3795
            TabIndex        =   33
            Top             =   5367
            Width           =   1290
         End
         Begin MSComCtl2.DTPicker dtp许可证效期 
            Height          =   345
            Left            =   1335
            TabIndex        =   32
            Top             =   5300
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   137363457
            CurrentDate     =   39227
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   30
            Tag             =   "许可证号"
            Top             =   4930
            Width           =   3915
         End
         Begin VB.TextBox txt批准文号 
            Height          =   300
            Left            =   1125
            MaxLength       =   40
            TabIndex        =   22
            Top             =   3350
            Width           =   3915
         End
         Begin VB.TextBox txt注册商标 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   24
            Top             =   3730
            Width           =   3915
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   4575
            MaxLength       =   1
            TabIndex        =   18
            Tag             =   "标识子码"
            Top             =   2595
            Width           =   465
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "标识主码"
            Top             =   2595
            Width           =   1605
         End
         Begin VB.CheckBox chk跟踪 
            Caption         =   "跟踪在用(&I)"
            Height          =   210
            Left            =   2430
            TabIndex        =   13
            Top             =   2205
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "换算系数"
            Text            =   "1"
            Top             =   2160
            Width           =   870
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   1125
            MaxLength       =   60
            TabIndex        =   20
            Tag             =   "生产商"
            Top             =   2970
            Width           =   3615
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   3
            Tag             =   "规格"
            Top             =   660
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "规格编码"
            Top             =   300
            Width           =   2160
         End
         Begin VB.ComboBox cbo单位 
            Height          =   300
            Index           =   0
            ItemData        =   "frmStuffSpec.frx":048C
            Left            =   1125
            List            =   "frmStuffSpec.frx":048E
            TabIndex        =   8
            Tag             =   "散装单位"
            Text            =   "支"
            Top             =   1770
            Width           =   1245
         End
         Begin VB.ComboBox cbo单位 
            Height          =   300
            Index           =   1
            ItemData        =   "frmStuffSpec.frx":0490
            Left            =   3810
            List            =   "frmStuffSpec.frx":0492
            TabIndex        =   10
            Tag             =   "包装单位"
            Text            =   "支"
            Top             =   1770
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker dtp注册证有效期 
            Height          =   345
            Left            =   1335
            TabIndex        =   28
            Top             =   4478
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   137297921
            CurrentDate     =   39227
         End
         Begin VB.Label lbl注册证 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "注册证有效期"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   4560
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "说明(&S)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   30
            Left            =   480
            TabIndex        =   121
            Top             =   5760
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(五笔)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   4530
            TabIndex        =   115
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(拼音)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   2490
            TabIndex        =   114
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "品名简码(&P)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   113
            Top             =   1500
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "商品名称"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   112
            Top             =   1110
            Width           =   720
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "注册证号(&T)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   4180
            Width           =   990
         End
         Begin VB.Label lblIn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "许可证效期(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   5382
            Width           =   1170
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "许可证号"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   4990
            Width           =   720
         End
         Begin VB.Label lbl批准文号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "批准文号(&W)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   3410
            Width           =   990
         End
         Begin VB.Label lbl注册商标 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "注册商标(&E)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   3790
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "标识子码(&D)"
            Height          =   180
            Index           =   16
            Left            =   3570
            TabIndex        =   17
            Top             =   2655
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "标识主码(&Z)"
            Height          =   180
            Index           =   17
            Left            =   120
            TabIndex        =   15
            Top             =   2685
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "换算系数(&X)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   2220
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "生产商(&M)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   300
            TabIndex        =   19
            Tag             =   "生产商"
            Top             =   3030
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "规格(&G)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   480
            TabIndex        =   2
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "规格编码(&N)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "散装单位(&U)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   1830
            Width           =   990
         End
         Begin VB.Label lbl住院单位 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "包装单位(&K)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2805
            TabIndex        =   9
            Top             =   1830
            Width           =   990
         End
      End
      Begin VB.Frame fra 
         Height          =   4515
         Index           =   2
         Left            =   -74850
         TabIndex        =   64
         Top             =   465
         Width           =   4365
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   21
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   89
            Tag             =   "增值税率"
            Top             =   3840
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   11
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   82
            Tag             =   "指导售价"
            Top             =   2628
            Width           =   3030
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   73
            Tag             =   "指导批价"
            Top             =   1419
            Width           =   1455
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   12
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   84
            Tag             =   "指导差价率"
            Text            =   "13.0435"
            Top             =   4545
            Visible         =   0   'False
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   16
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   71
            Tag             =   "当前售价"
            Top             =   1016
            Width           =   3030
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   76
            Tag             =   "采购扣率"
            Text            =   "100"
            Top             =   1822
            Width           =   1455
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   10
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   79
            Tag             =   "结算价"
            Top             =   2225
            Width           =   1455
         End
         Begin VB.ComboBox cbo价格属性 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Tag             =   "价格属性"
            Top             =   210
            Width           =   3090
         End
         Begin VB.TextBox txtEdit 
            Enabled         =   0   'False
            Height          =   300
            Index           =   14
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   88
            Tag             =   "差价让利"
            Text            =   "100"
            Top             =   3434
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   13
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   86
            Tag             =   "加成率"
            Text            =   "15.00"
            Top             =   3031
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   15
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   69
            Tag             =   "成本价格"
            Top             =   613
            Width           =   3030
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   4080
            TabIndex        =   120
            Top             =   3900
            Width           =   90
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   4080
            TabIndex        =   119
            Top             =   3494
            Width           =   90
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   4080
            TabIndex        =   118
            Top             =   3091
            Width           =   90
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4080
            TabIndex        =   117
            Top             =   4605
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "增值税率(&Z)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   90
            TabIndex        =   116
            Top             =   3900
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "指导售价(&K)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   90
            TabIndex        =   81
            Top             =   2688
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "指导批价"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   90
            TabIndex        =   72
            Top             =   1479
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "指导差率(&E)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   90
            TabIndex        =   83
            Top             =   4605
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2700
            TabIndex        =   77
            Top             =   1882
            Width           =   90
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "采购扣率(&X)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   90
            TabIndex        =   75
            Top             =   1882
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "结算价(&T)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   78
            Top             =   2285
            Width           =   810
         End
         Begin VB.Label lbl批价单位 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "元/片"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2700
            TabIndex        =   74
            Top             =   1479
            Width           =   735
         End
         Begin VB.Label lbl批价单位 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "元/片"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2700
            TabIndex        =   80
            Top             =   2285
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "价格属性(&P)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   90
            TabIndex        =   66
            Top             =   270
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "差价让利(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   90
            TabIndex        =   87
            Top             =   3494
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "当前售价(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   90
            TabIndex        =   70
            Top             =   1076
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "加成率"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   90
            TabIndex        =   85
            Top             =   3091
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "成本价格(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   90
            TabIndex        =   68
            Top             =   673
            Width           =   990
         End
      End
      Begin VB.TextBox txt病案费目 
         Height          =   300
         Left            =   -69045
         MaxLength       =   40
         TabIndex        =   96
         ToolTipText     =   "按*打开选择器"
         Top             =   1920
         Width           =   2115
      End
      Begin VB.Label lbl病案费目 
         Caption         =   "病案费目(&F)"
         Height          =   255
         Left            =   -70125
         TabIndex        =   123
         Top             =   1943
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收入项目(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   20
         Left            =   -70125
         TabIndex        =   90
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保类型(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   18
         Left            =   -70125
         TabIndex        =   92
         Top             =   1185
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "应用对象(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   -70125
         TabIndex        =   94
         Top             =   1575
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8400
      Top             =   8040
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
            Picture         =   "frmStuffSpec.frx":0494
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":0A2E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":0FC8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":1562
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl品种说明 
      Caption         =   "编码:0201    品名：一次性针管         英文名称:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   750
      TabIndex        =   111
      Top             =   240
      Width           =   7965
   End
   Begin VB.Label lbl 
      Caption         =   $"frmStuffSpec.frx":1AFC
      Height          =   390
      Index           =   0
      Left            =   -150
      TabIndex        =   103
      Top             =   8640
      Width           =   7125
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmStuffSpec.frx":1B83
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "frmStuffSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng诊疗ID As Long
Dim mstr材料ID As String         '当前编辑的材料ID
Private mlng分类id As Long       '当前选择的分类id

Dim mintSuccess As Integer
Dim mintEditType As gEditType    '编辑类型
Dim mblnChange As Boolean
Dim mstrPrivs As String         '权限串
Dim mblnFrist As Boolean        '第一次运行系统时
Dim mintCount As Integer
Dim mstr产地 As String
Dim mintUnit As Integer     '0-散装单位,1-包装单位
Dim mintCodeLength As Integer   '编码的长度,从数据库中读取出来的长度
Private Const mlngModule = 1711
Private mblnLoad As Boolean      '窗体只active一次
Private mintSet分批 As Integer  '设置分批属性
Private mblnInStrument As Boolean '是否共享安装了物资系统
Private mstr注册证号 As String   '纪录修改之前的注册证号
Private mstr注册证有效期 As String  '纪录修改之前的注册证有效期
Private mint注册修改参数 As Integer '0-只修改当前规格，1-同步修改品种下所有注册证号和注册证有效期

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Public Function ShowEditCard(ByVal frmMain As Object, intEditType As gEditType, ByVal lng诊疗ID As Long, ByVal lng分类id As Long, _
    Optional str材料ID As String = "", Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:编辑卫生材料
    '--入参数:frmMain-调用的主窗体
    '--       intEditType -编辑类型
    '--       lng诊疗ID-诊疗ID(品种ID)
    '--       str材料ID-编辑档案的当前ID
    '--       strPrivs-权限串
    '--出参数:
    '--返  回:编辑成功,返回ture,否则false
    '--编制:刘兴宏
    '--日期:2007/05/25
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    mintEditType = intEditType: mstrPrivs = strPrivs: mstr材料ID = str材料ID: mlng诊疗ID = lng诊疗ID: mlng分类id = lng分类id
    mintSuccess = 0
    
    frmStuffSpec.Show 1, frmMain
    
    ShowEditCard = mintSuccess > 0
End Function

Private Function GetDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:检查数据依赖性
    '--入参数:
    '--出参数:
    '--返  回:存在返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    gstrSQL = "Select 编码||'-'||名称 From 材料来源分类 Order By 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption

    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgBox "未设置材料来源分类（字典管理）！"
            Exit Function
        End If
        Me.cbo材料来源.Clear
        Do While Not .EOF
            Me.cbo材料来源.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    If Me.cbo材料来源.ListCount > 0 Then Me.cbo材料来源.ListIndex = 0
    
     
    gstrSQL = "Select 编码||'-'||名称 From 费用类型 where 性质=1 Order By 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     
    With rsTemp
        Me.cbo费用类型.Clear
        If .RecordCount = 0 Then
            ShowMsgBox "未设置用于卫材的医保类型（字典管理）！"
            Exit Function
        End If
        Do While Not .EOF
            Me.cbo费用类型.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    
    '刘兴宏:2007/05/25:增加材质分类
    gstrSQL = "Select 编码||'-'||名称 as 名称,简码 From 材料材质分类  order by 编码 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo材质分类.Clear
        Do While Not .EOF
            Me.cbo材质分类.AddItem zlStr.nvl(!名称)
            .MoveNext
        Loop
        If cbo材质分类.ListCount <> 0 Then cbo材质分类.ListIndex = 0
    End With
    
    '刘兴宏:2007/05/25:材料存储条件
    gstrSQL = "Select 编码||'-'||名称 as 名称,简码 From 材料存储条件 order by 编码 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo存储条件.Clear
        Do While Not .EOF
            Me.cbo存储条件.AddItem zlStr.nvl(!名称)
            .MoveNext
        Loop
        If cbo存储条件.ListCount <> 0 Then cbo存储条件.ListIndex = 0
    End With
    
    gstrSQL = "Select 编码||'-'||名称 as 名称,简码 From 材料货源情况  order by 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo货源.Clear
        Do While Not .EOF
            Me.cbo货源.AddItem zlStr.nvl(!名称)
            .MoveNext
        Loop
        cbo货源.ListIndex = 0
    End With
    
    If Me.cbo费用类型.ListCount > 0 Then Me.cbo费用类型.ListIndex = 0
    
    gstrSQL = "" & _
        "   Select ID,'['||编码||']'||名称 as 名称" & _
        "   From 收入项目" & _
        "   where 末级=1 and (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
        "   Order By 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgBox "未设置明细的收入项目！"
            Exit Function
        End If
        Me.cbo收入项目.Clear
        Do While Not .EOF
            Me.cbo收入项目.AddItem !名称: Me.cbo收入项目.ItemData(Me.cbo收入项目.NewIndex) = !Id
            .MoveNext
        Loop
        If Me.cbo收入项目.ListCount > 0 Then Me.cbo收入项目.ListIndex = 0
    End With
    
    mintUnit = Get定价单位
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    
   'mstrFormat = GetFmtString(mintUnit) 'IIf(mintUnit = 1, "#####0.0000;-#####0.0000; ;", "#####0.0000000;-#####0.0000000; ;")
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbo材料来源_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chk植入耗材_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo材质分类_Change()
    mblnChange = True
    
End Sub

Private Sub cbo材质分类_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cbo存储条件_Change()
    mblnChange = True
End Sub

Private Sub cbo存储条件_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub

Private Sub cbo单位_Click(Index As Integer)
    Call cbo单位_Change(Index)
End Sub

Private Sub cbo单位_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo费用类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub cbo服务对象_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo货源_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub cbo价格属性_Click()
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If txtEdit(i).Tag = "差价让利" Then
            txtEdit(14).Enabled = InStr(1, mstrPrivs, ";指导价格管理;") <> 0 And Not (cbo价格属性.Text = "定价")
            If txtEdit(14).Enabled Then
                txtEdit(14).BackColor = &H80000005
            Else
                txtEdit(14).BackColor = &H8000000F
            End If
            Exit For
        End If
    Next
End Sub


Private Sub cbo价格属性_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub


Private Sub cbo收入项目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub



Private Sub chkCostly_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandle
    If chkCostly.Value = 0 Then
        strSql = "select count(*) rec from 药品收发记录 a, 收发记录补充信息 b where a.药品id=[1] and a.id=b.收发id "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr材料ID)
        If rsTmp!rec > 0 Then
            rsTmp.Close
            If MsgBox("取消“高值材料”属性将使“卫材外购入库”中不能显示、录入“高值材料”信息。请确定吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                chkCostly.Value = 1
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk跟踪_Click()
    If mintEditType = g查看 Then Exit Sub
    If chk跟踪.Enabled = False Then Exit Sub
    
    If chk跟踪.Value = 1 Then
        chk核算材料.Enabled = True
    Else
        chk核算材料.Enabled = False
        chk核算材料.Value = 0
    End If
End Sub

Private Sub chk跟踪_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
    
End Sub

Private Sub chk跟踪病人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

 

Private Sub chk核算材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk屏蔽费别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk无菌性材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chk在用_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk库房_Click()
    Dim blnEnable As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '在库房分批的前提下，如果发料部门没有库存，则可设置其是否分批
    
    '    gstrSQL = "" & _
    '            "   Select nvl(Count(*),0) " & _
    '            "   From 药品库存 A,部门性质说明 B" & _
    '            "  Where A.药品ID=[1]" & _
    '            "       And A.库房ID=B.部门ID And (B.工作性质 Like '发料部门' Or B.工作性质 Like '%制剂室' )"
    '
    '    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
    '
    '    With rsTemp
    '        blnEnable = True
    '        If .Fields(0).Value <> 0 Then
    '            blnEnable = False
    '        End If
    '    End With
    If Me.chk库房.Value = 0 Then
        Me.chk在用.Value = 0: Me.chk在用.Enabled = False
        'Me.chk效期.Value = 0: Me.chk效期.Enabled = False
        Me.txtEdit(GetTxtIdx("保存期")).Text = "": Me.txtEdit(GetTxtIdx("保存期")).Enabled = False
    Else
        Me.chk在用.Enabled = True
        Me.txtEdit(GetTxtIdx("保存期")).Enabled = True
    End If
    SetCtlBackColor txtEdit(GetTxtIdx("保存期"))
End Sub

Private Function GetTxtIdx(ByVal strName As String) As Integer
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取文本框的索引
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If Trim(txtEdit(i).Tag) = strName Then
            GetTxtIdx = i
            Exit Function
        End If
    Next
    GetTxtIdx = -1
End Function

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:合法,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    Dim strName As String
    Dim bln不强制控制指导价格 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    bln不强制控制指导价格 = ISCHECK不强制控制指导价格
    
    ISValied = False
    
    '设置跟踪在用属性是否允许修改
    If mintEditType = g修改 Then
        '不跟踪在用->跟踪在用，检查门诊/住院费用记录表，有则不能修改
        If Me.chk跟踪.Value = 1 And chk跟踪.Tag = 0 Then
            gstrSQL = "Select 1 " & _
                " From (Select 1 From 门诊费用记录 Where 收费细目id = [1] " & _
                " Union All " & _
                " Select 1 From 住院费用记录 Where 收费细目id = [1]) " & _
                " Where Rownum < 2 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
            
            If Not rsTemp.EOF Then
                MsgBox "该规格已经产生过费用记录，所以不能修改【跟踪在用】属性，" & vbCrLf & _
                "请取消勾选后再保存！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '跟踪在用->不跟踪在用  检查药品收发记录，有则不能修改
        If Me.chk跟踪.Value = 0 And chk跟踪.Tag = 1 Then
            If Not Me.cbo价格属性.Enabled Then
                MsgBox "该规格已经产生过收发记录，所以不能修改【跟踪在用】属性，" & vbCrLf & _
                "请勾选后再保存！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    For i = 0 To txtEdit.UBound
        strName = txtEdit(i).Tag
        strTmp = Trim(txtEdit(i).Text)
        Select Case strName
        Case "规格编码", "规格", "换算系数", "名称"
            If strTmp = "" Then
                ShowMsgBox strName & "未输入，请输入" & strName & "！"
                Me.stbSpec.Tab = 0
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case "采购扣率"  ',"指导差价率""成本价格",
                '刘兴宏:主要是解决成本价格可以为零的情况,比如：疫苗.是免费的
                '问题:9569 2006-11-20
                If Val(strTmp) = 0 And txtEdit(i).Enabled Then
                    ShowMsgBox strName & "为0或未输入，请输入" & strName & "！"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
        Case "指导差价率"
            '刘兴宏:取掉指导差价率的限制,提示是否可输入零
            '东莞塘厦医院：允许将药品和卫材的指导差率和加成率设置为0.医院对部分药品和卫材在医院实行成本价销售,目前不能直接在目录里将加成率设置为0,但是可以在入库的时候修改为0.
            If strTmp = "" And txtEdit(i).Enabled Then
                If MsgBox(strName & "未输入，且将自动设置为0。" & vbCrLf & "是否继续保存？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
        Case "指导批价", "指导售价"
            If bln不强制控制指导价格 = False Then
                If Val(strTmp) = 0 And txtEdit(i).Enabled Then
                    ShowMsgBox strName & "为0或未输入，请输入" & strName & "！"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
'        Case "拼音简码"
'            Me.txtEdit(i).Text = zlStr.GetCodeByORCL(Me.txtEdit(GetTxtIdx("商品名")).Text, 0)
'        Case "五笔简码"
'            Me.txtEdit(i).Text = zlStr.GetCodeByORCL(Me.txtEdit(GetTxtIdx("商品名")).Text, 1)
        Case Else
            
        End Select
        
        If txtEdit(i).MaxLength <> 0 Then
            If LenB(StrConv(strTmp, vbFromUnicode)) > txtEdit(i).MaxLength Then
                ShowMsgBox strName & "超长,最多" & txtEdit(i).MaxLength & "个字符(" & txtEdit(i).MaxLength / 2 & "个汉字)！"
                Me.stbSpec.Tab = 0
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        Select Case strName
        Case "换算系数", "指导批价", "指导售价", "成本价格"
            If Val(strTmp) > 1000000 Then
                ShowMsgBox strName & "超过最大值1000000！"
                Me.stbSpec.Tab = IIf(strName = "换算系数", 0, 1)
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
            If strName = "换算系数" And Val(strTmp) <= 0 Then
                ShowMsgBox strName & "必需大于零！"
                Me.stbSpec.Tab = IIf(strName = "换算系数", 0, 1)
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
            
        Case "指导差价率", "差价让利", "采购扣率", "增值税率"
            If Val(strTmp) > 100 Then
                ShowMsgBox strName & "不能超过100！"
                Me.stbSpec.Tab = 1
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case "当前售价"
            If Me.cbo价格属性.ItemData(cbo价格属性.ListIndex) = 0 Then
                If Abs(Val(strTmp)) > 1000000 Then
                    ShowMsgBox "当前售价超过最大范围-1000000~1000000！"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                If bln不强制控制指导价格 = False Then
                    If Val(strTmp) > Val(Me.txtEdit(GetTxtIdx("指导售价"))) Then
                        ShowMsgBox "售价不能高于指导零售价！"
                        Me.stbSpec.Tab = 1
                        If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                        Exit Function
                    End If
                End If
            End If
        Case Else
        End Select
    Next
    
    '设置条码管理必须要分批管理
    If chkCode.Value = 1 Then
        If chk库房.Value = 0 Or chk在用.Value = 0 Then
            Me.stbSpec.Tab = 1
            MsgBox "启用条码管理，必须设置卫材为分批管理！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If chk跟踪.Value = 0 And chk核算材料.Value = 1 Then
        ShowMsgBox "非跟踪材料不能设置核算材料,请检查!:"
         Me.stbSpec.Tab = 1
         If chk核算材料.Enabled = True Then chk核算材料.SetFocus
        Exit Function
    End If
    
    If LenB(StrConv(Me.txt注册商标.Text, vbFromUnicode)) > 50 Then
        MsgBox "注册商标超长，最多50个字符或25个汉字！", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt注册商标.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txt批准文号.Text, vbFromUnicode)) > 40 Then
        MsgBox "批准文号超长，最多40个字符或20个汉字！", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt批准文号.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txt注册证号.Text, vbFromUnicode)) > 50 Then
        MsgBox "注册证号超长，最多50个字符或25个汉字！", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt注册证号.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txt备选码.Text, vbFromUnicode)) > 20 Then
        MsgBox "备选码超长，最多20个字符或10个汉字！", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt备选码.SetFocus
        Exit Function
    End If
    If Trim(Me.cbo单位(0).Text) = "" Then ShowMsgBox "请输入散装单位！": Me.stbSpec.Tab = 0: Me.cbo单位(0).SetFocus: Exit Function
    If LenB(StrConv(Me.cbo单位(0).Text, vbFromUnicode)) > 6 Then ShowMsgBox "散装单位超长(最多6个字符或3个汉字)！": Me.stbSpec.Tab = 0: Me.cbo单位(0).SetFocus: Exit Function
    If Trim(Me.cbo单位(1).Text) = "" Then ShowMsgBox "请输入包装单位！": Me.stbSpec.Tab = 0: Me.cbo单位(1).SetFocus: Exit Function
    If LenB(StrConv(Me.cbo单位(1).Text, vbFromUnicode)) > 6 Then ShowMsgBox "包装单位超长(最多6个字符或3个汉字)！": Me.stbSpec.Tab = 0: Me.cbo单位(1).SetFocus: Exit Function
    ISValied = True

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存卡片数据
    '--入参数:
    '--出参数:
    '--返  回:保存成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl当前售价 As Double, dbl指导售价 As Double, dbl成本价格 As Double, dbl指导批价 As Double
    Dim str站点 As String
    Dim lng材料ID As Long
    Dim lng分类id As Long
    Dim str货源 As String
    Dim str来源 As String
    Dim strValues As String
    
    str货源 = Trim(cbo货源.Text)
    If str货源 <> "" Then
        str货源 = Mid(str货源, InStr(1, str货源, "-") + 1)
    End If
    
    str来源 = Trim(cbo材料来源.Text)
    If str来源 <> "" Then
        str来源 = Mid(str来源, InStr(1, str来源, "-") + 1)
    End If
    
    err = 0
    On Error GoTo ErrHand:
    
    '------------------------------------------
    '数据保存
    If mintUnit <> 0 Then
        dbl指导售价 = Round(Val(txtEdit(11).Text) / Val(txtEdit(GetTxtIdx("换算系数")).Text), g_小数位数.obj_最大小数.零售价小数)
        dbl当前售价 = Round(Val(txtEdit(16).Text) / Val(txtEdit(GetTxtIdx("换算系数")).Text), g_小数位数.obj_最大小数.零售价小数)
        dbl成本价格 = Round(Val(txtEdit(15).Text) / Val(txtEdit(GetTxtIdx("换算系数")).Text), g_小数位数.obj_最大小数.成本价小数)
        dbl指导批价 = Round(Val(txtEdit(8).Text) / Val(txtEdit(GetTxtIdx("换算系数")).Text), g_小数位数.obj_最大小数.成本价小数)
    Else
        dbl当前售价 = Round(Val(txtEdit(16).Text), g_小数位数.obj_最大小数.零售价小数)
        dbl指导售价 = Round(Val(txtEdit(11).Text), g_小数位数.obj_最大小数.零售价小数)
        dbl成本价格 = Round(Val(txtEdit(15).Text), g_小数位数.obj_最大小数.成本价小数)
        dbl指导批价 = Round(Val(txtEdit(8).Text), g_小数位数.obj_最大小数.成本价小数)
    End If
    If mintEditType = g新增 Then
        lng材料ID = sys.NextId("收费项目目录")
        gstrSQL = "zl_卫生材料_Insert("
    Else
        lng材料ID = Val(mstr材料ID)
        gstrSQL = "zl_卫生材料_UPdate("
    End If
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '存储过程相关参数
    ' zl_卫生材料_Insert  or zl_卫生材料_UPdate的参数
    '  诊疗id_In       In 材料特性.诊疗id%Type,
    gstrSQL = gstrSQL & mlng诊疗ID & ","
    '  材料id_In       In 材料特性.材料id%Type,
    gstrSQL = gstrSQL & lng材料ID & ","
    '  编码_In         In 收费项目目录.编码%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(GetTxtIdx("规格编码")).Text) & "',"
    '  规格_In         In 收费项目目录.规格%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(GetTxtIdx("规格")).Text) & "',"
    '  产地_In         In 收费项目目录.产地%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("生产商")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  标识主码_In     In 收费项目目录.标识主码%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("标识主码")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  标识子码_In     In 收费项目目录.标识子码%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("标识子码")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  备选码_In       In 收费项目目录.备选码%Type := Null,
    strValues = Trim(txt备选码.Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  材料来源_In     In 材料特性.材料来源%Type := Null,
    gstrSQL = gstrSQL & IIf(str来源 = "", "NULL", "'" & str来源 & "'") & ","
    '  货源情况_In     In 材料特性.货源情况%Type := Null,
    gstrSQL = gstrSQL & IIf(str货源 = "", "NULL", "'" & str货源 & "'") & ","
    '  散装单位_In     In 收费项目目录.计算单位%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo单位(0).Text) = "", "NULL", "'" & Trim(cbo单位(0).Text) & "'") & ","
    '  包装单位_In     In 材料特性.包装单位%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo单位(1).Text) = "", "NULL", "'" & Trim(cbo单位(1).Text) & "'") & ","
    '  换算系数_In     In 材料特性.换算系数%Type := Null,
    strValues = Val(txtEdit(GetTxtIdx("换算系数")).Text):
    gstrSQL = gstrSQL & strValues & ","
    '  是否变价_In     In 收费项目目录.是否变价%Type := Null,
    gstrSQL = gstrSQL & IIf(cbo价格属性.ItemData(cbo价格属性.ListIndex) = 0, 0, 1) & ","
    '  指导批发价_In   In 材料特性.指导批发价%Type := Null,
    gstrSQL = gstrSQL & dbl指导批价 & ","
    '  扣率_In         In 材料特性.扣率%Type := 95,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("采购扣率")).Text) & ","
    '  指导零售价_In   In 材料特性.指导零售价%Type := Null,
    gstrSQL = gstrSQL & dbl指导售价 & ","
    '  指导差价率_In   In 材料特性.指导差价率%Type := Null,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("指导差价率")).Text) & ","
    '  费用类型_In     In 收费项目目录.费用类型%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo费用类型.Text) = "", "NULL", "'" & Mid(Me.cbo费用类型.Text, InStr(1, Me.cbo费用类型.Text, "-") + 1) & "'") & ","
    '  服务对象_In     In 收费项目目录.服务对象%Type := Null,
    gstrSQL = gstrSQL & cbo服务对象.ItemData(cbo服务对象.ListIndex) & ","
    '  屏蔽费别_In     In 收费项目目录.屏蔽费别%Type := 0,
    gstrSQL = gstrSQL & IIf(chk屏蔽费别.Value = 1, 1, 0) & ","
    '  库房分批_In     In 材料特性.库房分批%Type := Null,
    gstrSQL = gstrSQL & IIf(chk库房.Value = 1, 1, 0) & ","
    '  在用分批_In     In 材料特性.在用分批%Type := Null,
    gstrSQL = gstrSQL & IIf(chk在用.Value = 1, 1, 0) & ","
    '  最大效期_In     In 材料特性.最大效期%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("保存期")).Text)
    gstrSQL = gstrSQL & IIf(Val(strValues) <> 0, Val(strValues), "NULL") & ","
    '  灭菌效期_In     In 材料特性.灭菌效期%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("灭菌效期")).Text)
    gstrSQL = gstrSQL & IIf(Val(strValues) <> 0, Val(strValues), "NULL") & ","
    '  无菌性材料_In   In 材料特性.无菌性材料%Type := Null,
    gstrSQL = gstrSQL & IIf(chk无菌性材料.Value = 1, 1, 0) & ","
    '  一次性材料_In   In 材料特性.一次性材料%Type := Null,
    gstrSQL = gstrSQL & IIf(Chk一次性材料.Value = 1, 1, 0) & ","
    '  原材料_In       In 材料特性.原材料%Type := Null,
    gstrSQL = gstrSQL & IIf(chk原料.Value = 1, 1, 0) & ","
    '  差价让利比_In   In 材料特性.差价让利比%Type := 0,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("差价让利")).Text) & ","
    '  成本价_In       In 材料特性.成本价%Type := 0,
    gstrSQL = gstrSQL & dbl成本价格 & ","
    '  跟踪在用_In     In 材料特性.跟踪在用%Type := Null,
    gstrSQL = gstrSQL & chk跟踪.Value & ","
    '  核算材料_In     In 材料特性.核算材料%Type := 0,
    gstrSQL = gstrSQL & IIf(chk跟踪.Value = 1, chk核算材料.Value, 0) & ","
    '  当前售价_In     In 收费价目.现价%Type := 0,
    gstrSQL = gstrSQL & dbl当前售价 & ","
    '  收入id_In       In 收费价目.收入项目id%Type := Null,
    gstrSQL = gstrSQL & cbo收入项目.ItemData(cbo收入项目.ListIndex) & ","
    '  批准文号_In     In 材料特性.批准文号%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txt批准文号.Text) = "", "NULL", "'" & Trim(txt批准文号.Text) & "'") & ","
    '  注册商标_In     In 材料特性.注册商标%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txt注册商标.Text) = "", "NULL", "'" & Trim(txt注册商标.Text) & "'") & ","
    '  注册证号_In     In 材料特性.注册证号%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txt注册证号.Text) = "", "NULL", "'" & Trim(txt注册证号.Text) & "'") & ","
    '  许可证号_In     In 材料特性.许可证号%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("许可证号")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  许可证有效期_In In 材料特性.许可证有效期%Type := Null,
    If dtp许可证效期.Value = "" Or IsNull(dtp许可证效期.Value) Then
        gstrSQL = gstrSQL & "NULL" & ","
    Else
        gstrSQL = gstrSQL & "To_date('" & Format(dtp许可证效期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    End If
    '  材质分类_In     In 材料特性.材质分类%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo材质分类.Text) = "", "NULL", "'" & Mid(Me.cbo材质分类.Text, InStr(1, Me.cbo材质分类.Text, "-") + 1) & "'") & ","
    '  存储条件_In     In 材料特性.存储条件%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(cbo存储条件.Text) = "", "NULL", "'" & Mid(Me.cbo存储条件.Text, InStr(1, Me.cbo存储条件.Text, "-") + 1) & "'") & ","
    '  跟踪病人_In     In 材料特性.跟踪病人%Type := 0
    gstrSQL = gstrSQL & IIf(chk跟踪病人.Value = 1, "1", "0") & ","
    '  站点_In         In 收费项目目录.站点%Type := Null
    gstrSQL = gstrSQL & IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & str站点 & "'", "NULL") & ","
    '  品名_In         In 收费项目别名.名称%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("商品名")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("商品名")).Text) & "'") & ","
    '  拼音_In         In 收费项目别名.简码%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("拼音简码")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("拼音简码")).Text) & "'") & ","
    '  五笔_In         In 收费项目别名.简码%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("五笔简码")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("五笔简码")).Text) & "'") & ","
    '  增值税率_In     In 材料特性.增值税率%Type := Null
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("增值税率")).Text) & ","
    '  说明_In         In 收费项目目录.说明%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("说明")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("说明")).Text) & "'") & ","
    '  高值材料        In 材料特性.高值材料%Type := Null
    gstrSQL = gstrSQL & IIf(chkCostly.Value = 1, 1, 0) & ","
    '  条码管理        In 材料特性.是否条码管理%Type := Null
    gstrSQL = gstrSQL & IIf(chkCode.Value = 1, 1, 0) & ",'"
    '   病案费目
    gstrSQL = gstrSQL & txt病案费目.Text & "',"
    '   器械包单件
    gstrSQL = gstrSQL & IIf(chkInstrument.Value = 1, chkInstrument.Value, 0) & ","
    '  注册证有效期_In In 材料特性.注册证有效期%Type := Null
    If dtp注册证有效期.Value = "" Or IsNull(dtp注册证有效期.Value) Then
        gstrSQL = gstrSQL & "NULL,"
    Else
        gstrSQL = gstrSQL & "To_date('" & Format(dtp注册证有效期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    End If
    '  是否植入耗材
    gstrSQL = gstrSQL & IIf(chk植入耗材.Value = 1, 1, 0) & ","
    
    If mintEditType = g修改 Then
        gstrSQL = gstrSQL & mint注册修改参数 & ","
    End If
    '  加成率
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("加成率")).Text) & ","
    '  分零使用_In
    gstrSQL = gstrSQL & IIf(chk分零使用.Value = 1, 1, 0)
    gstrSQL = gstrSQL & ")"

    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chk库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk一次性材料_Click()
    If mintEditType = g查看 Then Exit Sub
    If Chk一次性材料.Value = 1 Then
        txtEdit(7).Enabled = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0
    Else
        '只有一次性材料才有灭菌效期。
        txtEdit(7).Enabled = False
        txtEdit(7).Text = ""
    End If
    
    SetCtlBackColor txtEdit(7)
End Sub

Private Sub Chk一次性材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chk原料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt病案费目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkInstrument_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkCostly_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub cmbStationNo_Change()
    mblnChange = True
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    stbSpec.Tab = 1
    If cbo价格属性.Enabled Then cbo价格属性.SetFocus
End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    Call cmdOK_Click
End Sub

Private Sub cmd病案_Click()
    On Error GoTo ErrHandle
    Dim strSql As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    
    strSql = "Select 编码 as id,上级 as 上级id, 名称, 简码, 末级 From 病案费目 Start With 上级 Is Null Connect By Prior 编码 = 上级"
    blnRe = frmTreeLeafSel.ShowTree(strSql, strID, str名称, "病案费目")
    '成功返回
    If blnRe Then
        '新的本级的宽度
        lbl病案费目.Tag = strID
        txt病案费目.Text = str名称
        stbSpec.Tab = 1
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd产地_Click()
    Dim rsTemp As New ADODB.Recordset
    Call Sel产地("")
End Sub
Private Sub cmdOK_Click()
    
    Dim i As Long
    '检查规格页面的输入项是否正确
    If ISValied = False Then Exit Sub
    

    If mintEditType <> g新增 And mintEditType <> g修改 Then
        Unload Me
        Exit Sub
    End If
    
    If mintEditType = g修改 Then
        If mstr注册证号 <> Trim(txt注册证号) Or mstr注册证有效期 <> CStr(Format(dtp注册证有效期.Value, "yyyy-mm-dd")) Then
            If MsgBox("是否把【注册证号】和【注册证有效期】同步修改到该品种下所有规格？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mint注册修改参数 = 1
            Else
                mint注册修改参数 = 0
            End If
        Else
            mint注册修改参数 = 0
        End If
    End If
    
    If SaveData = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    '保存参数值
    Call SaveReg
    
    If mintEditType = g新增 Then
        If ActiveControl Is cmdOK Then   '普通模式
            Unload Me
        ElseIf ActiveControl Is cmdSaveAddSpec Then        '连续增加规格
            For i = 0 To cbo单位(0).ListCount
                If Trim(cbo单位(0).Text) = cbo单位(0).List(i) Then
                    cbo单位(0).ListIndex = i: i = -1: Exit For
                End If
            Next
            If i >= 0 Then
                cbo单位(0).AddItem Trim(cbo单位(0).Text)
                cbo单位(0).ListIndex = cbo单位(0).NewIndex
            End If
            
            Call InitCardData(False)
            
            Me.stbSpec.Tab = 0
            If txtEdit(GetTxtIdx("规格")).Enabled Then txtEdit(GetTxtIdx("规格")).SetFocus
        ElseIf ActiveControl Is cmdSaveAddItem Then '连续增加品种
            Unload Me
            If frmStuffBreed.ShowEditCard(frmStuffMgr, g新增, "", mlng分类id, gstrPrivs) = False Then
                Exit Sub
            End If
        End If
    Else
        Unload Me
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub
Private Sub cmd帮助_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Function GetMaxCode() As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取最大编号
    '--入参数:
    '--出参数:
    '--返  回:最大编号
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCode As ADODB.Recordset
    Dim strTemp As String
    Dim intCodeType As Integer
    Dim str编码 As String
    
    On Error GoTo ErrHandle
    intCodeType = Val(zlDatabase.GetPara("编码递增模式", glngSys, mlngModule))
    gstrSQL = "Select 编码 From 诊疗项目目录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng诊疗ID)
    strTemp = zlStr.nvl(rsTemp!编码)
    
    If intCodeType = 0 Or Len(strTemp) >= 17 Then
        '取最大编码
        gstrSQL = "Select Nvl(编码, '00000000000000') As 编码" & vbNewLine & _
                        "From (Select 编码" & vbNewLine & _
                        "       From 收费项目目录 A, 材料特性 B" & vbNewLine & _
                        "       Where a.类别 = '4' And a.Id = b.材料id" & vbNewLine & _
                        "       Order By Length(编码) Desc, 编码 Desc)" & vbNewLine & _
                        "Where Rownum = 1"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            str编码 = zlCommFun.IncStr(!编码)
            GetMaxCode = str编码
        End With
   
    Else

        gstrSQL = "Select a.Id, c.编码, c.名称, c.规格" & vbNewLine & _
                        "From 诊疗项目目录 A, 材料特性 B, 收费项目目录 C" & vbNewLine & _
                        "Where a.Id = b.诊疗id And b.材料id = c.Id And a.分类id In (Select ID From 诊疗分类目录 Where 类型 = 7) And a.Id =[1] " & vbNewLine & _
                        "Order By ID"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng诊疗ID)
        
        If Len(strTemp) >= 14 Or 14 - Len(strTemp) - IIf(intCodeType = 1, 1, 0) = 0 Then
            str编码 = "01"
            str编码 = IIf(intCodeType = 1, "4", "") & strTemp & str编码
        Else
            str编码 = Mid("00000000000000", 1, 14 - Len(strTemp) - IIf(intCodeType = 1, 1, 0))
            str编码 = IIf(intCodeType = 1, "4", "") & strTemp & str编码
            str编码 = zlCommFun.IncStr(str编码)
        End If
        
        GetMaxCode = str编码
    
        Do While True
            rsTemp.Filter = ""
            rsTemp.Filter = "编码='" & GetMaxCode & "'"
            If rsTemp.RecordCount = 0 Then
                Exit Do
            End If
            GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "Select 编码 From 收费项目目录 "
    Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    Do While True
        rsCode.Filter = ""
        rsCode.Filter = "编码='" & GetMaxCode & "'"
        If rsCode.RecordCount = 0 Then
            Exit Do
        End If
        GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    Loop

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Set分批()
    '库房分批属性设置
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mintSet分批 = 0 Then
        gstrSQL = "Select b.库房分批, b.在用分批" & _
                   " From 材料特性 B, (Select Max(a.Id) As ID From 收费项目目录 A, 材料特性 B Where a.Id = b.材料id And b.诊疗id = [1]) C" & _
                   " Where b.材料id = c.Id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "库房分批设置", mlng诊疗ID)
        
        If rsTemp.RecordCount > 0 Then
            chk库房.Value = IIf(IsNull(rsTemp!库房分批), "0", rsTemp!库房分批)
            chk在用.Value = IIf(IsNull(rsTemp!在用分批), "0", rsTemp!在用分批)
            chk库房.Enabled = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0
            chk在用.Enabled = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0
        End If
    ElseIf mintSet分批 = 1 Then
        chk库房.Value = 1
        chk在用.Value = 0
        chk库房.Enabled = False
        chk在用.Enabled = False
    ElseIf mintSet分批 = 2 Then
        chk库房.Value = 1
        chk在用.Value = 1
        chk库房.Enabled = False
        chk在用.Enabled = False
    ElseIf mintSet分批 = 3 Then
        chk库房.Value = 0
        chk在用.Value = 0
        chk库房.Enabled = False
        chk在用.Enabled = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitCardData(Optional bln单位 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始卡片数据
    '--入参数:bln单位-是否重新获取单位
    '--出参数:
    '--返  回:加载成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim str计算单位 As String
    Dim rsTemp As New ADODB.Recordset
    Dim dbl换算系数 As Double
    '--恢复参数值
    On Error GoTo ErrHandle
    
    Call LoadReg    '加载注册信息值
    If bln单位 Then
        '求出当前单位
        gstrSQL = " Select distinct a.计算单位,b.包装单位 From 收费项目目录 a,材料特性 b where a.id=b.材料id and b.诊疗ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng诊疗ID)
        cbo单位(0).Clear
        cbo单位(1).Clear
        
        With rsTemp
            .Sort = "计算单位"
            str计算单位 = ""
            Do While Not .EOF
                If str计算单位 <> zlStr.nvl(!计算单位) Then
                    str计算单位 = zlStr.nvl(!计算单位)
                    cbo单位(0).AddItem zlStr.nvl(!计算单位)
                End If
                .MoveNext
            Loop
            .Sort = "包装单位"
            str计算单位 = ""
            Do While Not .EOF
                If str计算单位 <> zlStr.nvl(!包装单位) Then
                    str计算单位 = zlStr.nvl(!包装单位)
                    cbo单位(1).AddItem zlStr.nvl(!包装单位)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    '确定分类信息
    gstrSQL = "Select id, 编码,名称,计算单位,站点 from 诊疗项目目录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng诊疗ID)
   If Not rsTemp.EOF Then
        '读取站点信息
        With cmbStationNo
            For i = 1 To .ListCount - 1
                If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!站点) Then
                    .ListIndex = i: Exit For
                End If
            Next
        End With
        
        
        lbl品种说明.Caption = "品种信息：[" & zlStr.nvl(rsTemp!编码) & "] " & zlStr.nvl(rsTemp!名称)
        str计算单位 = zlStr.nvl(rsTemp!计算单位)
        gstrSQL = "Select 名称 From 诊疗项目别名 where 诊疗项目id=[1] and 性质=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng诊疗ID)
        If rsTemp.RecordCount = 0 Then
            lbl品种说明.Caption = lbl品种说明.Caption & Space(8) & "英文名称:"
        Else
            lbl品种说明.Caption = lbl品种说明.Caption & Space(8) & "英文名称:" & zlStr.nvl(rsTemp!名称)
        End If
        For i = 0 To cbo单位(0).ListCount
            If str计算单位 = cbo单位(0).List(i) Then
                If cbo单位(0).ListIndex < 0 Then
                    cbo单位(0).ListIndex = i
                End If
                i = -1
                Exit For
            End If
        Next
        If i <> -1 Then
            cbo单位(0).AddItem str计算单位
            cbo单位(0).ListIndex = cbo单位(0).NewIndex
        End If
        If cbo单位(0).ListIndex >= 0 Then
            str计算单位 = Trim(cbo单位(0).Text)
        End If
   Else
        ShowMsgBox "不存在指定的品种,不能继续!"
        Exit Function
   End If
   
    '--取缺省收入项
   cbo收入项目.Tag = Val(zlDatabase.GetPara("收入项目对应", glngSys, mlngModule))
    For mintCount = 0 To Me.cbo收入项目.ListCount - 1
        If Me.cbo收入项目.ItemData(mintCount) = Val(Me.cbo收入项目.Tag) Then
            Me.cbo收入项目.ListIndex = mintCount: Exit For
        End If
    Next
       
   If mintEditType = g新增 Then
        '增加时，重新提取编码号，清空规格和生产商
        Call 获取上次录入规格信息(mlng诊疗ID)
        Me.txtEdit(GetTxtIdx("规格编码")).Text = "": Me.txtEdit(GetTxtIdx("规格")).Text = "": Me.txtEdit(GetTxtIdx("生产商")).Text = "": Me.lblFound.Caption = ""
        gstrSQL = "Select 编码 from 收费项目目录 where 1=2"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        mintCodeLength = rsTemp.Fields("编码").DefinedSize
        txtEdit(0).MaxLength = mintCodeLength
        Me.txtEdit(0).Text = GetMaxCode
        '默认规格
        With cbo单位(0)
            For i = 0 To .ListCount
                If .List(i) = str计算单位 Then
                    .ListIndex = i
                End If
            Next
            If .ListIndex >= 0 Then
                If .List(.ListIndex) <> str计算单位 Then
                    .AddItem str计算单位
                End If
            End If
        End With
        dtp许可证效期.Value = sys.Currentdate
        dtp许可证效期.Value = ""
        dtp注册证有效期.Value = sys.Currentdate
        dtp注册证有效期.Value = ""
        
        Call Set分批
        
        Exit Function
   End If

   '其他需读取卡片数据
    '----------数据装载-------------------------------------
    gstrSQL = "select I.编码 as 规格编码,I.名称,I.规格,I.产地 as 生产商,S.货源情况,S.材料来源, " & _
             "        I.计算单位,S.换算系数,S.包装单位,I.是否变价,S.指导批发价 as 指导批价,S.扣率 as 采购扣率,S.指导零售价 as 指导售价," & _
             "        S.指导差价率,S.加成率,S.差价让利比 as 差价让利,S.成本价 as 成本价格, " & _
             "        I.标识主码,I.标识子码,i.病案费目,I.备选码,I.费用类型,I.服务对象,I.屏蔽费别, " & _
             "        S.库房分批,S.在用分批,S.最大效期 as 保存期,S.灭菌效期,S.无菌性材料," & _
             "        S.一次性材料,S.原材料,S.批准文号,S.注册商标,S.注册证号,s.注册证有效期,S.高值材料,S.是否条码管理," & IIf(mblnInStrument, " s.器械包卫材单件, ", "") & _
             "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间," & _
             "        S.跟踪在用,S.核算材料,S.许可证号,S.许可证有效期,S.材质分类,S.存储条件,S.跟踪病人,I.站点,S.增值税率,I.说明,S.是否植入耗材,S.是否分零  " & _
             "  from 收费项目目录 I,材料特性 S " & _
             "  where I.ID=S.材料ID and I.id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
             
    
    Dim strFieldsName As String
    
    With rsTemp
        mintCodeLength = .Fields("规格编码").DefinedSize
        txtEdit(0).MaxLength = mintCodeLength
        If .RecordCount > 0 Then
            txt批准文号.Text = zlStr.nvl(!批准文号)
            txt注册商标.Text = zlStr.nvl(!注册商标)
            txt注册证号.Text = zlStr.nvl(!注册证号)
            mstr注册证号 = zlStr.nvl(!注册证号)
            txt备选码.Text = zlStr.nvl(!备选码)
            For i = 0 To txtEdit.UBound
                strFieldsName = txtEdit(i).Tag
                Select Case strFieldsName
                Case "灭菌效期", "保存期"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), "######")
                Case "指导批价", "成本价格"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!换算系数, 1)), mFMT.FM_成本价)
                Case "指导售价"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!换算系数, 1)), mFMT.FM_零售价)
                Case "采购扣率", "指导差价率", "差价让利", "加成率"
                        txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBCJL)
                Case "拼音", "五笔", "结算价", "当前售价"
                Case "许可证号"
                    txtEdit(i).Text = zlStr.nvl(!许可证号)
                Case "许可证有效期"
                Case "增值税率"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBJCL)
                Case Else
                    '“商品名,拼音简码,五笔简码”单独处理
                    If InStr("商品名;拼音简码;五笔简码", strFieldsName) = 0 Then txtEdit(i).Text = zlStr.nvl(.Fields(strFieldsName))
                End Select
            Next
            chk跟踪病人.Value = IIf(Val(zlStr.nvl(!跟踪病人)) = 1, 1, 0)
            chk核算材料.Value = IIf(Val(zlStr.nvl(!核算材料)) = 1, 1, 0)
            Me.chk核算材料.Enabled = (Me.chk跟踪.Value = 1)
            
            If IsNull(!许可证有效期) Then
                dtp许可证效期.Value = ""
            Else
                dtp许可证效期.Value = Format(!许可证有效期, "yyyy-mm-dd")
            End If
            If IsNull(!注册证有效期) Then
                dtp注册证有效期.Value = ""
                mstr注册证有效期 = ""
            Else
                dtp注册证有效期.Value = Format(!注册证有效期, "yyyy-mm-dd")
                mstr注册证有效期 = Format(!注册证有效期, "yyyy-mm-dd")
            End If
            
            '计算单位
            For mintCount = 0 To Me.cbo单位(0).ListCount - 1
                If cbo单位(0).List(mintCount) = zlStr.nvl(!计算单位) Then
                    cbo单位(0).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo单位(0).ListIndex < 0 Then
                If zlStr.nvl(!计算单位) <> "" Then
                    cbo单位(0).AddItem zlStr.nvl(!计算单位)
                    cbo单位(0).ListIndex = cbo单位(0).NewIndex
                End If
            End If
            '包装单位
            For mintCount = 0 To Me.cbo单位(1).ListCount - 1
                If cbo单位(1).List(mintCount) = zlStr.nvl(!包装单位) Then
                    cbo单位(1).ListIndex = mintCount
                    Exit For
                End If
            Next
            If cbo单位(1).ListIndex < 0 Then
                If zlStr.nvl(!包装单位) <> "" Then
                    cbo单位(1).AddItem zlStr.nvl(!计算单位)
                    cbo单位(1).ListIndex = cbo单位(1).NewIndex
                End If
            End If
            dbl换算系数 = zlStr.nvl(!换算系数, 1)
            '--材料来源
            For mintCount = 0 To Me.cbo货源.ListCount - 1
                If Mid(Me.cbo货源.List(mintCount), InStr(1, Me.cbo货源.List(mintCount), "-") + 1) = zlStr.nvl(!货源情况) Then
                    Me.cbo货源.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo材料来源.ListCount - 1
                If Mid(Me.cbo材料来源.List(mintCount), InStr(1, Me.cbo材料来源.List(mintCount), "-") + 1) = zlStr.nvl(!材料来源) Then
                    Me.cbo材料来源.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo材质分类.ListCount - 1
                If Mid(Me.cbo材质分类.List(mintCount), InStr(1, Me.cbo材质分类.List(mintCount), "-") + 1) = zlStr.nvl(!材质分类) Then
                    Me.cbo材质分类.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo存储条件.ListCount - 1
                If Mid(Me.cbo存储条件.List(mintCount), InStr(1, Me.cbo存储条件.List(mintCount), "-") + 1) = zlStr.nvl(!存储条件) Then
                    Me.cbo存储条件.ListIndex = mintCount: Exit For
                End If
            Next
            
            
          '时价
            For mintCount = 0 To Me.cbo价格属性.ListCount - 1
                If cbo价格属性.ItemData(mintCount) = zlStr.nvl(!是否变价, 0) Then
                    cbo价格属性.ListIndex = mintCount
                    Exit For
                End If
            Next
            
            lbl批价单位(0).Caption = "元/" & IIf(mintUnit = 0, zlStr.nvl(!计算单位), zlStr.nvl(!包装单位))
            lbl批价单位(1).Caption = "元/" & IIf(mintUnit = 0, zlStr.nvl(!计算单位), zlStr.nvl(!包装单位))
            
            cbo价格属性.ListIndex = IIf(IsNull(!是否变价), 0, !是否变价)
            
            Chk一次性材料.Value = zlStr.nvl(!一次性材料, 0): Call Chk一次性材料_Click
            chk无菌性材料.Value = zlStr.nvl(!无菌性材料, 0)
            chk原料.Value = zlStr.nvl(!原材料, 0)
            chk植入耗材.Value = zlStr.nvl(!是否植入耗材, 0)
            
            For mintCount = 0 To Me.cbo费用类型.ListCount - 1
                If Mid(Me.cbo费用类型.List(mintCount), InStr(1, Me.cbo费用类型.List(mintCount), "-") + 1) = IIf(IsNull(!费用类型), "", !费用类型) Then
                    Me.cbo费用类型.ListIndex = mintCount: Exit For
                End If
            Next
            
            Me.cbo服务对象.ListIndex = IIf(IsNull(!服务对象), 0, !服务对象)
            Me.chk屏蔽费别.Value = IIf(IsNull(!屏蔽费别), 0, !屏蔽费别)
            
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "注：该规格于" & Format(!建档时间, "YYYY年MM月DD日") & "建立，" & Format(!撤档时间, "YYYY年MM月DD日") & "停用"
            Else
                Me.lblFound.Caption = ""
            End If
            
            Me.chk库房.Value = zlStr.nvl(!库房分批, 0)
            Me.chk在用.Value = zlStr.nvl(!在用分批, 0)
            Me.chk跟踪.Value = zlStr.nvl(!跟踪在用, 0)
            Me.chkCostly.Value = zlStr.nvl(!高值材料, 0)
            Me.chkCode.Value = zlStr.nvl(!是否条码管理, 0)
            Me.chk分零使用.Value = zlStr.nvl(!是否分零, 0)
            If mblnInStrument = True Then
                Me.chkInstrument.Value = zlStr.nvl(!器械包卫材单件, 0)
            End If

            Me.chk在用.Tag = Me.chk在用.Value
            txt病案费目.Text = IIf(IsNull(!病案费目), "", !病案费目)
            
            If Me.chk库房.Value = 0 Then
                Me.chk在用.Enabled = False: Me.chk在用.Value = 0
            Else
                Me.chk在用.Enabled = True
                Me.chk在用.Value = Me.chk在用.Tag
            End If
            
            '读取站点信息
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!站点) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
            
        End If
        
   End With
         
    '提取商品名和简码
    gstrSQL = "select 名称,性质,简码,码类 from 收费项目别名 where 收费细目id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
    With rsTemp
        Me.txtEdit(19).MaxLength = .Fields("简码").DefinedSize
        Me.txtEdit(20).MaxLength = .Fields("简码").DefinedSize
        Me.txtEdit(GetTxtIdx("商品名")).MaxLength = .Fields("名称").DefinedSize
        Do While Not .EOF
            If !性质 = 3 And !码类 = 1 Then
                Me.txtEdit(GetTxtIdx("商品名")).Text = IIf(IsNull(!名称), "", !名称)
                Me.txtEdit(GetTxtIdx("拼音简码")).Text = IIf(IsNull(!简码), "", !简码)
            End If
            If !性质 = 3 And !码类 = 2 Then
                Me.txtEdit(GetTxtIdx("商品名")).Text = IIf(IsNull(!名称), "", !名称)
                Me.txtEdit(GetTxtIdx("五笔简码")) = IIf(IsNull(!简码), "", !简码)
            End If
            .MoveNext
        Loop
    End With
         
    '提取显示当前售价
    If Me.cbo价格属性.ListIndex <> 0 Then
        '时价材料，取库存金额/库存数量做为其价格，无库存时取价表定价
        gstrSQL = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                " from 收费价目 P," & _
                "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                "      From 药品库存 Where 药品ID=[1]) K" & _
                " where P.收费细目id=[1] and (Sysdate Between p.执行日期 And p.终止日期 or Sysdate>=p.执行日期 And p.终止日期 Is Null)" & _
                GetPriceClassString("P")
    Else
        '非时价材料调价，取其价格记录中的价格
        gstrSQL = "select P.现价,P.收入项目id" & _
                " from 收费价目 P" & _
                " where P.收费细目id=[1] and (Sysdate Between p.执行日期 And p.终止日期 or Sysdate>=p.执行日期 And p.终止日期 Is Null)" & _
                GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtEdit(16).Text = Format(!现价 * IIf(mintUnit = 0, 1, dbl换算系数), mFMT.FM_零售价)
            For mintCount = 0 To Me.cbo收入项目.ListCount - 1
                If Me.cbo收入项目.ItemData(mintCount) = !收入项目id Then
                    Me.cbo收入项目.ListIndex = mintCount: Exit For
                End If
            Next
        End If
    End With

    If Val(mstr材料ID) <> 0 Then
        '--较证执行数据
        gstrSQL = "Select ID from 收费价目 where 收费细目id=[1] and nvl(变动原因,0)=0" & _
                GetPriceClassString("")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
        
        Do While Not rsTemp.EOF
                gstrSQL = "zl_材料收发记录_Adjust(" & Val(zlStr.nvl(rsTemp!Id)) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                rsTemp.MoveNext
        Loop
    End If
    
    '根据是否有发生，确定：材料属性、成本价格、零售价格可修改否
    gstrSQL = " Select nvl(Count(*),0) From 药品收发记录 Where 药品ID=[1] And rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
    
    With rsTemp
        If .Fields(0).Value > 0 Then
            Me.cbo价格属性.Enabled = False
            Me.txtEdit(15).Enabled = False
            Me.txtEdit(16).Enabled = False
'            Me.cbo收入项目.Enabled = False
        Else
            Me.cbo价格属性.Enabled = cbo价格属性.Enabled
            Me.txtEdit(15).Enabled = Me.txtEdit(15).Enabled  '成本价
            Me.txtEdit(16).Enabled = cbo价格属性.Enabled       '当前售价
'            Me.cbo收入项目.Enabled = cbo价格属性.Enabled
        End If
        SetCtlBackColor txtEdit(15)
        SetCtlBackColor txtEdit(16)
    End With
    
    If Me.chk跟踪.Value = 1 Then
'        Me.chk跟踪.Enabled = Me.cbo价格属性.Enabled
        chk跟踪.Tag = 1
    Else
        chk跟踪.Tag = 0
    End If
    
    If Val(mstr材料ID) <> 0 Then
        '如果存在未执行的价格,则不充许修改相关价格
        gstrSQL = "Select 现价 from 收费价目 where 收费细目id=[1] and nvl(变动原因,0)=0" & _
                GetPriceClassString("")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
        
        If rsTemp.EOF = False Then
            Me.cbo价格属性.Enabled = False
            'Me.txtEDIT(15).Enabled = False
            Me.txtEdit(16).Enabled = False
            Me.cbo收入项目.Enabled = False
        End If
    End If
    '根据是否有库存，确定：分批特性可修改否
    
    gstrSQL = "" & _
        "   Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
        "   Where A.药品ID=[1] And A.库房ID=B.部门ID And B.工作性质 In ('卫材库','物资库房', '虚拟库房')"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
        
    With rsTemp
        
        If .Fields(0).Value > 0 Then
            Me.chk库房.Enabled = False
        Else
            Me.chk库房.Enabled = True
        End If
    End With
    
    If Me.chk库房.Value = 1 Then
        gstrSQL = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                 " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '发料部门' Or B.工作性质 Like '%制剂室')"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr材料ID))
                
        With rsTemp
            If .Fields(0).Value > 0 Then
                Me.chk在用.Enabled = False
                If Me.chk库房.Enabled Then Me.chk库房.Enabled = IIf(chk在用.Value = 1, False, True)
            Else
                Me.chk在用.Enabled = True
            End If
        End With
    End If
    
    InitCardData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub 获取上次录入规格信息(ByVal lng诊疗ID As Long)
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:获取上次录入的规格信息
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim lng材料ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, dbl换算系数 As Double
    gstrSQL = "Select ID as 材料id" & _
              "  From 收费项目目录 A," & _
              "      (Select Max(a.建档时间) As 建档时间 From 收费项目目录 A, 材料特性 B Where a.Id = b.材料id And b.诊疗id = [1]) B " & _
              "  Where a.建档时间 = b.建档时间 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
    If rsTemp.EOF Then Exit Sub
    lng材料ID = Val(zlStr.nvl(rsTemp!材料ID))
    If lng材料ID = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
   '其他需读取卡片数据
    '----------数据装载-------------------------------------
    gstrSQL = "select I.编码 as 规格编码,I.名称,I.规格,I.产地 as 生产商,S.货源情况,S.材料来源, " & _
             "        I.计算单位,S.换算系数,S.包装单位,I.是否变价,S.指导批发价 as 指导批价,S.扣率 as 采购扣率,S.指导零售价 as 指导售价," & _
             "        S.指导差价率,S.加成率,S.差价让利比 as 差价让利,S.成本价 as 成本价格, " & _
             "        I.标识主码,I.标识子码,I.费用类型,I.服务对象,I.屏蔽费别, " & _
             "        S.库房分批,S.在用分批,S.最大效期 as 保存期,S.灭菌效期,S.无菌性材料,S.一次性材料,S.原材料,S.批准文号,S.注册商标," & _
             "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,S.跟踪在用,S.核算材料,s.注册证有效期,S.许可证号,S.许可证有效期,S.材质分类,S.存储条件,I.站点,S.增值税率,I.说明,S.是否植入耗材 " & _
             "  from 收费项目目录 I,材料特性 S " & _
             "  where I.ID=S.材料ID and I.id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
             
    
    Dim strFieldsName As String
    
    With rsTemp
        If .RecordCount > 0 Then
            txt批准文号.Text = zlStr.nvl(!批准文号)
            txt注册商标.Text = zlStr.nvl(!注册商标)
                  
            For i = 0 To txtEdit.UBound
                strFieldsName = txtEdit(i).Tag
                Select Case strFieldsName
                Case "灭菌效期", "保存期"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), "######")
                Case "指导批价", "成本价格"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!换算系数, 1)), mFMT.FM_成本价)
                Case "指导售价"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!换算系数, 1)), mFMT.FM_零售价)
                Case "采购扣率", "指导差价率", "差价让利", "加成率"
                        txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBCJL)
                Case "拼音", "五笔", "结算价", "当前售价"
                Case "许可证号"
                    txtEdit(i).Text = zlStr.nvl(!许可证号)
                Case "规格", "规格编码"
                    txtEdit(i).Text = ""
                Case "增值税率"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBJCL)
                Case Else
                    '“商品名,拼音简码,五笔简码”单独处理
                    If InStr("商品名;拼音简码;五笔简码", strFieldsName) = 0 Then txtEdit(i).Text = zlStr.nvl(.Fields(strFieldsName))
                End Select
            Next
            
            If IsNull(!许可证有效期) Then
                dtp许可证效期.Value = ""
            Else
                dtp许可证效期.Value = Format(!许可证有效期, "yyyy-mm-dd")
            End If
            If IsNull(!注册证有效期) Then
                dtp注册证有效期.Value = ""
            Else
                dtp注册证有效期.Value = Format(!注册证有效期, "yyyy-mm-dd")
            End If
            
            '计算单位
            For mintCount = 0 To Me.cbo单位(0).ListCount - 1
                If cbo单位(0).List(mintCount) = zlStr.nvl(!计算单位) Then
                    cbo单位(0).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo单位(0).ListIndex < 0 Then
                If zlStr.nvl(!计算单位) <> "" Then
                    cbo单位(0).AddItem zlStr.nvl(!计算单位)
                    cbo单位(0).ListIndex = cbo单位(0).NewIndex
                End If
            End If
            '包装单位
            For mintCount = 0 To Me.cbo单位(1).ListCount - 1
                If cbo单位(1).List(mintCount) = zlStr.nvl(!包装单位) Then
                    cbo单位(1).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo单位(1).ListIndex < 0 Then
                If zlStr.nvl(!包装单位) <> "" Then
                    cbo单位(1).AddItem zlStr.nvl(!计算单位)
                    cbo单位(1).ListIndex = cbo单位(1).NewIndex
                End If
            End If
            dbl换算系数 = zlStr.nvl(!换算系数, 1)
            '--材料来源
            For mintCount = 0 To Me.cbo货源.ListCount - 1
                If Mid(Me.cbo货源.List(mintCount), InStr(1, Me.cbo货源.List(mintCount), "-") + 1) = zlStr.nvl(!货源情况) Then
                    Me.cbo货源.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo材料来源.ListCount - 1
                If Mid(Me.cbo材料来源.List(mintCount), InStr(1, Me.cbo材料来源.List(mintCount), "-") + 1) = zlStr.nvl(!材料来源) Then
                    Me.cbo材料来源.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo材质分类.ListCount - 1
                If Mid(Me.cbo材质分类.List(mintCount), InStr(1, Me.cbo材质分类.List(mintCount), "-") + 1) = zlStr.nvl(!材质分类) Then
                    Me.cbo材质分类.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--材料来源
            For mintCount = 0 To Me.cbo存储条件.ListCount - 1
                If Mid(Me.cbo存储条件.List(mintCount), InStr(1, Me.cbo存储条件.List(mintCount), "-") + 1) = zlStr.nvl(!存储条件) Then
                    Me.cbo存储条件.ListIndex = mintCount: Exit For
                End If
            Next
            
            
          '时价
            For mintCount = 0 To Me.cbo价格属性.ListCount - 1
                If cbo价格属性.ItemData(mintCount) = zlStr.nvl(!是否变价, 0) Then
                    cbo价格属性.ListIndex = mintCount
                    Exit For
                End If
            Next
            
            lbl批价单位(0).Caption = "元/" & IIf(mintUnit = 0, zlStr.nvl(!计算单位), zlStr.nvl(!包装单位))
            lbl批价单位(1).Caption = "元/" & IIf(mintUnit = 0, zlStr.nvl(!计算单位), zlStr.nvl(!包装单位))
            
            cbo价格属性.ListIndex = IIf(IsNull(!是否变价), 0, !是否变价)
            
            Chk一次性材料.Value = zlStr.nvl(!一次性材料, 0)
            chk无菌性材料.Value = zlStr.nvl(!无菌性材料, 0)
            chk原料.Value = zlStr.nvl(!原材料, 0)
            chk植入耗材.Value = zlStr.nvl(!是否植入耗材, 0)
            
            For mintCount = 0 To Me.cbo费用类型.ListCount - 1
                If Mid(Me.cbo费用类型.List(mintCount), InStr(1, Me.cbo费用类型.List(mintCount), "-") + 1) = IIf(IsNull(!费用类型), "", !费用类型) Then
                    Me.cbo费用类型.ListIndex = mintCount: Exit For
                End If
            Next
            
            If InStr(1, mstrPrivs, ";服务对象;") <> 0 Then
                Me.cbo服务对象.ListIndex = IIf(IsNull(!服务对象), 0, !服务对象)
            Else
                cbo服务对象.Enabled = False
            End If
            
            Me.chk屏蔽费别.Value = IIf(IsNull(!屏蔽费别), 0, !屏蔽费别)
                 
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "注：该规格于" & Format(!建档时间, "YYYY年MM月DD日") & "建立，" & Format(!撤档时间, "YYYY年MM月DD日") & "停用"
            Else
                Me.lblFound.Caption = ""
            End If
            
            Me.chk库房.Value = zlStr.nvl(!库房分批, 0)
            Me.chk在用.Value = zlStr.nvl(!在用分批, 0)
            Me.chk跟踪.Value = zlStr.nvl(!跟踪在用, 0)
            Me.chk核算材料.Value = Val(zlStr.nvl(!核算材料))
             
            Me.chk在用.Tag = Me.chk在用.Value
            
            If Me.chk库房.Value = 0 Then
                Me.chk在用.Enabled = False: Me.chk在用.Value = 0
            Else
                Me.chk在用.Enabled = True
                Me.chk在用.Value = Me.chk在用.Tag
            End If
            '读取站点信息
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!站点) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
        End If
        
   End With
         
    '提取显示当前售价
    If Me.cbo价格属性.ListIndex <> 0 Then
        '时价材料，取库存金额/库存数量做为其价格，无库存时取价表定价
        gstrSQL = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                " from 收费价目 P," & _
                "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                "      From 药品库存 Where 药品ID=[1]) K" & _
                " where P.收费细目id=[1] and (Sysdate Between p.执行日期 And p.终止日期 or Sysdate>=p.执行日期 And p.终止日期 Is Null)" & _
                GetPriceClassString("P")
    Else
        '非时价材料调价，取其价格记录中的价格
        gstrSQL = "select P.现价,P.收入项目id" & _
                " from 收费价目 P" & _
                " where P.收费细目id=[1] and (Sysdate Between p.执行日期 And p.终止日期 or Sysdate>=p.执行日期 And p.终止日期 Is Null)" & _
                GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtEdit(16).Text = Format(!现价 * IIf(mintUnit = 0, 1, dbl换算系数), mFMT.FM_零售价)
            For mintCount = 0 To Me.cbo收入项目.ListCount - 1
                If Me.cbo收入项目.ItemData(mintCount) = !收入项目id Then
                    Me.cbo收入项目.ListIndex = mintCount: Exit For
                End If
            Next
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetPopedom()
    Dim intI As Long
    Dim blnStuffModify As Boolean
    
    chkCode.Visible = zlStr.IsHavePrivs(mstrPrivs, "设置条码管理")
    
    If mintEditType = g新增 Or mintEditType = g修改 Then
   
        If mintEditType = g修改 Then
            '看是否允许修改相关卫材档案信息
            blnStuffModify = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0
            For intI = 0 To txtEdit.UBound
                 txtEdit(intI).Enabled = blnStuffModify
            Next
            
            txt批准文号.Enabled = blnStuffModify
            txt注册商标.Enabled = blnStuffModify
            txt注册证号.Enabled = blnStuffModify
            
            SetCtlBackColor txt批准文号
            SetCtlBackColor txt注册商标
            
            cbo单位(0).Enabled = blnStuffModify
            cbo单位(1).Enabled = blnStuffModify
            
            chk原料.Enabled = blnStuffModify
            chk屏蔽费别.Enabled = blnStuffModify
            cbo材料来源.Enabled = blnStuffModify
            chk分零使用.Enabled = blnStuffModify
            
            cmd产地.Enabled = blnStuffModify
            cbo货源.Enabled = blnStuffModify
            cbo服务对象.Enabled = blnStuffModify
            cbo材料来源.Enabled = blnStuffModify
            cbo存储条件.Enabled = blnStuffModify
            dtp许可证效期.Enabled = blnStuffModify
            dtp注册证有效期.Enabled = blnStuffModify
            chk库房.Enabled = blnStuffModify
            chk在用.Enabled = blnStuffModify
            
            chk跟踪.Enabled = blnStuffModify
            
            chk核算材料.Enabled = blnStuffModify
            cbo材质分类.Enabled = blnStuffModify
            chk跟踪病人.Enabled = blnStuffModify
            
            chk植入耗材.Enabled = blnStuffModify
            Chk一次性材料.Enabled = blnStuffModify
            chkCostly.Enabled = blnStuffModify
            chkCode.Enabled = blnStuffModify
            chk无菌性材料.Enabled = blnStuffModify
            fra分批核算.Enabled = blnStuffModify
            txt备选码.Enabled = blnStuffModify
            cmbStationNo.Enabled = blnStuffModify
            chkInstrument.Enabled = blnStuffModify
            SetCtlBackColor txt备选码
            
        Else
            txt批准文号.Enabled = True
            txt注册商标.Enabled = True
            txt备选码.Enabled = True
            SetCtlBackColor txt备选码
        End If
        
        Me.txtEdit(9).Enabled = InStr(1, mstrPrivs, ";管理扣率;") <> 0     '扣率
        Me.txtEdit(12).Enabled = InStr(1, mstrPrivs, ";指导价格管理;") <> 0       '指导差价率
        Me.txtEdit(8).Enabled = Me.txtEdit(12).Enabled                          '指导批价
        Me.txtEdit(11).Enabled = Me.txtEdit(12).Enabled                          '指导售价
        Me.txtEdit(13).Enabled = Me.txtEdit(12).Enabled                          '加成率
        Me.txtEdit(14).Enabled = Me.txtEdit(12).Enabled
        
        Me.cbo价格属性.Enabled = InStr(1, mstrPrivs, ";售价管理;") <> 0
        Me.txtEdit(15).Enabled = InStr(1, mstrPrivs, ";成本价管理;") <> 0                  '成本价格
        Me.txtEdit(16).Enabled = Me.cbo价格属性.Enabled                 '当前售价
        Me.cbo收入项目.Enabled = InStr(1, mstrPrivs, ";调整收入项目;") <> 0
        Me.cbo费用类型.Enabled = InStr(1, mstrPrivs, ";医保用料目录;") <> 0
        Me.cbo服务对象.Enabled = InStr(1, mstrPrivs, ";服务对象;") <> 0
        
        For intI = 0 To txtEdit.UBound
            SetCtlBackColor txtEdit(intI)
        Next
    
        Exit Sub
    Else
        txt病案费目.Enabled = False
        cmd病案.Enabled = False
    End If
    For intI = 0 To txtEdit.UBound
        txtEdit(intI).Enabled = False
        SetCtlBackColor txtEdit(intI)
    Next
    
    txt批准文号.Enabled = False
    txt注册商标.Enabled = False
    txt注册证号.Enabled = False
    txt备选码.Enabled = False
    SetCtlBackColor txt备选码
    SetCtlBackColor txt批准文号
    SetCtlBackColor txt注册商标
    SetCtlBackColor txt注册证号
    
    cbo单位(0).Enabled = False
    cbo单位(1).Enabled = False
    
    chk原料.Enabled = False
    chk屏蔽费别.Enabled = False
    cbo材料来源.Enabled = False
    chk分零使用.Enabled = False
    
    cmd产地.Enabled = False
    cbo货源.Enabled = False
    cbo价格属性.Enabled = False
    cbo费用类型.Enabled = False
    cbo收入项目.Enabled = False
    cbo服务对象.Enabled = False
    cbo材质分类.Enabled = False
    cbo存储条件.Enabled = False
    dtp许可证效期.Enabled = False
    dtp注册证有效期.Enabled = False
    chk库房.Enabled = False
    chk在用.Enabled = False
    chk跟踪.Enabled = False
    chk核算材料.Enabled = False
    chk跟踪病人.Enabled = False
    Chk一次性材料.Enabled = False
    chk植入耗材.Enabled = False
    chkCostly.Enabled = False
    chkCode.Enabled = False
    chkInstrument.Enabled = False
    chk无菌性材料.Enabled = False
    fra分批核算.Enabled = False
    cmbStationNo.Enabled = False
    cmdOK.Visible = False
    cmdCancel.Caption = "关闭(&C)"
End Sub

Private Sub dtp许可证效期_Change()
    mblnChange = True
End Sub

Private Sub dtp许可证效期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub dtp注册证有效期_Change()
    mblnChange = True
End Sub

Private Sub dtp注册证有效期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    If mblnLoad = True Then Exit Sub
    If mblnFrist = False Then Exit Sub
    mblnFrist = True
    
    '初始站点
    cmbStationNo.Visible = gSystem_Para.bln存在站点
    lblStationNo.Visible = cmbStationNo.Visible
    gstrSQL = "Select Count(1) 器械包单件 From all_Tab_Columns Where Table_Name = '材料特性' And Column_Name = '器械包卫材单件'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否共享安装物资系统")
    If rsTemp!器械包单件 = 0 Then '没有共享安装物资系统
        chkInstrument.Visible = False
    End If
    
    gstrSQL = "Select Count(1) 物资系统  From zlSystems Where 编号 = 400 And 共享号 = 100"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否共享安装物资系统")
    mblnInStrument = False
    If rsTemp!物资系统 > 0 Then
        mblnInStrument = True
    Else
        chkInstrument.Visible = False
    End If
    
    mintSet分批 = Val(zlDatabase.GetPara("卫材分批属性自动设置", glngSys, mlngModule, 0))
    '----------依赖关系判断-------------------------------------
    If GetDepend = False Then
        Unload Me
        Exit Sub
    End If
    
    '----------程序权限控制-------------------------------------
    Call SetPopedom
    
    '----------初始卡片数据-------------------------------------
    Call InitCardData
    
    '----------默认第一选项卡-----------------------------------
    If mintEditType = g修改 Then
      If InStr(1, mstrPrivs, ";卫材品名管理;") <> 0 Then
          Me.stbSpec.Tab = 0
      Else
          Me.stbSpec.Tab = 1
      End If
    Else
      Me.stbSpec.Tab = 0
    End If
    mblnLoad = True
     
    If InStr(1, mstrPrivs, ";调整跟踪在用;") = 0 Then
       chk跟踪.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub txt病案费目_GotFocus()
    txt病案费目.SelStart = 0
    txt病案费目.SelLength = Len(txt病案费目)
    txt病案费目.SetFocus
End Sub

Private Sub txt病案费目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Then
        txt病案费目.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Form_Load()
    Dim aryTemp() As String
    Dim strSql As String
    Dim rsrecord As ADODB.Recordset
    
    mblnFrist = True
    On Error GoTo ErrHandle
            
    If mintEditType <> g新增 Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    strSql = "select 编号,名称 from zlnodelist"
    Set rsrecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
    With cmbStationNo
        .AddItem ""
        Do While Not rsrecord.EOF
            .AddItem rsrecord!编号 & "-" & rsrecord!名称
            rsrecord.MoveNext
        Loop
    End With
    
    '----------------装入可选的基础数据----------------------
    With Me.cbo价格属性
        .Clear
        aryTemp = Split("0-定价;1-时价", ";")
        For mintCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(mintCount): .ItemData(.NewIndex) = mintCount
        Next
        .ListIndex = 0
    End With
    
    With Me.cbo服务对象
        aryTemp = Split("0-不应用于病人;1-门诊;2-住院;3-门诊和住院", ";")
        For mintCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(mintCount): .ItemData(.NewIndex) = mintCount
        Next
        If InStr(1, mstrPrivs, ";服务对象;") <> 0 Then
            .ListIndex = 3
        Else
            .ListIndex = 0
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnLoad = False
End Sub

Private Sub txtEDIT_Change(Index As Integer)
    Dim strTag As String
    strTag = txtEdit(Index).Tag
    
    Select Case strTag
        Case "采购扣率"
            '计算结算价=指导批价*扣率/100
            Me.txtEdit(10).Text = Format(Val(Me.txtEdit(8).Text) * Val(Me.txtEdit(Index).Text) / 100, mFMT.FM_成本价)
        Case "指导批价"
            '计算结算价=指导批价*扣率/100
            Me.txtEdit(10).Text = Format(Val(Me.txtEdit(Index).Text) * Val(Me.txtEdit(9).Text) / 100, mFMT.FM_成本价)
        Case "标识主码", "标识子码"
                txtEdit(Index).Text = UCase(txtEdit(Index).Text)
                txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
        Case "商品名"
            '拼音和五笔
            Me.txtEdit(19).Text = zlStr.GetCodeByORCL(Me.txtEdit(Index).Text, 0, Me.txtEdit(19).MaxLength)
            Me.txtEdit(20).Text = zlStr.GetCodeByORCL(Me.txtEdit(Index).Text, 1, Me.txtEdit(20).MaxLength)
        Case "加成率"
            If Val(txtEdit(Index).Text) > 9900 Then txtEdit(Index).Text = 9900
            If Val(txtEdit(Index).Text) < 0 Then txtEdit(Index).Text = 0
    End Select
End Sub

Private Sub txtEDIT_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strTag As String
    Dim intDigit As Integer
    Dim strKey As String
    
    If Index = 0 Then
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        End Select
        KeyAscii = 0
        Exit Sub
    End If
    
    strKey = txtEdit(Index).Text
    strTag = txtEdit(Index).Tag
    Select Case strTag
        Case "换算系数", "灭菌效期", "保存期"
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
        Case "指导批价", "采购扣率", "结算价", "指导售价", "指导差价率", "加成率", "差价让利", "成本价格", "当前售价", "增值税率"
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m金额式
            
            If strTag = "指导批价" Or strTag = "结算价" Or strTag = "成本价格" Or strTag = "指导售价" Or strTag = "当前售价" Then
                Select Case strTag
                    Case "指导批价", "结算价", "成本价格"
                        intDigit = Len(Mid(mFMT.FM_成本价, InStr(1, mFMT.FM_成本价, ".") + 1))
                    Case "指导售价", "当前售价"
                        intDigit = Len(Mid(mFMT.FM_成本价, InStr(1, mFMT.FM_零售价, ".") + 1))
                End Select
                
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If txtEdit(Index).SelLength = Len(strKey) Then Exit Sub
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Case Else
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    End Select
    
    If strTag = "生产商" Then     '产地
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        If KeyAscii <> vbKeyReturn Then Exit Sub
        If txtEdit(Index).Text <> "" Then
            Call Sel产地(txtEdit(Index).Text)
        End If
        Exit Sub
    End If
    
    If strTag = "规格" Then
        If InStr("^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Exit Sub
    End If
    
    If strTag = "商品名" Then
        If InStr("^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Exit Sub
    End If
    
    If strTag = "五笔" Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Sel产地(ByVal strKey As String)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:选择产地
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim vRect  As RECT, lngH As Long
    Dim objTxt As Object
    Dim blnCancel As Boolean
    
    Dim strTemp As String
    Set objTxt = txtEdit(GetTxtIdx("生产商"))
    
    strTemp = strKey
    
    strTemp = GetMatchingSting(strTemp)
    If strKey = "" Then
        gstrSQL = "Select Rownum as ID, 编码,名称,简码 From 材料生产商 Order By 编码 "
    Else
        gstrSQL = "Select Rownum as ID,编码,名称,简码 From 材料生产商 where 编码 Like [1]  Or 名称 Like [1] Or 简码 Like [1] Order By 编码 "
    End If
    
    vRect = zlControl.GetControlRect(objTxt.hwnd)
    lngH = objTxt.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "产地选择器", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strTemp)
   
   '     frmParent=显示的父窗体
   '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
   '     bytStyle=选择器风格
   '       为0时:列表风格:ID,…
   '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
   '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
   '     strTitle=选择器功能命名,也用于个性化区分
   '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
   '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
   '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
   '             bytStyle=1时,可以是编码或名称
   '     strNote=选择器的说明文字
   '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
   '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
   '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
   '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
   '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
   '     blnSearch=是否显示行号,并可以输入行号定位
    If blnCancel Then
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    If rsTemp Is Nothing Then
        If mstr产地 <> strKey And strKey <> "" Then
                If Asc(strKey) > 0 Then
                    MsgBox "没有找到匹配的生产商，请重新输入！", vbInformation, gstrSysName
                    If objTxt.Enabled Then objTxt.SetFocus
                    mstr产地 = ""
                    Exit Sub
                End If
        
                If MsgBox("没有找到相关的生产商，增加该生产商吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If objTxt.Enabled Then objTxt.SetFocus
                    mstr产地 = ""
                Else
                    If zlSureManufacturer = False Then
                        MsgBox "生产商的编码超长，无法自动增加。" & vbCrLf & "请输入或选择现有的材料生产商！", vbInformation, gstrSysName
                        objTxt.Text = "": mstr产地 = "": Exit Sub
                    Else
                        Dim str编码 As String, str名称 As String
                        str名称 = strKey
                        If AutoAdd生产商(str编码, str名称, Me.Caption) = False Then
                            mstr产地 = ""
                            If objTxt.Enabled Then objTxt.SetFocus
                            Exit Sub
                        Else
                            mstr产地 = strKey
                        End If
                        Call OS.PressKey(vbKeyTab): Exit Sub
                    End If
                End If
        End If
        Exit Sub
    End If
    objTxt.Text = zlStr.nvl(rsTemp!名称)
    If objTxt.Enabled Then objTxt.SetFocus
    Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtEDIT_LostFocus(Index As Integer)
    Dim cur价格 As Double
    Dim strTag   As String
    Dim dbl加成率 As Double
    Dim dbl差价率 As Double
    strTag = txtEdit(Index).Tag
    Select Case strTag
        Case "指导批价", "结算价"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_成本价)
        Case "指导售价", "当前售价"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_零售价)
            
            If strTag = "当前售价" Then
                If Val(txtEdit(16).Text) <> 0 Then
                    txtEdit(11).Text = txtEdit(16).Text
                End If
                '满足这些条件才计算加成率  15-成本价,11-指导售价,16-现售价,14-差价让利,13-加成率,12-指导差价率
                If Val(Trim(txtEdit(15).Text)) > 0 And Val(Trim(txtEdit(11).Text)) > 0 And Val(Trim(txtEdit(16).Text)) > 0 And Val(Trim(txtEdit(16).Text)) <= Val(Trim(txtEdit(11).Text)) And Val(Trim(txtEdit(14).Text)) / 100 <> 0 Then
                    If Val(Trim(txtEdit(14).Text)) / 100 = 1 Then
                        dbl加成率 = Val(Trim(txtEdit(16).Text)) / Val(Trim(txtEdit(15).Text)) - 1
                    Else
                        dbl加成率 = ((Val(Trim(txtEdit(16).Text)) - Val(Trim(txtEdit(11).Text)) * (1 - Val(Trim(txtEdit(14).Text)))) / Val(Trim(txtEdit(14).Text))) / Val(Trim(txtEdit(15).Text)) - 1
                    End If
                    
                    If dbl加成率 < 0 Then Exit Sub
                    
                    dbl加成率 = dbl加成率 * 100
                    
                    txtEdit(13).Text = Format(dbl加成率, "0.00")
                    
                    '通过加成率计算指导差价率
                    dbl差价率 = dbl加成率
                    Call Calc(dbl差价率, False)
                    
                    txtEdit(12).Text = Format(dbl差价率, "0.00000")
                End If
            End If
        Case "差价让利"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBCJL)
        Case "成本价格"
            Dim dblSalePrice As Double
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_成本价)
            If Val(txtEdit(16).Text) = 0 And Val(txtEdit(Index).Text) <> 0 Then
                '当前售价为零时,需重算售价
                '成本价*(1+加成率)
                dblSalePrice = Val(txtEdit(Index).Text) * (1 + Val(Me.txtEdit(13).Text) / 100)  '
                '成本价*（1+加成率）+(指导售价 -（成本价*（1+加成率））)*(1-差价让利率)
                dblSalePrice = dblSalePrice + (Val(Me.txtEdit(11).Text) - dblSalePrice) * (1 - Val(Me.txtEdit(14)) / 100)
                
                If Val(txtEdit(11).Text) <> 0 Then
                    If dblSalePrice > Val(Me.txtEdit(11).Text) Then
                        '大于指导售价,则按指导价算
                        dblSalePrice = Val(Me.txtEdit(11).Text)
                    End If
                End If
                Me.txtEdit(16).Text = Format(dblSalePrice, mFMT.FM_零售价)
            End If
            
            If Val(txtEdit(15).Text) <> 0 And Val(txtEdit(8).Text) = 0 Then
                txtEdit(8).Text = txtEdit(15).Text
            End If
        Case "加成率"
            '重新计算指导差价率和加成率
            cur价格 = Val(txtEdit(13).Text)
            Call Calc(cur价格, False)
            
            '加成率
            Me.txtEdit(13).Text = Format(txtEdit(13).Text, GFM_VBJCL)
            '指导差价率
            Me.txtEdit(12).Text = Format(cur价格, GFM_VBCJL)
        Case "指导差价率"
            '重新计算指导差价率和加成率
            
            cur价格 = Val(txtEdit(12).Text) '指导差价率
            
            If cur价格 < 100 Then
                Call Calc(cur价格, True)
                '指导差价率
                Me.txtEdit(Index).Text = Format(txtEdit(Index).Text, GFM_VBCJL)
                
                '加成率
                Me.txtEdit(13).Text = Format(cur价格, GFM_VBJCL)
            Else
                '不允许出现指导差价率大于等于100的情况，因此需要从加成率反算回来
                cur价格 = Val(txtEdit(13).Text)
                Call Calc(cur价格, False)
                Me.txtEdit(Index).Text = Format(cur价格, GFM_VBCJL)
            End If
        Case "采购扣率"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBKL)
        Case "增值税率"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBJCL)
    End Select
    
    '关闭输入法
    ImeLanguage False
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
    Dim strTag As String
    strTag = txtEdit(Index).Tag
    zlControl.TxtSelAll txtEdit(Index)
    OS.OpenIme True
    Select Case strTag
        Case "名称", "规格", "生产商", "商品名"
            '打开输入法
            ImeLanguage True
        Case "标识主码"
            OS.OpenIme False
    End Select
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            OS.PressKey vbKeyTab
        End If
End Sub

Private Sub cbo单位_GotFocus(Index As Integer)
    Me.cbo单位(Index).SelStart = 0: Me.cbo单位(Index).SelLength = 100
    ImeLanguage True
End Sub

Private Sub cbo单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
           Exit Sub
        Case Else
            zlControl.TxtCheckKeyPress cbo单位(Index), KeyAscii, m文本式
        End Select
        Exit Sub
    End If
End Sub

Private Sub cbo单位_LostFocus(Index As Integer)
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    ImeLanguage False
    strTmp = cbo单位(Index).Text
    blnAdd = True
    For i = 0 To cbo单位(Index).ListCount - 1
        If cbo单位(Index).List(i) = Trim(strTmp) Then
            blnAdd = False
            Exit For
        End If
    Next
    If blnAdd And strTmp <> "" Then
        cbo单位(Index).AddItem strTmp
    End If
    If Index <> 0 Then Exit Sub
    Me.lbl批价单位(0).Caption = "元/" & cbo单位(Index).Text
    Me.lbl批价单位(1).Caption = "元/" & cbo单位(Index).Text

End Sub


Private Sub cbo单位_Change(Index As Integer)
    If mintUnit = 0 Then
        If Index = 1 Then Exit Sub
    Else
        If Index = 0 Then Exit Sub
    End If
    
    Me.lbl批价单位(0).Caption = "元/" & cbo单位(Index).Text
    Me.lbl批价单位(1).Caption = "元/" & cbo单位(Index).Text
End Sub


Private Sub stbSpec_Click(PreviousTab As Integer)
    If Me.msf产地.Visible Then stbSpec.Tab = 0: Me.msf产地.SetFocus: Exit Sub
    
    Select Case stbSpec.Tab
    Case 0
        If Me.txtEdit(0).Enabled Then Me.txtEdit(0).SetFocus
    Case 1
        If Me.txtEdit(8).Enabled Then Me.txtEdit(8).SetFocus
        If Me.cbo价格属性.Enabled Then Me.cbo价格属性.SetFocus
    End Select
End Sub

Private Function zlSureManufacturer() As Boolean
    '-------------------------------------------------------------
    '功能：判断是否可继续增加生产商（生产商编码字段宽度为:10）
    '-------------------------------------------------------------
    Dim strTemp  As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    zlSureManufacturer = False
    gstrSQL = "Select Max(编码) 编码 From 材料生产商"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        If .EOF Then zlSureManufacturer = True: Exit Function
        If IsNull(!编码) Then zlSureManufacturer = True: Exit Function
        
        '如果超长则退出
        strTemp = .Fields(0).Value
        mintCount = Len(strTemp)
        strTemp = strTemp + 1
        If Len(strTemp) > 10 Then Exit Function
        If mintCount >= Len(strTemp) Then
            strTemp = String(mintCount - Len(strTemp), "0") & strTemp
        End If
    End With
    zlSureManufacturer = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Calc(dbl价格 As Double, Optional ByVal bln差价率 As Boolean = True)
    '如果传入的是差价率，计算加成率并返回；否则计算差价率并返回
    '加成率与差价率间，存在下列对应关系
    '加成率=1/(1-差价率)-1
    '差价率=1-1/(1+加成率)
    dbl价格 = dbl价格 / 100
    If bln差价率 Then
        dbl价格 = 1 / (1 - dbl价格) - 1
    Else
        dbl价格 = 1 - 1 / (1 + dbl价格)
    End If
    dbl价格 = dbl价格 * 100
End Sub
  
Private Sub txt备选码_Change()
    mblnChange = True
End Sub

Private Sub txt备选码_GotFocus()
    Call OS.OpenIme(False)
    zlControl.TxtSelAll txt备选码
End Sub

Private Sub txt备选码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmbStationNo.Visible = True Then
        OS.PressKey vbKeyTab
    Else
        stbSpec.Tab = 1
        If cbo价格属性.Enabled Then cbo价格属性.SetFocus
    End If
End Sub

Private Sub txt备选码_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt备选码, KeyAscii, m文本式
End Sub

Private Sub txt批准文号_GotFocus()
    Me.txt批准文号.SelStart = 0: Me.txt批准文号.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt批准文号_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt批准文号_LostFocus()
    Call OS.OpenIme(False)
End Sub
Private Sub txt注册商标_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub


Private Sub txt注册商标_LostFocus()
    Call OS.OpenIme(False)
End Sub
Private Sub txt注册商标_GotFocus()
    Me.txt注册商标.SelStart = 0: Me.txt注册商标.SelLength = 100
    Call OS.OpenIme(True)
End Sub
Private Sub SaveReg()
    '功能:保存相关的注册信息
    Dim strReg As String
    Call zlDatabase.SetPara("上次指导差价率", Val(Me.txtEdit(12)), glngSys, mlngModule)
    Call zlDatabase.SetPara("上次加成率", Val(Me.txtEdit(13)), glngSys, mlngModule)
End Sub
Private Sub LoadReg()
    '功能:加载注册信息值
    Dim strReg As String
    Dim blnHavePriv As Boolean
    blnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置") And zlStr.IsHavePrivs(mstrPrivs, "指导价格管理")
    
    strReg = zlDatabase.GetPara("上次指导差价率", glngSys, mlngModule)
    txtEdit(12).Text = Format(IIf(Val(strReg) = 0, 13.0435, Val(strReg)), GFM_VBCJL)
    strReg = zlDatabase.GetPara("上次加成率", glngSys, mlngModule)
    txtEdit(13).Text = Format(IIf(Val(strReg) = 0, 15, Val(strReg)), GFM_VBCJL)
End Sub


Private Sub txt注册证号_Change()
    mblnChange = True
End Sub

Private Sub txt注册证号_GotFocus()
    zlControl.TxtSelAll txt注册证号
    OS.OpenIme False
End Sub

Private Sub txt注册证号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt注册证号_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt注册证号, KeyAscii, m文本式
End Sub

