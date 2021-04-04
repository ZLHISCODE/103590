VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm发药参数设置 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Frm发药参数设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7380
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   3600
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "基础(&1)"
      TabPicture(0)   =   "Frm发药参数设置.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "辅助(&2)"
      TabPicture(1)   =   "Frm发药参数设置.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frm发药参数设置"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "打印(&3)"
      TabPicture(2)   =   "Frm发药参数设置.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbo中药处方格式"
      Tab(2).Control(1)=   "cbo中药配药格式"
      Tab(2).Control(2)=   "cbo西药处方格式"
      Tab(2).Control(3)=   "cbo西药配药格式"
      Tab(2).Control(4)=   "chkPreview"
      Tab(2).Control(5)=   "cbo药品标签"
      Tab(2).Control(6)=   "cbo配药后"
      Tab(2).Control(7)=   "Cbo发药后"
      Tab(2).Control(8)=   "Fra打印"
      Tab(2).Control(9)=   "Fra刷新"
      Tab(2).Control(10)=   "lbl中药处方格式"
      Tab(2).Control(11)=   "lbl中药配药格式"
      Tab(2).Control(12)=   "lbl西药处方格式"
      Tab(2).Control(13)=   "lbl西药配药格式"
      Tab(2).Control(14)=   "lbl药品标签"
      Tab(2).Control(15)=   "Lbl发药"
      Tab(2).Control(16)=   "lbl配药"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "票据(&4)"
      TabPicture(3)   =   "Frm发药参数设置.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl票据"
      Tab(3).Control(1)=   "cmd打印设置"
      Tab(3).Control(2)=   "cbo票据设置"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "来源科室(&5)"
      TabPicture(4)   =   "Frm发药参数设置.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lvw来源科室"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "处方类型(&6)"
      TabPicture(5)   =   "Frm发药参数设置.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraSetColor"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "排队叫号(&7)"
      TabPicture(6)   =   "Frm发药参数设置.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Fra语音设备设置"
      Tab(6).Control(1)=   "frm显示设备设置"
      Tab(6).Control(2)=   "chk启用排队叫号"
      Tab(6).Control(3)=   "chkUseDisplay"
      Tab(6).ControlCount=   4
      Begin VB.ComboBox cbo中药处方格式 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   480
         Width           =   2280
      End
      Begin VB.ComboBox cbo中药配药格式 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   480
         Width           =   2280
      End
      Begin VB.ComboBox cbo西药处方格式 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   840
         Width           =   2280
      End
      Begin VB.ComboBox cbo西药配药格式 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chkPreview 
         Caption         =   "打印处方签时先预览再打印"
         Height          =   195
         Left            =   -70920
         TabIndex        =   114
         Top             =   1608
         Width           =   2640
      End
      Begin VB.Frame Frm发药参数设置 
         Caption         =   "  选项 "
         Height          =   4965
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   6675
         Begin VB.ComboBox cbo回车方式 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   3120
            Width           =   3360
         End
         Begin VB.CheckBox chkDispensing 
            Caption         =   "操作“呼叫”功能的同时通知药品自动化设备“准备发药”"
            Height          =   225
            Left            =   120
            TabIndex        =   109
            Top             =   3960
            Width           =   5460
         End
         Begin VB.CheckBox chk扫描后呼叫 
            Caption         =   "待发药处方扫描后自动呼叫"
            Height          =   225
            Left            =   120
            TabIndex        =   108
            Top             =   3600
            Width           =   2460
         End
         Begin VB.ComboBox cbo金额显示 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   240
            Width           =   2160
         End
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "设备配置(&S)"
            Height          =   350
            Left            =   2160
            TabIndex        =   98
            Top             =   4320
            Width           =   1500
         End
         Begin VB.ListBox lst卡类型 
            Appearance      =   0  'Flat
            Columns         =   1
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   2280
            Style           =   1  'Checkbox
            TabIndex        =   95
            Top             =   2280
            Width           =   2760
         End
         Begin VB.Frame fraline1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   1560
            TabIndex        =   86
            Top             =   2040
            Width           =   650
         End
         Begin VB.ComboBox cbo药品名称显示 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   240
            Width           =   2160
         End
         Begin VB.CheckBox chk大小单位 
            Caption         =   "用两种单位显示药品数量"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   91
            Top             =   840
            Width           =   2400
         End
         Begin VB.CheckBox Chk显示付数栏 
            Caption         =   "显示付数栏"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   1200
         End
         Begin VB.CheckBox chk配药扫描 
            Caption         =   "配药模式启用扫描器（两次扫描确认）"
            Height          =   225
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   4140
         End
         Begin VB.CheckBox chkOverTime 
            Caption         =   "单独显示"
            Height          =   225
            Left            =   120
            TabIndex        =   88
            Top             =   1800
            Width           =   1020
         End
         Begin VB.TextBox txtOverTime 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1665
            TabIndex        =   87
            Text            =   "1440"
            Top             =   1800
            Width           =   460
         End
         Begin VB.CheckBox chk发药刷卡 
            Caption         =   "发药模式两次刷卡发药"
            Height          =   225
            Left            =   120
            TabIndex        =   85
            Top             =   2280
            Width           =   2100
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "查找时系统自动回车方式"
            Height          =   180
            Left            =   120
            TabIndex        =   126
            Top             =   3180
            Width           =   1980
         End
         Begin VB.Label lbl金额显示 
            AutoSize        =   -1  'True
            Caption         =   "金额显示"
            Height          =   180
            Left            =   3600
            TabIndex        =   100
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "智能卡及其他设备定义 "
            Height          =   180
            Left            =   120
            TabIndex        =   99
            Top             =   4440
            Width           =   1890
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "药品名称显示"
            Height          =   180
            Left            =   120
            TabIndex        =   94
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label lblOverTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "超过       分钟未发药的药品处方"
            Height          =   180
            Left            =   1200
            TabIndex        =   93
            Top             =   1815
            Width           =   2790
         End
      End
      Begin VB.ComboBox cbo药品标签 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1560
         Width           =   2280
      End
      Begin VB.CheckBox chkUseDisplay 
         Caption         =   "显示排队队列"
         Height          =   255
         Left            =   -74640
         TabIndex        =   77
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cbo配药后 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1200
         Width           =   2280
      End
      Begin VB.ComboBox Cbo发药后 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1200
         Width           =   2280
      End
      Begin VB.CheckBox chk启用排队叫号 
         Caption         =   "启用排队叫号"
         Height          =   255
         Left            =   -74640
         TabIndex        =   52
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame fraSetColor 
         Caption         =   " 处方打印机定义 "
         Height          =   4965
         Left            =   -74880
         TabIndex        =   42
         Top             =   500
         Width           =   6795
         Begin VSFlex8Ctl.VSFlexGrid vsfPrinter 
            Height          =   3852
            Left            =   120
            TabIndex        =   116
            Top             =   600
            Width           =   6492
            _cx             =   11451
            _cy             =   6794
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
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
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
            Editable        =   1
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
         Begin VB.CommandButton cmdDefaultPrinter 
            BackColor       =   &H00000000&
            Caption         =   "恢复默认打印机(&P)"
            Height          =   300
            Left            =   3960
            MaskColor       =   &H00000000&
            TabIndex        =   51
            Top             =   4560
            Width           =   2655
         End
         Begin VB.Label Label6 
            Caption         =   "选择处方对应的打印机（仅对西药处方签、配药单）！"
            Height          =   432
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   4920
         End
      End
      Begin MSComctlLib.ListView Lvw来源科室 
         Height          =   4965
         Left            =   -74850
         TabIndex        =   40
         Top             =   495
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   8758
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.ComboBox cbo票据设置 
         Height          =   276
         Left            =   -73650
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   540
         Width           =   2565
      End
      Begin VB.CommandButton cmd打印设置 
         Caption         =   "打印设置(&P)"
         Height          =   345
         Left            =   -74400
         TabIndex        =   35
         Top             =   990
         Width           =   3315
      End
      Begin VB.Frame Fra打印 
         Caption         =   "  自动打印"
         Height          =   1692
         Left            =   -74850
         TabIndex        =   39
         Top             =   2040
         Width           =   6705
         Begin VB.CheckBox chk卫材发料单 
            Caption         =   "打印卫材发料单"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4320
            TabIndex        =   115
            Top             =   0
            Width           =   1920
         End
         Begin VB.CheckBox chkAllType 
            Caption         =   "自动打印配药单时打印票据的所有格式"
            Height          =   195
            Left            =   2760
            TabIndex        =   112
            Top             =   315
            Width           =   3360
         End
         Begin VB.CheckBox chk药品标签 
            Caption         =   "打印药品标签"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2640
            TabIndex        =   41
            Top             =   0
            Width           =   1440
         End
         Begin VB.CheckBox chk记帐单 
            Caption         =   "打印时包含记帐单据"
            Height          =   195
            Left            =   345
            TabIndex        =   18
            Top             =   315
            Width           =   1920
         End
         Begin VB.ListBox lst打印窗口 
            Appearance      =   0  'Flat
            Columns         =   1
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   2760
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   600
            Width           =   2760
         End
         Begin VB.CheckBox Chk打印配药单 
            Caption         =   "打印配药单"
            Height          =   210
            Left            =   1320
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Opt打印配药单本部门 
            Caption         =   "打印本部门的配药单"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   19
            Top             =   615
            Width           =   1935
         End
         Begin VB.OptionButton Opt打印配药单本窗口 
            Caption         =   "打印本窗口的配药单"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   20
            Top             =   1027
            Width           =   1935
         End
         Begin VB.OptionButton Opt打印配药单选择 
            Caption         =   "打印指定窗口配药单"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   21
            Top             =   1440
            Width           =   2100
         End
      End
      Begin VB.Frame Fra刷新 
         Caption         =   "  自动刷新 "
         Height          =   1680
         Left            =   -74850
         TabIndex        =   23
         Top             =   3840
         Width           =   6705
         Begin VB.TextBox txt打印间隔 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   47
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox Txt打印退费单号 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   31
            Top             =   1320
            Width           =   1125
         End
         Begin VB.TextBox Txt延迟打印 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   28
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox Txt刷新间隔 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   25
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lblRefreshComment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "已启用消息机制替代"
            Height          =   180
            Left            =   3120
            TabIndex        =   97
            Top             =   660
            Width           =   1620
         End
         Begin VB.Label lblPrintComment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "已启用消息机制替代"
            Height          =   180
            Left            =   3120
            TabIndex        =   96
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "打印间隔"
            Height          =   180
            Left            =   840
            TabIndex        =   49
            Top             =   300
            Width           =   720
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "秒"
            Height          =   180
            Index           =   4
            Left            =   2760
            TabIndex        =   48
            Top             =   285
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "分"
            Height          =   180
            Index           =   2
            Left            =   2760
            TabIndex        =   32
            Top             =   1380
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "秒"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   29
            Top             =   1020
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "秒"
            Height          =   180
            Index           =   0
            Left            =   2760
            TabIndex        =   26
            Top             =   660
            Width           =   180
         End
         Begin VB.Label Lbl打印退费单号 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "打印退费单据间隔"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   1380
            Width           =   1440
         End
         Begin VB.Label Lbl延迟打印 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "延迟打印"
            Height          =   180
            Left            =   840
            TabIndex        =   27
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Lbl刷新间隔 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "刷新间隔"
            Height          =   180
            Left            =   840
            TabIndex        =   24
            Top             =   660
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "  选项 "
         Height          =   5025
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6795
         Begin VB.CheckBox chkSame 
            Caption         =   "允许配药人和核查人相同"
            Height          =   195
            Left            =   3480
            TabIndex        =   113
            Top             =   3413
            Width           =   2640
         End
         Begin VB.ComboBox cboCheck 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   3360
            Width           =   2280
         End
         Begin VB.ComboBox cbo待配药 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   4560
            Width           =   2280
         End
         Begin VB.CheckBox chkIsDosage 
            Caption         =   "当前药房需要配药环节"
            Height          =   225
            Left            =   195
            TabIndex        =   105
            Top             =   1855
            Width           =   2940
         End
         Begin VB.CheckBox chkIsDosageOk 
            Caption         =   "当前药房需要配药确认(病人签到)环节"
            Height          =   225
            Left            =   195
            TabIndex        =   104
            Top             =   1560
            Width           =   3540
         End
         Begin VB.CheckBox chkSign 
            Caption         =   "签到时自动进行配药(药房窗口签到有效)"
            Height          =   180
            Left            =   195
            TabIndex        =   103
            Top             =   2150
            Width           =   3615
         End
         Begin VB.CheckBox chkCheckStuff 
            Caption         =   "发药后检查卫材发放情况"
            Height          =   180
            Left            =   195
            TabIndex        =   102
            Top             =   2400
            Width           =   2295
         End
         Begin VB.ComboBox cbo处方类型 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1080
            Width           =   2280
         End
         Begin VB.TextBox txt配药时限 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3150
            TabIndex        =   45
            Top             =   3060
            Width           =   525
         End
         Begin VB.CheckBox chk自动配药 
            Caption         =   "自动配药模式"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   3120
            Width           =   1440
         End
         Begin VB.ComboBox cbo单位 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   660
            Width           =   2280
         End
         Begin VB.ComboBox cbo记帐处方 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   4170
            Width           =   2280
         End
         Begin VB.TextBox txt查询天数 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            TabIndex        =   15
            Text            =   "1"
            Top             =   4560
            Width           =   885
         End
         Begin VB.ListBox lst发药窗口 
            Appearance      =   0  'Flat
            Columns         =   1
            ForeColor       =   &H80000012&
            Height          =   1710
            IMEMode         =   3  'DISABLE
            Left            =   4680
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   300
            Width           =   1800
         End
         Begin VB.ComboBox cbo收费处方 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3780
            Width           =   2280
         End
         Begin VB.ComboBox Cbo配药人 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2760
            Width           =   2280
         End
         Begin VB.ComboBox Cbo药房 
            ForeColor       =   &H80000012&
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   1320
            TabIndex        =   3
            Text            =   "Cbo药房"
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "核查人"
            Height          =   180
            Left            =   360
            TabIndex        =   111
            Top             =   3420
            Width           =   540
         End
         Begin VB.Label lbl配药打印状态 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "待配药处方"
            Height          =   180
            Left            =   60
            TabIndex        =   107
            Top             =   4620
            Width           =   900
         End
         Begin VB.Label lbl门诊住院 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊住院处方"
            Height          =   180
            Left            =   120
            TabIndex        =   81
            Top             =   1140
            Width           =   1080
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "分钟"
            Height          =   180
            Left            =   3720
            TabIndex        =   46
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "自动配药时限"
            Height          =   180
            Left            =   2040
            TabIndex        =   44
            Top             =   3120
            Width           =   1080
         End
         Begin VB.Label lbl单位 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "药房属性"
            Height          =   180
            Left            =   480
            TabIndex        =   12
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbl记帐处方 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "记帐处方"
            Height          =   180
            Left            =   240
            TabIndex        =   10
            Top             =   4230
            Width           =   720
         End
         Begin VB.Label lbl天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "天"
            Height          =   180
            Left            =   5280
            TabIndex        =   16
            Top             =   4620
            Width           =   180
         End
         Begin VB.Label lbl查询天数 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查询天数"
            Height          =   180
            Left            =   3480
            TabIndex        =   14
            Top             =   4620
            Width           =   720
         End
         Begin VB.Label lbl收费处方 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "收费处方"
            Height          =   180
            Left            =   240
            TabIndex        =   8
            Top             =   3840
            Width           =   720
         End
         Begin VB.Label Lbl配药人 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "配药人"
            Height          =   180
            Left            =   360
            TabIndex        =   6
            Top             =   2820
            Width           =   540
         End
         Begin VB.Label Lbl发药窗口 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "发药窗口"
            Height          =   180
            Left            =   3840
            TabIndex        =   4
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Lbl药房 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "发药药房"
            Height          =   180
            Left            =   480
            TabIndex        =   2
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame frm显示设备设置 
         Height          =   855
         Left            =   -74850
         TabIndex        =   53
         Top             =   840
         Width           =   6795
         Begin VB.CommandButton cmd显示设备设置 
            Caption         =   "设备设置"
            Height          =   300
            Left            =   4320
            TabIndex        =   55
            Top             =   300
            Width           =   1100
         End
         Begin VB.ComboBox cbo显示硬件类别 
            Height          =   300
            ItemData        =   "Frm发药参数设置.frx":03CE
            Left            =   1560
            List            =   "Frm发药参数设置.frx":03D0
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "显示设备类别"
            Height          =   180
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Fra语音设备设置 
         Height          =   3735
         Left            =   -74850
         TabIndex        =   57
         Top             =   1800
         Width           =   6795
         Begin VB.OptionButton optCallWay 
            Caption         =   "启用本地语音"
            Height          =   330
            Index           =   0
            Left            =   240
            TabIndex        =   76
            Top             =   320
            Width           =   1455
         End
         Begin VB.CheckBox chkUseSound 
            Caption         =   "启用语音呼叫"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optCallWay 
            Caption         =   "启用远端语音"
            Height          =   330
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Frame frm语音广播设置 
            Height          =   1455
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   6495
            Begin VB.OptionButton optSoundType 
               Caption         =   "微软语音"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   71
               Top             =   338
               Width           =   1095
            End
            Begin VB.OptionButton optSoundType 
               Caption         =   "系统语音"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   70
               Top             =   338
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.TextBox txtSpeed 
               Height          =   270
               Left            =   1080
               TabIndex        =   66
               Text            =   "65"
               Top             =   685
               Width           =   495
            End
            Begin VB.CommandButton cmdTestSound 
               Caption         =   "测试语音"
               Height          =   350
               Left            =   4080
               TabIndex        =   65
               Top             =   645
               Width           =   1100
            End
            Begin VB.TextBox txtPlayCount 
               Height          =   270
               Left            =   1080
               TabIndex        =   64
               Text            =   "1"
               Top             =   1035
               Width           =   615
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "语音类型"
               Height          =   180
               Left            =   120
               TabIndex        =   69
               Top             =   375
               Width           =   720
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "语音语速：      (范围在0到100之间，推荐65)"
               Height          =   180
               Left            =   120
               TabIndex        =   68
               Top             =   730
               Width           =   3780
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "播放次数为        次"
               Height          =   180
               Left            =   120
               TabIndex        =   67
               Top             =   1080
               Width           =   1800
            End
         End
         Begin VB.Frame Fra远端语音设置 
            Height          =   1215
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Width           =   6495
            Begin VB.TextBox txtLoopQueryTime 
               Height          =   270
               Left            =   1800
               MaxLength       =   3
               TabIndex        =   78
               Text            =   "10"
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboWorkStation 
               Height          =   300
               Left            =   1200
               TabIndex        =   61
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "数据轮询间隔时间为        秒"
               Height          =   180
               Left            =   120
               TabIndex        =   79
               Top             =   885
               Width           =   2520
            End
            Begin VB.Label labRemoteComputerName 
               Caption         =   "远端站点名："
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   405
               Width           =   1215
            End
         End
      End
      Begin VB.Label lbl中药处方格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中药处方格式"
         Height          =   180
         Left            =   -71460
         TabIndex        =   124
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label lbl中药配药格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中药配药格式"
         Height          =   180
         Left            =   -74940
         TabIndex        =   122
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label lbl西药处方格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "西药处方格式"
         Height          =   180
         Left            =   -71460
         TabIndex        =   120
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lbl西药配药格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "西药配药格式"
         Height          =   180
         Left            =   -74940
         TabIndex        =   118
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lbl药品标签 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品标签"
         Height          =   180
         Left            =   -74580
         TabIndex        =   83
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Lbl发药 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方签"
         Height          =   180
         Left            =   -70920
         TabIndex        =   75
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lbl配药 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "配药单"
         Height          =   180
         Left            =   -74400
         TabIndex        =   74
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lbl票据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "票据(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74370
         TabIndex        =   33
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   38
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   37
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   36
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm发药参数设置.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm发药参数设置.frx":06EC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm发药参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--注册表相关变量--
Private intDays As Integer
Private intUnit As Integer                              '缺省单位（0-自适应;1-门诊药房单位;2-住院药房单位）
Private intPrint As Integer                             '不打印未配药单据(0)
Private int校验方式 As Integer                          '校验方式
Private int校验配药人 As Integer                        '配药时是否校验配药人
Private int校验发药人 As Integer                        '发药时是否校验发药人
Private mint记帐单 As Integer                           '打印配药单时是否包含记帐单
Private mint药品标签 As Integer                         '打印药品标签
Private mint卫材发料单 As Integer                       '打印卫材发料单
Private strPrintWindow As String                        '打印未配药单据为3时有效
'0-不打印未配药单据
'1-打印本部门所有未配药单据
'2-打印本窗口所有未配药单据
'3-选择打印(发药窗口)

Private IntRefresh As Integer                           '刷新间隔(0)
Private intPrintDelay As Integer                        '延迟打印(60)
Private intPrintHandbackNO As Integer                   '打印退费单据号(0)
Private mintPrintInterval As Integer                    '打印配药单间隔(0)
Private lng药房ID As Long                               '药房(设置本机所对应的药房)
Private Str窗口 As String                               '发药窗口(设置本机所对应的发药窗口)
Private str配药人 As String                             '设置配药人
Private mint自动配药 As Integer                         '是否使用自动配药功能：0-不使用；1-使用
Private mint自动配药时限 As Integer                     '超过该时限就需要验证配药人：默认为始终不验证配药人
Private mint刷验证 As Integer                           '发药后是否进行刷卡验证：0-不刷卡;1-要刷卡
Private mint配药扫描 As Integer                         '配药模式启用扫描器：0-不启用;1-启用
Private mint启用排队叫号 As Integer                     '是否启用排队叫号功能
Private mintSign As Integer                             '签到时进行配药
Private mblnLoadDrug As Boolean
Private mblnUseMsg As Boolean                           '是否已启用消息机制
Private mstr两次刷卡发药 As String                      '两次刷卡发药，格式：卡类别1,卡类别2......，如果无内容表示不启用
Private mint发生时间过滤 As Integer                     '药品医嘱按发生时间（首次时间）过滤：0-按产生时间过滤，1-按发生时间过滤
Private mint金额显示方式 As Integer                     '0-显示应收金额，1-显示实收金额，2-显示应收金额和实收金额
Private mint病人取药模式 As Integer                     '病人取药模式：0-不启用，1-启用
Private mint发药后检查 As Integer                       '发药后是否检查有本药房此病人有未发的卫材单据
Private mint扫描后呼叫 As Integer                       '0-不自动呼叫,1-扫描后自动语音呼叫
Private mstr核查人 As String
Private mintRowNum As Integer


Private mintShowName As Integer                         '药品名称显示方式：0-名称和编码；1-仅编码；2-仅名称
Private mintType As Integer                             '处方类型：0-显示门诊和住院处方；1-只显示门诊处方；2-只显示住院处方

Private IntShowCol As Integer                           '在处方明细中是否显示付数(0)
Private mintShowBill收费 As Integer                     '收费处方显示范围
Private mintShowBill记帐 As Integer                     '记帐处方显示范围
Private mintShowBill配药 As Integer                     '待配药单打印状态显示范围
Private IntAutoPrint As Integer                         '发药后打印处方单(1)
Private mint配药后自动打印 As Integer                   '自动打印配药单
Private mint发药后自动打印药品标签 As Integer           '发药后自动打印药品标签
Private mstrWin As String                               '发药窗口串
Private mint回车方式 As Integer                         '通过录入或刷卡查找时系统自动添加回车处理的方式，0-系统不自动回车,1-当录入达到项目或卡号长度时自动回车

Private mIntCol类型 As Integer
Private mintCol格式 As Integer
Private mintCol打印机 As Integer

Private Const mconstr类型 = "普通;儿科;急诊;精神Ⅱ类;精神Ⅰ类;麻醉"
Private Const mconlng颜色 = "&HFFFFFF;&HC0FFC0;&HC0FFFF;&HFFFFFF;&HC0C0FF;&HC0C0FF"

Public mstrPrivs As String                              '权限串
Private mblnSetPara As Boolean                          '是否具有参数设置权限
Private mstrRPTDefaultScheme_Recipt As String           '处方签报表的默认格式

'排队叫号使用的参数

Private Type Type_Call
    int启用排队叫号 As Integer
    int语音类型 As Integer
'    int显示模式 As Integer
    int显示排队队列 As Integer
    int启用语音呼叫 As Integer
    int叫号方式 As Integer
    str远端呼叫站点 As String
    int语音广播语速 As Integer
    int语音播放次数 As Integer
    int轮询时间 As Integer
End Type

Private mType_Call As Type_Call
'--本程序所使用的东东--
Public RecPart As New ADODB.Recordset                   '药房
Private RecPeople As New ADODB.Recordset                '药房发药人
Private BlnStartUp  As Boolean                          '是否启动成功
Public strShow As String                                '显示串
Private mstrSourceDep As String                         '来源科室串

Private mstrPrinters As String                          '本地打印机列表，用;分隔

'处方类型：普通、急诊、儿科、麻醉、精一、精二
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

'默认处方颜色：普通－白色；急诊－淡黄色；儿科－淡绿色；麻醉、精一－淡红色；精二－白色
Private Const mconlng普通 = &HFFFFFF
Private Const mconlng儿科 = &HC0FFC0
Private Const mconlng急诊 = &HC0FFFF
Private Const mconlng精二 = &HFFFFFF
Private Const mconlng精一 = &HC0C0FF
Private Const mconlng麻醉 = &HC0C0FF

Public Property Get In_启用发药() As Boolean
    In_启用发药 = mblnLoadDrug
End Property

Public Property Let In_启用发药(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property

Public Property Get In_启用消息() As Boolean
    In_启用消息 = mblnUseMsg
End Property

Public Property Let In_启用消息(ByVal vNewValue As Boolean)
    mblnUseMsg = vNewValue
End Property









Private Sub LoadList()
    Dim rs西药格式 As New ADODB.Recordset
    Dim rs中药格式 As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str配药格式 As String
    Dim str处方格式 As String
    Dim str编号 As String
    Dim strPrinter As String
    Dim strPrinters As String
    Dim strColor As String
    Dim myPrinter As Printer
    Dim n As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mIntCol类型 = 0
    mintCol格式 = 1
    mintCol打印机 = 2
    
    '获取报表格式
    '--西药
    str编号 = "ZL1_BILL_1341_3"
    
    gstrSQL = "Select b.说明 From zlReports A, zlRPTFMTs B Where a.Id = b.报表id And a.编号 = [1] order by b.序号"
    
    Set rs西药格式 = zldatabase.OpenSQLRecord(gstrSQL, "提取西药报表格式", str编号)
    
    '--中药
    str编号 = "ZL1_BILL_1341_4"
    
    gstrSQL = "Select b.说明 From zlReports A, zlRPTFMTs B Where a.Id = b.报表id And a.编号 = [1] order by b.序号"
    
    Set rs中药格式 = zldatabase.OpenSQLRecord(gstrSQL, "提取中药报表格式", str编号)
    
    '获取处方类型的颜色设置
    strColor = zldatabase.GetPara("处方颜色", glngSys, 1341, "", , mblnSetPara)
    
    '获取保存的打印机参数设置
    strPrinter = zldatabase.GetPara("处方对应的打印机", glngSys, 1341, "", , mblnSetPara)
    
    '获取对应的打印格式设置
    str配药格式 = zldatabase.GetPara("配药单打印格式", glngSys, 1341, "2;2", , mblnSetPara)
    str处方格式 = zldatabase.GetPara("处方签打印格式", glngSys, 1341, "1;1", , mblnSetPara)
    
    '添加打印格式至下拉列表
    With rs西药格式
        For n = 1 To .RecordCount
            cbo西药配药格式.AddItem !说明
            cbo西药配药格式.ItemData(cbo西药配药格式.NewIndex) = n
            cbo西药处方格式.AddItem !说明
            cbo西药处方格式.ItemData(cbo西药处方格式.NewIndex) = n
            .MoveNext
        Next
    End With
    
    With rs中药格式
        For n = 1 To .RecordCount
            cbo中药配药格式.AddItem !说明
            cbo中药配药格式.ItemData(cbo中药配药格式.NewIndex) = n
            cbo中药处方格式.AddItem !说明
            cbo中药处方格式.ItemData(cbo中药处方格式.NewIndex) = n
            .MoveNext
        Next
    End With
    
    '加载用户设置的打印格式
    '--西药
    For i = 0 To cbo西药配药格式.ListCount - 1
        If Val(Split(str配药格式, ";")(0)) = cbo西药配药格式.ItemData(i) Then
            cbo西药配药格式.ListIndex = i
        End If
    Next
    
    For i = 0 To cbo西药处方格式.ListCount - 1
        If Val(Split(str处方格式, ";")(0)) = cbo西药处方格式.ItemData(i) Then
            cbo西药处方格式.ListIndex = i
        End If
    Next
    '--中药
    For i = 0 To cbo中药配药格式.ListCount - 1
        If Val(Split(str配药格式, ";")(1)) = cbo中药配药格式.ItemData(i) Then
            cbo中药配药格式.ListIndex = i
        End If
    Next
    
    For i = 0 To cbo中药处方格式.ListCount - 1
        If Val(Split(str处方格式, ";")(1)) = cbo中药处方格式.ItemData(i) Then
            cbo中药处方格式.ListIndex = i
        End If
    Next
    
    '载入本地打印机列表
    mstrPrinters = ""
    For Each myPrinter In Printers
        mstrPrinters = IIf(mstrPrinters = "", "", mstrPrinters & ";") & myPrinter.DeviceName
    Next
    
    For n = 0 To UBound(Split(mstrPrinters, ";"))
        If Split(mstrPrinters, ";")(n) <> "" Then
            strPrinters = strPrinters & "|" & Split(mstrPrinters, ";")(n)
        End If
    Next
    strPrinters = Mid(strPrinters, 2)
    
    '装载本地记录集
    With rsData
        If .State = 1 Then .Close
        
        .Fields.Append "类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "格式", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "打印机", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '判断数据的合法性，并对其修正
    If UBound(Split(strPrinter, ";")) <> UBound(Split(mconstr类型, ";")) Then
        For n = 0 To UBound(Split(mconstr类型, ";"))
            strPrinter = strPrinter & ";"
        Next
    End If
    
    '向本地记录集载入用户保存的打印机设置
    For n = 0 To UBound(Split(mconstr类型, ";"))
        rs西药格式.MoveFirst
        If InStr(strPrinter, "?") = 0 Then
            For i = 1 To rs西药格式.RecordCount
                rsData.AddNew
                
                rsData!类型 = Split(mconstr类型, ";")(n)
                rsData!格式 = rs西药格式!说明
                rsData!打印机 = Split(strPrinter, ";")(n)
    
                rsData.Update
                rs西药格式.MoveNext
            Next
        Else
            For i = 0 To UBound(Split(Split(strPrinter, ";")(n), ","))
                rsData.AddNew
                
                rsData!类型 = Split(mconstr类型, ";")(n)
                rsData!格式 = Mid(Split(Split(strPrinter, ";")(n), ",")(i), 1, InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") - 1)
                rsData!打印机 = Mid(Split(Split(strPrinter, ";")(n), ",")(i), InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") + 1)
             
            Next
        End If
        rsData.Update
    Next
        
    With vsfPrinter
        .rows = rs西药格式.RecordCount * 6
        .Cols = 3
        .AllowSelection = False
        .ColAlignment(mIntCol类型) = flexAlignCenterCenter
        .RowHeight(-1) = 250
        .ColWidth(mIntCol类型) = 900
        .ColWidth(mintCol格式) = 1500
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(mIntCol类型) = True
        
        '加载打印机选项至表格
        .ColComboList(mintCol打印机) = strPrinters
        
        '加载[类型&颜色]、[格式]
        For n = 0 To UBound(Split(mconstr类型, ";"))
            rs西药格式.MoveFirst
            For i = 1 To rs西药格式.RecordCount
                .TextMatrix(n * rs西药格式.RecordCount + i - 1, mIntCol类型) = Split(mconstr类型, ";")(n)
                
                If strColor <> "" Then
                    .Cell(flexcpBackColor, n * rs西药格式.RecordCount + i - 1, mIntCol类型) = Val(Split(strColor, ";")(n))
                Else
                    .Cell(flexcpBackColor, n * rs西药格式.RecordCount + i - 1, mIntCol类型) = Split(mconlng颜色, ";")(n)
                End If
                
                .TextMatrix(n * rs西药格式.RecordCount + i - 1, mintCol格式) = rs西药格式!说明
                
                rs西药格式.MoveNext
            Next
        Next
        
        '载入用户保存的打印机设置
        For n = 0 To .rows - 1
            rsData.Filter = "类型 = '" & .TextMatrix(n, mIntCol类型) & "' and 格式 = '" & .TextMatrix(n, mintCol格式) & "'"
            If rsData.RecordCount > 0 Then
                If InStr(strPrinters & "|", rsData!打印机 & "|") > 0 Then   '检查该打印机本地是否存在
                    .TextMatrix(n, mintCol打印机) = rsData!打印机
                End If
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ReadFromReg()
    Dim strTmp As String
    Dim intOverTime As Integer
    Dim intParaType As Integer
    
    On Error Resume Next
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "参数设置")

    '取私有参数
    mintShowBill收费 = Val(zldatabase.GetPara("收费处方显示方式", glngSys, 1341, 3, Array(lbl收费处方, cbo收费处方), mblnSetPara))
    mintShowBill记帐 = Val(zldatabase.GetPara("记帐处方显示方式", glngSys, 1341, 3, Array(lbl记帐处方, cbo记帐处方), mblnSetPara))
    mintShowBill配药 = Val(zldatabase.GetPara("待配药单据打印显示方式", glngSys, 1341, 0, Array(lbl配药打印状态, cbo待配药), mblnSetPara))
    intDays = Val(zldatabase.GetPara("查询天数", glngSys, 1341, 1, Array(lbl查询天数, txt查询天数, lbl天数), mblnSetPara))
    mint记帐单 = Val(zldatabase.GetPara("打印包含记帐单", glngSys, 1341, 0, Array(chk记帐单), mblnSetPara))
    intPrintHandbackNO = Val(zldatabase.GetPara("打印退费单据间隔", glngSys, 1341, 0, Array(Lbl打印退费单号, Txt打印退费单号, lblNote(2)), mblnSetPara))
    intPrintDelay = Val(zldatabase.GetPara("打印延迟", glngSys, 1341, 60, Array(Lbl延迟打印, Txt延迟打印, lblNote(1)), mblnSetPara))
    IntRefresh = Val(zldatabase.GetPara("刷新间隔", glngSys, 1341, 0, Array(Lbl刷新间隔, Txt刷新间隔, lblNote(0)), mblnSetPara))
    mintPrintInterval = Val(zldatabase.GetPara("打印间隔", glngSys, 1341, 0, Array(Label3, txt打印间隔, lblNote(4)), mblnSetPara))
    IntShowCol = Val(zldatabase.GetPara("显示付数", glngSys, 1341, 0, Array(Chk显示付数栏), mblnSetPara))
    IntAutoPrint = Val(zldatabase.GetPara("发药后自动打印", glngSys, 1341, 0, Array(Lbl发药, Cbo发药后), mblnSetPara))
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0, Array(lbl单位, cbo单位), mblnSetPara))
    mint配药后自动打印 = Val(zldatabase.GetPara("配药后自动打印", glngSys, 1341, 2, Array(lbl配药, cbo配药后), mblnSetPara))
    mint发药后自动打印药品标签 = Val(zldatabase.GetPara("发药后打印药品标签", glngSys, 1341, 2, Array(lbl药品标签, cbo药品标签), mblnSetPara))
    
    mint配药扫描 = Val(zldatabase.GetPara("配药模式扫描器确认", glngSys, 1341, 0, Array(chk配药扫描), mblnSetPara))
    intOverTime = Val(zldatabase.GetPara("超时未发药品显示时间间隔", glngSys, 1341, 0, Array(chkOverTime, lblOverTime, txtOverTime, fraline1), mblnSetPara))
    mintType = Val(zldatabase.GetPara("发门诊住院处方", glngSys, 1341, 0, Array(lbl门诊住院, cbo处方类型), mblnSetPara))
    mintSign = Val(zldatabase.GetPara("签到时进行配药", glngSys, 1341, 0, Array(chkSign), mblnSetPara))
    mstr两次刷卡发药 = zldatabase.GetPara("两次刷卡发药", glngSys, 1341, "", Array(chk发药刷卡, lst卡类型), mblnSetPara)
    
    mint金额显示方式 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0, Array(lbl金额显示, cbo金额显示), mblnSetPara))
    
    mint发药后检查 = zldatabase.GetPara("发药后检查卫材发放情况", glngSys, 1341, 0, Array(chkCheckStuff), mblnSetPara)
    mint扫描后呼叫 = Val(zldatabase.GetPara("待发药单据扫描后自动呼叫", glngSys, 1341, 0, Array(chk扫描后呼叫), mblnSetPara))
    mint回车方式 = Val(zldatabase.GetPara("查找时系统自动回车方式", glngSys, 1341, 0, Array(cbo回车方式), mblnSetPara))
   
    '0-不打印未配药单据
    '1-打印本部门所有未配药单据
    '2-打印本窗口所有未配药单据
    '3-选择打印(发药窗口)
    intPrint = Val(zldatabase.GetPara("发现新单据是否打印", glngSys, 1341, 0, Array(Chk打印配药单), mblnSetPara))
    
    mint药品标签 = Val(zldatabase.GetPara("打印药品标签", glngSys, 1341, 0, Array(chk药品标签), mblnSetPara))
    mint卫材发料单 = Val(zldatabase.GetPara("打印卫材发料单", glngSys, 1341, 0, Array(chk卫材发料单), mblnSetPara))
    lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1341, 0, Array(lbl药房, , Cbo药房), mblnSetPara))
    Str窗口 = zldatabase.GetPara("发药窗口", glngSys, 1341, "", Array(Lbl发药窗口, lst发药窗口), mblnSetPara)
    str配药人 = zldatabase.GetPara("配药人", glngSys, 1341, "", Array(Lbl配药人, cbo配药人), mblnSetPara)
    strPrintWindow = zldatabase.GetPara("打印指定发药窗口", glngSys, 1341, "", Array(Opt打印配药单选择, lst打印窗口), mblnSetPara)
    mstrSourceDep = zldatabase.GetPara("来源科室", glngSys, 1341, "", Array(Lvw来源科室), mblnSetPara)
    mint自动配药 = Val(zldatabase.GetPara("自动配药", glngSys, 1341, 0, Array(chk自动配药), mblnSetPara))
    mint自动配药时限 = Val(zldatabase.GetPara("自动配药时限", glngSys, 1341, 0, Array(Label1, txt配药时限, Label2), mblnSetPara))
    mstr核查人 = zldatabase.GetPara("核查人", glngSys, 1341, "", Array(lblCheck, cboCheck), mblnSetPara)
    
    If lng药房ID <> 0 Then
        Call SetDispense
    End If
    
    strTmp = zldatabase.GetPara("列设置", glngSys, 1341, "0", Array(Label4, cbo药品名称显示), mblnSetPara)
    If InStr(1, strTmp, "|") > 0 Then
        mintShowName = Val(Mid(strTmp, 1, 1))
    Else
        mintShowName = Val(strTmp)
    End If
    If mintShowName > 2 Or mintShowName < 0 Then mintShowName = 0
    
    chk大小单位.Value = Val(zldatabase.GetPara("显示大小单位", glngSys, 1341, 0, Array(chk大小单位), mblnSetPara))
    chkAllType.Value = (zldatabase.GetPara("打印票据的所有格式", glngSys, 1341, 0, Array(chkAllType), mblnSetPara))
    chkSame.Value = (zldatabase.GetPara("允许核查人和配药人相同", glngSys, 1341, 0, Array(chkSame), mblnSetPara))
    chkPreview.Value = zldatabase.GetPara("打印处方签时先预览再打印", glngSys, 1341, 0, Array(chkPreview), mblnSetPara)
    
    If intOverTime < 0 Or intOverTime > 1440 Then
        intOverTime = 0
    End If
    intOverTime = Int(intOverTime)
    chkOverTime.Value = IIf(intOverTime = 0, 0, 1)
    If chkOverTime.Value = 0 Then
        txtOverTime.Text = ""
        txtOverTime.Enabled = False
    Else
        txtOverTime.Text = intOverTime
        txtOverTime.Enabled = True
    End If
    
    '处方颜色和打印机设置
    Call LoadList
    
    '排队叫号参数按部门存取
    With mType_Call
        .int叫号方式 = Val(zldatabase.GetPara("叫号方式", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int启用排队叫号 = Val(zldatabase.GetPara("启用排队叫号", glngSys, 1341, 0, Array(chk启用排队叫号), mblnSetPara, intParaType, lng药房ID))
        .int启用语音呼叫 = Val(zldatabase.GetPara("启用语音呼叫", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
'        .int显示模式 = Val(zldatabase.GetPara("显示模式", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int显示排队队列 = Val(zldatabase.GetPara("显示排队队列", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音播放次数 = Val(zldatabase.GetPara("语音播放次数", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音广播语速 = Val(zldatabase.GetPara("语音广播语速", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音类型 = Val(zldatabase.GetPara("语音类型", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .str远端呼叫站点 = zldatabase.GetPara("远端呼叫站点", glngSys, 1341, "", Null, mblnSetPara, intParaType, lng药房ID)
        .int轮询时间 = Val(zldatabase.GetPara("呼叫轮询时间", glngSys, 1341, 10, Null, mblnSetPara, intParaType, lng药房ID))
    End With
End Function

Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select 编码 || '-' || 名称 科室, Id " & _
            " From 部门表 " & _
            " Where Id In (Select 部门id From 部门性质说明 Where 工作性质 = '临床' And 服务对象 In (1,2,3)) And " & _
            " (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By 编码 || '-' || 名称 "

    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "SetSourceDep")
    Call SQLTest

    With rs
        If .EOF Then
            MsgBox "没有设置该类部门！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw来源科室.ListItems.Clear
        Do While Not .EOF
            Lvw来源科室.ListItems.Add , "_" & !Id, !科室, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw来源科室.ListItems("_" & !Id).Checked = True
                End If
            End If
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

Private Sub Cbo药房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo药房.ListCount = 0 Then Exit Sub
    
    If Cbo药房.ListIndex >= 0 Then
        If Val(Cbo药房.Tag) = Cbo药房.ItemData(Cbo药房.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, Cbo药房, Trim(Cbo药房.Text), str工作性质, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有药房"), False, True), "0,1,2,3") = False Then
        Exit Sub
    End If
    If Cbo药房.ListIndex >= 0 Then
        Cbo药房.Tag = Cbo药房.ItemData(Cbo药房.ListIndex)
    End If
End Sub

Private Sub Cbo药房_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cbo药房_Validate(Cancel As Boolean)
    If Cbo药房.ListCount > 0 Then
        If Cbo药房.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub chkIsDosage_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
    
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "已启用消息机制替代", "待配药环节已启用消息机制替代自动刷新")
End Sub

Private Sub chkIsDosageOk_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
End Sub

Private Sub chkOverTime_Click()
    If chkOverTime.Value = 1 Then
        txtOverTime.Enabled = True
        If Int(Val(txtOverTime.Text)) = 0 Then
            txtOverTime.Text = "30"
        End If
    Else
        txtOverTime.Enabled = False
    End If
End Sub

Private Sub chkUseDisplay_Click()
    If Me.chkUseDisplay.Value = 0 Then
        frm显示设备设置.Enabled = False
    Else
        frm显示设备设置.Enabled = True
    End If
End Sub

Private Sub chkUseSound_Click()
    If Me.chkUseSound.Value = 1 Then
        frm语音广播设置.Enabled = True
        Fra远端语音设置.Enabled = True
        Me.optCallWay(0).Enabled = True
        Me.optCallWay(1).Enabled = True
    Else
        frm语音广播设置.Enabled = False
        Fra远端语音设置.Enabled = False
        Me.optCallWay(0).Enabled = False
        Me.optCallWay(1).Enabled = False
    End If
End Sub

Private Sub chk发药刷卡_Click()
    lst卡类型.Enabled = (chk发药刷卡.Value = 1)
End Sub

Private Sub chk自动配药_Click()
    If chk自动配药.Value = 1 Then
        txt配药时限.Enabled = chk自动配药.Enabled
    Else
        txt配药时限.Enabled = False
    End If
End Sub

Private Sub cmdDefaultPrinter_Click()
    Dim strDefault As String
    Dim n As Integer
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    
    '取报表的格式名称（默认取第一个格式）
    If mstrRPTDefaultScheme_Recipt = "" Then
        Set rsData = DeptSendWork_Get发药单格式("ZL1_BILL_1341_3")
        If Not rsData.EOF Then mstrRPTDefaultScheme_Recipt = rsData!格式
    End If
    
    '兼容以前的版本，依次从不同的位置取值
'    If mstrRPTDefaultScheme_Recipt <> "" Then strDefault = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3\" & mstrRPTDefaultScheme_Recipt, "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3\所有格式", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        
    If strDefault = "" Or InStr(1, ";" & mstrPrinters & ";", ";" & strDefault & ";") = 0 Then
        '如果默认打印机为空，或者不在本地打印机列表中时
        MsgBox "没有设置西药处方签对应的打印机，请在“票据(4)”中设置！", vbInformation, gstrSysName
        TabShow.Tab = 3
        Exit Sub
    Else
        '设置默认的打印机
        For n = 0 To vsfPrinter.rows - 1
            vsfPrinter.TextMatrix(n, mintCol打印机) = strDefault
        Next
        
    End If
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, 100, 1341)
End Sub

Private Sub cmdTestSound_Click()
    On Error GoTo errHandle
    If optSoundType(1).Value = True Then
        '微软语音
        Call zlCall_MsSoundPlay("请、" & "黄志杰、" & "黄志杰、" & "、到一号窗口", Val(txtSpeed.Text))
    Else
        '系统语音
        Call zlCall_SystemSoundPlay("请、" & "黄志杰、" & "黄志杰、" & "、到一号窗口", Val(txtSpeed.Text))
    End If
    Exit Sub
errHandle:
    Call SaveErrLog
End Sub
Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case cbo票据设置.ListIndex
    Case 0
        '西药处方签
        strBill = "ZL1_BILL_1341_3"
    Case 1
        '中药处方签
        strBill = "ZL1_BILL_1341_4"
    Case 2
        '处方发药清单
        strBill = "ZL1_BILL_1341_2"
    Case 3
        '处方退药通知单
        strBill = "ZL1_BILL_1341_1"
    Case 4
        '记帐处方统计表
        strBill = "ZL1_INSIDE_1341"
    Case 5
        '西药药品标签
        strBill = "ZL1_BILL_1341_6"
    Case 6
        '中草药药品标签
        strBill = "ZL1_BILL_1341_7"
    Case 7
        '已退费单据
        strBill = "ZL1_BILL_1341_8"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub cmd显示设备设置_Click()
    If gobjLEDShow Is Nothing Then
        If Not CreateObject_LED(Val(cbo显示硬件类别.ItemData(cbo显示硬件类别.ListIndex))) Then Exit Sub
    End If
        
    If Not gobjLEDShow Is Nothing Then
        Call gobjLEDShow.zlDrugSetup(Me, mstrWin)
    End If
End Sub

Private Sub lst打印窗口_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Cbo药房_Click()
    Dim intDO As Integer
    Dim bln门诊 As Boolean, bln住院 As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim intParaType As Integer
    Dim n As Integer
    
    On Error GoTo errHandle
    
    '窗口加载过程中导致的Click不执行
    If BlnStartUp = False Then Exit Sub
    
    '不可能，如果没有设置药房，主界面都进不了
    If Me.Cbo药房.ListCount = 0 Then Exit Sub
    
    lng药房ID = Cbo药房.ItemData(Cbo药房.ListIndex)
    
    '重新读取药房对应的发药窗口及药房人员
    Call ReadWindowsAndPeople
    
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341))
    
    '设置配药参数
    SetDispense
    
    '根据药房显示单位
    gstrSQL = " Select distinct 服务对象 From 部门性质说明" & _
              " Where 部门ID=[1] And 工作性质 like '%药房'" & _
              " Order By 服务对象 Desc"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取药房服务对象]", Cbo药房.ItemData(Cbo药房.ListIndex))
    
    rstemp.Filter = "服务对象=3"
    If rstemp.RecordCount <> 0 Then bln门诊 = True: bln住院 = True
    rstemp.Filter = "服务对象=2"
    If rstemp.RecordCount <> 0 Then bln住院 = True
    rstemp.Filter = "服务对象=1"
    If rstemp.RecordCount <> 0 Then bln门诊 = True
    rstemp.Filter = 0
    
    With cbo单位
        .Clear
        .AddItem "1-自适应"
        .ItemData(.NewIndex) = 0
        If bln门诊 Then
            .AddItem "2-门诊药房"
            .ItemData(.NewIndex) = 1
        End If
        If bln住院 Then
            .AddItem "3-住院药房"
            .ItemData(.NewIndex) = 2
        End If
        .ListIndex = 0
        
        For intDO = 0 To .ListCount - 1
            If .ItemData(intDO) = intUnit Then
                .ListIndex = intDO
                Exit For
            End If
        Next
    End With
    
    '重新提取排队叫号参数
    With mType_Call
        .int叫号方式 = Val(zldatabase.GetPara("叫号方式", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int启用排队叫号 = Val(zldatabase.GetPara("启用排队叫号", glngSys, 1341, 0, Array(chk启用排队叫号), mblnSetPara, intParaType, lng药房ID))
        .int启用语音呼叫 = Val(zldatabase.GetPara("启用语音呼叫", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
'        .int显示模式 = Val(zldatabase.GetPara("显示模式", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int显示排队队列 = Val(zldatabase.GetPara("显示排队队列", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音播放次数 = Val(zldatabase.GetPara("语音播放次数", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音广播语速 = Val(zldatabase.GetPara("语音广播语速", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .int语音类型 = Val(zldatabase.GetPara("语音类型", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lng药房ID))
        .str远端呼叫站点 = zldatabase.GetPara("远端呼叫站点", glngSys, 1341, "", Null, mblnSetPara, intParaType, lng药房ID)
        .int轮询时间 = Val(zldatabase.GetPara("呼叫轮询时间", glngSys, 1341, 10, Null, mblnSetPara, intParaType, lng药房ID))
    End With
    
   
    '重新设置排队叫号控件状态
    With mType_Call
        chk启用排队叫号.Value = .int启用排队叫号
        chkUseDisplay.Value = .int显示排队队列
        chkUseSound.Value = .int启用语音呼叫
        
        If .int叫号方式 = 0 Then
            optCallWay(0).Value = True
        Else
            optCallWay(1).Value = True
        End If
        
        optSoundType(.int语音类型).Value = 1
        txtSpeed.Text = .int语音广播语速
        txtPlayCount.Text = .int语音播放次数
        Me.cboWorkStation.Text = .str远端呼叫站点
        txtLoopQueryTime.Text = .int轮询时间
        
        chkUseDisplay_Click
        chkUseSound_Click
        
        If Me.optCallWay(0).Value = True Then
            optCallWay_Click 0
        Else
            optCallWay_Click 1
        End If
    End With


    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk打印配药单_Click()
    Dim ConState As Boolean
    
    ConState = (Chk打印配药单.Value = 1 And Chk打印配药单.Enabled = True)
    chk药品标签.Enabled = ConState
    If Not ConState And chk药品标签.Value = 1 Then
        chk药品标签.Value = 0
    End If
    
    chk卫材发料单.Enabled = ConState
    If Not ConState And chk卫材发料单.Value = 1 Then
        chk卫材发料单.Value = 0
    End If
    
    Opt打印配药单本部门.Enabled = ConState
    Opt打印配药单本窗口.Enabled = ConState
    Opt打印配药单选择.Enabled = ConState
    If Not ConState Then lst打印窗口.Enabled = False
    
    If BlnStartUp = False Then Exit Sub
    
    If ConState Then
        If Opt打印配药单本部门.Enabled = True Then Opt打印配药单本部门.SetFocus
    End If
End Sub

Private Sub Chk打印配药单_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim IntPrintStyle As Integer, i As Integer
    Dim strWin1 As String, strWin2 As String
    Dim intTemp As Integer
    Dim n As Integer
    Dim strPrinters As String
    Dim intSendCount As Integer
    Dim strCardType As String
    
    If Trim(txt查询天数.Text) = "" Then
        txt查询天数.Text = "1"
'        MsgBox "请输入查询天数（1天-365天）！", vbInformation, gstrSysName
'        txt查询天数.SetFocus
'        Exit Sub
    End If
    If Not IsNumeric(txt查询天数.Text) Then
        MsgBox "查询天数中含有非法字符！", vbInformation, gstrSysName
        If txt查询天数.Enabled = True Then txt查询天数.SetFocus
        Exit Sub
    End If
    If Val(txt查询天数.Text) < 1 Or Val(txt查询天数.Text) > 365 Then
        MsgBox "查询天数不能小于1天或大于365天！", vbInformation, gstrSysName
        If txt查询天数.Enabled = True Then txt查询天数.SetFocus
        Exit Sub
    End If
    
    If Trim(Txt刷新间隔) <> "" Then
        If Not IsNumeric(Txt刷新间隔) Then
            MsgBox "刷新间隔中含有非法字符！", vbInformation, gstrSysName
            If Txt刷新间隔.Enabled = True Then Txt刷新间隔.SetFocus
            Exit Sub
        End If
        If Val(Txt刷新间隔) < 0 Or Val(Txt刷新间隔) > 60 Then
            MsgBox "刷新间隔值超过范围（0至60）！", vbInformation, gstrSysName
            If Txt刷新间隔.Enabled = True Then Txt刷新间隔.SetFocus
            Exit Sub
        End If
        Txt刷新间隔 = CInt(Txt刷新间隔)
    End If
    If Trim(txt打印间隔) <> "" Then
        If Not IsNumeric(txt打印间隔) Then
            MsgBox "打印间隔中含有非法字符！", vbInformation, gstrSysName
            If txt打印间隔.Enabled = True Then txt打印间隔.SetFocus
            Exit Sub
        End If
        If Val(txt打印间隔) < 0 Or Val(txt打印间隔) > 60 Then
            MsgBox "打印间隔值超过范围（0至60）！", vbInformation, gstrSysName
            If txt打印间隔.Enabled = True Then txt打印间隔.SetFocus
            Exit Sub
        End If
        txt打印间隔 = CInt(txt打印间隔)
    End If
    If Trim(Txt延迟打印) <> "" Then
        If Not IsNumeric(Txt延迟打印) Then
            MsgBox "延迟打印中含有非法字符！", vbInformation, gstrSysName
            If Txt延迟打印.Enabled = True Then Txt延迟打印.SetFocus
            Exit Sub
        End If
        If Val(Txt延迟打印) < 0 Or Val(Txt延迟打印) > 60 Then
            MsgBox "延迟打印值超过范围（0至60）！", vbInformation, gstrSysName
            If Txt延迟打印.Enabled = True Then Txt延迟打印.SetFocus
            Exit Sub
        End If
        Txt延迟打印 = CInt(Txt延迟打印)
    End If
    If Trim(Txt打印退费单号) <> "" Then
        If Not IsNumeric(Txt打印退费单号) Then
            MsgBox "退费单号中含有非法字符！", vbInformation, gstrSysName
            If Txt打印退费单号.Enabled = True Then Txt打印退费单号.SetFocus
            Exit Sub
        End If
        If Val(Txt打印退费单号) < 0 Or Val(Txt打印退费单号) > 60 Then
            MsgBox "打印退费单值超过范围（0至60）！", vbInformation, gstrSysName
            If Txt打印退费单号.Enabled = True Then Txt打印退费单号.SetFocus
            Exit Sub
        End If
        Txt打印退费单号 = CInt(Txt打印退费单号)
    End If
    
    '检查本机所管窗口:如果有,至少要选择一个
    For i = 0 To lst发药窗口.ListCount - 1
        If lst发药窗口.Selected(i) Then
            strWin1 = strWin1 & ",'" & lst发药窗口.List(i) & "'"
            intSendCount = intSendCount + 1
        End If
    Next
    
    '如果启用排队叫号，则本机只能设置一个发药窗口
    If intSendCount > 1 And chk启用排队叫号.Value = 1 Then
        MsgBox "已启用排队叫号，只能设置一个发药窗口！", vbInformation, gstrSysName
        If lst发药窗口.Enabled = True Then lst发药窗口.SetFocus: Exit Sub
    End If
    
    If mblnLoadDrug And intSendCount > 1 Then
        MsgBox "已启用门诊自动发药，只能设置一个发药窗口！", vbInformation, gstrSysName
        If lst发药窗口.Enabled = True Then lst发药窗口.SetFocus: Exit Sub
    End If
    
    strWin1 = Mid(strWin1, 2)
    If strWin1 = "" And lst发药窗口.ListCount > 0 Then
        MsgBox "请指定本工作站所对应的发药窗口。", vbInformation, gstrSysName
        If lst发药窗口.Enabled = True Then lst发药窗口.SetFocus: Exit Sub
    End If
'    If UBound(Split(strWin1, ",")) + 1 = lst发药窗口.ListCount Then strWin1 = ""
       
    
    '检查打印发药窗口:不管是否有,至少要选择一个
    For i = 0 To lst打印窗口.ListCount - 1
        If lst打印窗口.Selected(i) Then
            strWin2 = strWin2 & ",'" & lst打印窗口.List(i) & "'"
        End If
    Next
    strWin2 = Mid(strWin2, 2)
    If strWin2 = "" And Chk打印配药单.Value = 1 And Opt打印配药单选择.Value Then
        MsgBox "选择打印指定窗口的配药单时必须要设置对应的发药窗口！", vbInformation, gstrSysName
        If lst打印窗口.Enabled = True Then lst打印窗口.SetFocus: Exit Sub
    End If
    If UBound(Split(strWin2, ",")) + 1 = lst打印窗口.ListCount Then strWin2 = ""
    
    '来源科室
    mstrSourceDep = ""
    With Me.Lvw来源科室
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With
    
    '处方对应的打印机
    With vsfPrinter
        intTemp = .rows / 6
        For n = 0 To .rows - 1
            If (n + 1) Mod intTemp = 1 And (n + 1) > intTemp Then strPrinters = strPrinters & ";"
            strPrinters = strPrinters & IIf(strPrinters = "" Or Right(strPrinters, 1) = ";", "", ",") & .TextMatrix(n, mintCol格式) & "?" & .TextMatrix(n, mintCol打印机)
        Next
    End With
        
    '两次刷卡的卡类别
    If chk发药刷卡.Value = 1 Then
        If lst卡类型.ListCount > 0 Then
            For i = 0 To lst卡类型.ListCount - 1
                If lst卡类型.Selected(i) Then
                    strCardType = IIf(strCardType = "", strCardType, strCardType & ",") & lst卡类型.ItemData(i)
                End If
            Next
        End If
    End If
        
    On Error Resume Next
    
    '保存私有参数
    zldatabase.SetPara "列设置", Me.cbo药品名称显示.ListIndex, glngSys, 1341

    zldatabase.SetPara "收费处方显示方式", cbo收费处方.ListIndex, glngSys, 1341
    zldatabase.SetPara "记帐处方显示方式", cbo记帐处方.ListIndex, glngSys, 1341
    zldatabase.SetPara "待配药单据打印显示方式", cbo待配药.ListIndex, glngSys, 1341
    zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, 1341
    zldatabase.SetPara "打印包含记帐单", IIf(chk记帐单.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "打印退费单据间隔", Val(Txt打印退费单号), glngSys, 1341
    zldatabase.SetPara "打印延迟", Val(Txt延迟打印), glngSys, 1341
    zldatabase.SetPara "刷新间隔", Val(Txt刷新间隔), glngSys, 1341
    zldatabase.SetPara "打印间隔", Val(txt打印间隔), glngSys, 1341
    
    zldatabase.SetPara "药房属性", cbo单位.ListIndex, glngSys, 1341
    zldatabase.SetPara "显示付数", Chk显示付数栏.Value, glngSys, 1341
    zldatabase.SetPara "发药后自动打印", Me.Cbo发药后.ListIndex, glngSys, 1341
    zldatabase.SetPara "配药后自动打印", Me.cbo配药后.ListIndex, glngSys, 1341
    zldatabase.SetPara "发药后打印药品标签", Me.cbo药品标签.ListIndex, glngSys, 1341
    zldatabase.SetPara "显示大小单位", chk大小单位.Value, glngSys, 1341
    zldatabase.SetPara "配药模式扫描器确认", chk配药扫描.Value, glngSys, 1341
    zldatabase.SetPara "超时未发药品显示时间间隔", IIf(chkOverTime.Value = 0, 0, Int(Val(txtOverTime.Text))), glngSys, 1341
    zldatabase.SetPara "发门诊住院处方", Me.cbo处方类型.ListIndex, glngSys, 1341
    zldatabase.SetPara "两次刷卡发药", strCardType, glngSys, 1341
    zldatabase.SetPara "金额显示方式", cbo金额显示.ListIndex, glngSys, 1341
    zldatabase.SetPara "发药后检查卫材发放情况", chkCheckStuff.Value, glngSys, 1341
    zldatabase.SetPara "待发药单据扫描后自动呼叫", chk扫描后呼叫.Value, glngSys, 1341
    zldatabase.SetPara "打印处方签时先预览再打印", chkPreview.Value, glngSys, 1341
    zldatabase.SetPara "查找时系统自动回车方式", cbo回车方式.ListIndex, glngSys, 1341
    
    If chkDispensing.Visible Then
        zldatabase.SetPara "呼叫时通知开始发药", Me.chkDispensing.Value, glngSys, 1341
    Else
        zldatabase.SetPara "呼叫时通知开始发药", "0", glngSys, 1341
    End If
     
    '打印
    IntPrintStyle = Chk打印配药单.Value
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt打印配药单本部门.Value, 1, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt打印配药单本窗口.Value, 2, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt打印配药单选择.Value, 3, 1)
    zldatabase.SetPara "发现新单据是否打印", IntPrintStyle, glngSys, 1341
    zldatabase.SetPara "打印指定发药窗口", strWin2, glngSys, 1341
    zldatabase.SetPara "打印药品标签", IIf(chk药品标签.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "打印卫材发料单", IIf(chk卫材发料单.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "打印票据的所有格式", IIf(chkAllType.Value = 1, 1, 0), glngSys, 1341
            
    '配药
    zldatabase.SetPara "发药药房", Cbo药房.ItemData(Cbo药房.ListIndex), glngSys, 1341
    zldatabase.SetPara "发药窗口", strWin1, glngSys, 1341
    zldatabase.SetPara "配药人", IIf(cbo配药人.Text <> "当前操作员", cbo配药人.Text, "|当前操作员|"), glngSys, 1341
    zldatabase.SetPara "自动配药", IIf(chk自动配药.Value = 1, 1, 0), glngSys, 1341
    zldatabase.SetPara "自动配药时限", Val(txt配药时限.Text), glngSys, 1341
    zldatabase.SetPara "签到时进行配药", chkSign.Value, glngSys, 1341
    zldatabase.SetPara "核查人", IIf(cboCheck.Text <> "当前操作员", cboCheck.Text, "|当前操作员|"), glngSys, 1341
    zldatabase.SetPara "允许核查人和配药人相同", IIf(chkSame.Value = 1, 1, 0), glngSys, 1341
    
    '保存排队叫号的参数
    zldatabase.SetPara "叫号方式", IIf(Me.optCallWay(0).Value = True, 0, 1), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "启用排队叫号", Me.chk启用排队叫号.Value, glngSys, 1341, mblnSetPara, Cbo药房.ItemData(Cbo药房.ListIndex)
    zldatabase.SetPara "启用语音呼叫", Me.chkUseSound.Value, glngSys, 1341, mblnSetPara
    zldatabase.SetPara "显示排队队列", chkUseDisplay.Value, glngSys, 1341, mblnSetPara
    zldatabase.SetPara "语音播放次数", Val(txtPlayCount.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "语音广播语速", Val(txtSpeed.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "语音类型", IIf(optSoundType(0).Value = True, 0, 1), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "远端呼叫站点", Me.cboWorkStation.Text, glngSys, 1341, mblnSetPara, Cbo药房.ItemData(Cbo药房.ListIndex)
    zldatabase.SetPara "呼叫轮询时间", Val(Me.txtLoopQueryTime.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "显示设备类别", cbo显示硬件类别.ItemData(cbo显示硬件类别.ListIndex), glngSys, 1341, mblnSetPara
    
    
    '来源科室
    zldatabase.SetPara "来源科室", mstrSourceDep, glngSys, 1341
    
    '配药单&处方签打印格式
    zldatabase.SetPara "配药单打印格式", cbo西药配药格式.ItemData(cbo西药配药格式.ListIndex) & ";" & cbo中药配药格式.ItemData(cbo中药配药格式.ListIndex), glngSys, 1341
    zldatabase.SetPara "处方签打印格式", cbo西药处方格式.ItemData(cbo西药处方格式.ListIndex) & ";" & cbo中药处方格式.ItemData(cbo中药处方格式.ListIndex), glngSys, 1341
    
    '处方对应的打印机
    zldatabase.SetPara "处方对应的打印机", strPrinters, glngSys, 1341
    
    frm药品处方发药New.BlnSetParaSuccess = True
    
    '保存配药和配药确认环节
    gstrSQL = "Zl_药房配药控制_Update("
    gstrSQL = gstrSQL & Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex)
    gstrSQL = gstrSQL & "," & Me.chkIsDosage.Value
    gstrSQL = gstrSQL & "," & Me.chkIsDosageOk.Value
    gstrSQL = gstrSQL & ")"
    
    Call zldatabase.ExecuteProcedure(gstrSQL, "cmdOK_Click")
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    
    '初始化chkDispensing
    Call InitDispensing
    
    '读取注册表
    Call ReadFromReg
    '根据设置显示
    Call WriteCons
    '来源科室
    Call SetSourceDep
    
    BlnStartUp = True
    RestoreWinState Me, App.ProductName
End Sub

Private Function ReadWindowsAndPeople()
    Dim intParaType As Integer
    
    '--读取该药房的发药窗口及配药人--
    
    
        '发药窗口（要打印的发药窗口下拉框中不加入"所有发药窗口"）
'        If .State = 1 Then .Close
'        gstrSQL = " Select 名称 From 发药窗口 Where 药房ID=" & Cbo药房.ItemData(Cbo药房.ListIndex)
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        .Open gstrSQL, gcnOracle
'        Call SQLTest

    Dim lngLEDModal As Long
    
    On Error GoTo errHandle
    gstrSQL = " Select 名称 From 发药窗口 Where 药房ID=[1]"
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cbo药房.ItemData(Cbo药房.ListIndex))
    
    mstrWin = ""
    
    With RecPeople
        Me.lst发药窗口.Clear
        Me.lst打印窗口.Clear
        lst发药窗口.Columns = 2 '代码中设效果比较好
        lst打印窗口.Columns = 2
        Do While Not .EOF
            lst发药窗口.AddItem !名称
            lst打印窗口.AddItem !名称
            
            lst发药窗口.Selected(lst发药窗口.NewIndex) = True
            If Opt打印配药单选择.Value Then
                lst打印窗口.Selected(lst打印窗口.NewIndex) = True
            End If
            
            mstrWin = IIf(mstrWin = "", "", mstrWin & ",") & !名称
            
            .MoveNext
        Loop

        If lst发药窗口.ListCount > 0 Then lst发药窗口.ListIndex = 0
        If lst打印窗口.ListCount > 0 Then lst打印窗口.ListIndex = 0
    End With
    '配药人
    gstrSQL = " Select 姓名 From 人员表  Where ID in " & _
             " (Select Distinct 人员ID From 人员性质说明 Where 人员性质='药房发药人' " & _
             " And 人员ID IN (Select 人员ID From 部门人员 Where 部门ID=[1]))" & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cbo药房.ItemData(Cbo药房.ListIndex))
    
    With RecPeople
        Me.cbo配药人.Clear
        Me.cbo配药人.AddItem "当前操作员"
        Do While Not .EOF
            cbo配药人.AddItem !姓名
            .MoveNext
        Loop
        cbo配药人.ListIndex = 0
    End With
    
    With RecPeople
        If .RecordCount <> 0 Then
            .MoveFirst
        End If
        Me.cboCheck.Clear
        Me.cboCheck.AddItem "当前操作员"
        Do While Not .EOF
            cboCheck.AddItem !姓名
            .MoveNext
        Loop
        cboCheck.ListIndex = 0
    End With
    
    lngLEDModal = zldatabase.GetPara("显示设备类别", glngSys, 1341, "101", Null, mblnSetPara, intParaType, lng药房ID)
    cbo显示硬件类别.Clear
    
    gstrSQL = "Select 部件类型,部件名,Nvl(启用,0) AS 启用,说明 From 排队LED显示部件  "
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, "提取该LED显示接口的注册信息")
    
    While RecPeople.EOF = False
        cbo显示硬件类别.AddItem zlStr.nvl(RecPeople!说明)
        cbo显示硬件类别.ItemData(cbo显示硬件类别.ListCount - 1) = zlStr.nvl(RecPeople!部件类型, 0)
        If lngLEDModal = zlStr.nvl(RecPeople!部件类型, 0) Then
            cbo显示硬件类别.ListIndex = cbo显示硬件类别.ListCount - 1
        End If
        RecPeople.MoveNext
    Wend
    
    If cbo显示硬件类别.ListCount > 0 And cbo显示硬件类别.ListIndex = -1 Then
        cbo显示硬件类别.ListIndex = 0
    End If
    
    '添加站点列表
    ReadWorkStationInf
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteCons()
    Dim IntLocate As Integer
    Dim rsData As ADODB.Recordset
    
    '根据用户设置显示
    
    RecPart.MoveFirst               '不可能为空（否则连主界面都进不了）
    
    txt查询天数.Text = intDays
    '装入下拉框数据
    With Me.Cbo药房
        Do While Not RecPart.EOF
            .AddItem RecPart!名称
            .ItemData(.NewIndex) = RecPart!Id
            RecPart.MoveNext
        Loop
        .ListIndex = 0
    End With
    With Me.Cbo发药后
        .AddItem "1-发药后提示是否打印"
        .AddItem "2-发药后自动打印"
        .AddItem "3-发药后不打印"
        .ListIndex = IntAutoPrint
    End With
    
    With Me.cbo配药后
        .AddItem "1-配药后提示是否打印"
        .AddItem "2-配药后自动打印"
        .AddItem "3-配药后不打印"
        .ListIndex = mint配药后自动打印
    End With
    
    With Me.cbo药品标签
        .AddItem "1-发药后提示是否打印"
        .AddItem "2-发药后自动打印"
        .AddItem "3-发药后不打印"
        .ListIndex = mint发药后自动打印药品标签
    End With
    
    With cbo收费处方
        .Clear
        .AddItem "1-不显示任何处方"
        .AddItem "2-显示未收费处方"
        .AddItem "3-显示已收费处方"
        .AddItem "4-显示所有的处方"
        .ListIndex = 0
    End With
    With cbo记帐处方
        .Clear
        .AddItem "1-不显示任何处方"
        .AddItem "2-显示未审核处方"
        .AddItem "3-显示已审核处方"
        .AddItem "4-显示所有的处方"
        .ListIndex = 0
    End With
    
    With cbo待配药
        .Clear
        .AddItem "0-显示所有配药单"
        .AddItem "1-显示未打印配药单"
        .AddItem "2-显示已打印配药单"
        .ListIndex = 0
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-西药处方签"
        .AddItem "2-中药处方签"
        .AddItem "3-处方发药清单"
        .AddItem "4-处方退药通知单"
        .AddItem "5-记帐处方统计表"
        .AddItem "6_西药药品标签"
        .AddItem "7_中草药药品标签"
        .AddItem "8_已退费单据"
        .ListIndex = 0
    End With
    
    With Me.cbo药品名称显示
        .Clear
        .AddItem "0-显示药品编码与名称"
        .AddItem "1-仅显示药品编码"
        .AddItem "2-仅显示药品名称"
        .ListIndex = 0
    End With
    
    With Me.cbo金额显示
        .Clear
        .AddItem "0-显示应收金额"
        .AddItem "1-显示实收金额"
        .AddItem "2-显示应收和实收金额"
        .ListIndex = 0
    End With
    
    With Me.cbo处方类型
        .Clear
        .AddItem "0-显示门诊和住院处方"
        .AddItem "1-只显示门诊处方"
        .AddItem "2-只显示住院处方"
        .ListIndex = mintType
    End With
    
    With Me.cbo回车方式
        .Clear
        .AddItem "0-系统不自动回车"
        .AddItem "1-当录入达到项目或卡号长度时自动回车"
    End With
    
    '装入基本数据
    cbo收费处方.ListIndex = mintShowBill收费
    cbo记帐处方.ListIndex = mintShowBill记帐
    cbo待配药.ListIndex = mintShowBill配药
    Chk显示付数栏.Value = IntShowCol
    
    Chk打印配药单.Value = IIf(intPrint = 0, 0, 1)
    
    cbo金额显示.ListIndex = mint金额显示方式

    Opt打印配药单本部门.Value = IIf(intPrint = 1, True, False)
    Opt打印配药单本窗口.Value = IIf(intPrint = 2, True, False)
    Opt打印配药单选择.Value = IIf(intPrint = 3, True, False)
    
    Txt刷新间隔 = Format(IntRefresh, "#####;-#####; ;")
    txt打印间隔 = Format(mintPrintInterval, "#####;-#####; ;")
    Txt延迟打印 = Format(intPrintDelay, "#####;-#####; ;")
    Txt打印退费单号 = Format(intPrintHandbackNO, "#####;-#####; ;")
    
    If txt打印间隔.Enabled Then txt打印间隔.Enabled = Not mblnUseMsg
    lblPrintComment.Visible = mblnUseMsg
    
    If Txt刷新间隔.Enabled Then Txt刷新间隔.Enabled = Not mblnUseMsg And chkIsDosage.Value = 0
    lblRefreshComment.Visible = mblnUseMsg
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "已启用消息机制替代", "待配药环节已启用消息机制替代自动刷新")
    
    If lng药房ID <> 0 Then                                  '定位药房
        '不存在该药房则提示
        For IntLocate = 0 To Me.Cbo药房.ListCount - 1
            If Me.Cbo药房.ItemData(IntLocate) = lng药房ID Then
                Me.Cbo药房.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cbo药房.ListCount - 1) Then
            MsgBox "请重新设置药房（原来设置的药房已失效）！", vbInformation, gstrSysName
            If Cbo药房.ListCount >= 1 Then Cbo药房.ListIndex = 0
        End If
    End If
    BlnStartUp = True
    Cbo药房_Click                                           '不管设置药房否，均提取该药房所含发药窗口及配药人
    BlnStartUp = False
    
    '定位发药窗口
    If Str窗口 <> "" Then
        For IntLocate = 0 To lst发药窗口.ListCount - 1
            If InStr(Str窗口, "'" & lst发药窗口.List(IntLocate) & "'") > 0 Then
                lst发药窗口.Selected(IntLocate) = True
            Else
                lst发药窗口.Selected(IntLocate) = False
            End If
        Next
        If lst发药窗口.ListCount > 0 Then lst发药窗口.ListIndex = 0
    End If
    
    If str配药人 <> "" Then                                 '显示
        '不存在该配药人则提示
        If str配药人 = "|当前操作员|" Then
            cbo配药人.ListIndex = 0
        Else
            For IntLocate = 1 To cbo配药人.ListCount - 1
                If cbo配药人.List(IntLocate) = str配药人 Then
                    cbo配药人.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (cbo配药人.ListCount - 1) Then
                MsgBox "请重新设置配药人（原来设置的配药人已不在本部门）！", vbInformation, gstrSysName
                If cbo配药人.ListCount >= 1 Then cbo配药人.ListIndex = 0
            End If
        End If
    End If
    
    If mstr核查人 <> "" Then
        '不存在该核查人则提示
        If mstr核查人 = "|当前操作员|" Then
            cboCheck.ListIndex = 0
        Else
            For IntLocate = 1 To cboCheck.ListCount - 1
                If cboCheck.List(IntLocate) = mstr核查人 Then
                    cboCheck.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (cboCheck.ListCount - 1) Then
                MsgBox "请重新设置核查人（原来设置的核查人已不在本部门）！", vbInformation, gstrSysName
                If cboCheck.ListCount >= 1 Then cboCheck.ListIndex = 0
            End If
        End If
    End If
    
    '定位打印发药窗口
    If strPrintWindow <> "" Then
        For IntLocate = 0 To lst打印窗口.ListCount - 1
            If InStr(strPrintWindow, "'" & lst打印窗口.List(IntLocate) & "'") > 0 Then
                lst打印窗口.Selected(IntLocate) = True
            Else
                lst打印窗口.Selected(IntLocate) = False
            End If
        Next
        If lst打印窗口.ListCount > 0 Then lst打印窗口.ListIndex = 0
    End If
    
    Me.cbo药品名称显示.ListIndex = mintShowName
    
    chk自动配药.Value = IIf(mint自动配药 = 1, 1, 0)
    chk记帐单.Value = IIf(mint记帐单 = 1, 1, 0)
    chk药品标签.Value = IIf(Chk打印配药单.Value = 1, IIf(mint药品标签 = 1, 1, 0), 0)
    chk卫材发料单.Value = IIf(Chk打印配药单.Value = 1, IIf(mint卫材发料单 = 1, 1, 0), 0)
    txt配药时限.Text = mint自动配药时限
    txt配药时限.Enabled = (mint自动配药 = 1 And chk自动配药.Enabled = True)
    chk配药扫描.Value = IIf(mint配药扫描 = 1, 1, 0)
    chkSign.Value = IIf(mintSign = 1, 1, 0)
    Me.chk扫描后呼叫.Value = IIf(mint扫描后呼叫 = 1, 1, 0)
    
    If mint回车方式 >= 0 And mint回车方式 <= 1 Then
        cbo回车方式.ListIndex = mint回车方式
    Else
        cbo回车方式.ListIndex = 0
    End If
    
    '设置排队叫号的参数
    With mType_Call
        chk启用排队叫号.Value = .int启用排队叫号
        chkUseDisplay.Value = .int显示排队队列
        chkUseSound.Value = .int启用语音呼叫
        
        If .int叫号方式 = 0 Then
            optCallWay(0).Value = True
        Else
            optCallWay(1).Value = True
        End If
        
        optSoundType(.int语音类型).Value = 1
        txtSpeed.Text = .int语音广播语速
        txtPlayCount.Text = .int语音播放次数
        Me.cboWorkStation.Text = .str远端呼叫站点
        txtLoopQueryTime.Text = .int轮询时间
    End With
    
    chkUseDisplay_Click
    chkUseSound_Click
    
    If Me.optCallWay(0).Value = True Then
        optCallWay_Click 0
    Else
        optCallWay_Click 1
    End If
    
    '两次刷卡模式和卡类别
    chk发药刷卡.Value = IIf(mstr两次刷卡发药 = "", 0, 1)
    lst卡类型.Enabled = (chk发药刷卡.Value = 1)
    
    gstrSQL = "Select ID, 编码, 名称 From 医疗卡类别 Order By 编码"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "WriteCons")
    If rsData.RecordCount > 0 Then
        lst卡类型.Clear
        lst卡类型.Columns = 2
        Do While Not rsData.EOF
            lst卡类型.AddItem rsData!名称
            lst卡类型.ItemData(lst卡类型.NewIndex) = rsData!Id
            
            If mstr两次刷卡发药 <> "" Then
                If InStr(1, "," & mstr两次刷卡发药 & ",", "," & rsData!Id & ",") > 0 Then
                    lst卡类型.Selected(lst卡类型.NewIndex) = True
                End If
            End If
            
            rsData.MoveNext
        Loop

        If lst卡类型.ListCount > 0 Then lst卡类型.ListIndex = 0
    Else
        chk发药刷卡.Enabled = False
        lst卡类型.Enabled = False
    End If
    chkCheckStuff.Value = IIf(mint发药后检查 = 1, 1, 0)
 End Function

Private Sub optCallWay_Click(index As Integer)
    If index = 0 Then
        Fra远端语音设置.Enabled = False
        frm语音广播设置.Enabled = True
    Else
        Fra远端语音设置.Enabled = True
        frm语音广播设置.Enabled = False
    End If
End Sub

Private Sub optSoundType_Click(index As Integer)
    If optSoundType(0).Value = True Then
        Label10.Caption = "语音语速：      (范围在0到100之间，推荐65)"
        txtSpeed.Text = "65"
    Else
        Label10.Caption = "语音语速：      (范围在-10到10之间，推荐-4)"
        txtSpeed.Text = "-4"
    End If
End Sub

Private Sub Opt打印配药单本部门_Click()
    lst打印窗口.Enabled = False
End Sub

Private Sub Opt打印配药单本部门_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Opt打印配药单本窗口_Click()
    lst打印窗口.Enabled = False
End Sub

Private Sub Opt打印配药单本窗口_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Opt打印配药单选择_Click()
    lst打印窗口.Enabled = Opt打印配药单选择.Enabled
    If BlnStartUp = False Then Exit Sub
    
    If Opt打印配药单选择.Value Then
        If lst打印窗口.Enabled = True Then lst打印窗口.SetFocus
    End If
End Sub
Private Sub Opt打印配药单选择_GotFocus()
    TabShow.Tab = 2
End Sub



Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        If Me.Cbo药房.Enabled = True Then Me.Cbo药房.SetFocus
    Case 2
        If Me.Chk打印配药单.Enabled = True Then Me.Chk打印配药单.SetFocus
    Case 3
        If Me.cbo票据设置.Enabled = True Then Me.cbo票据设置.SetFocus
    End Select
End Sub

Private Sub txtOverTime_Change()
    txtOverTime.Text = Int(Val(txtOverTime.Text))
    If Val(txtOverTime.Text) > 1440 Then
        txtOverTime.Text = "1440"
    End If
End Sub

Private Sub txtOverTime_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt打印间隔_GotFocus()
    GetFocus txt打印间隔
End Sub


Private Sub txt配药时限_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub


Private Sub Txt打印退费单号_GotFocus()
    GetFocus Txt打印退费单号
End Sub

Private Sub Txt刷新间隔_GotFocus()
    GetFocus Txt刷新间隔
End Sub

Private Sub Txt延迟打印_GotFocus()
    GetFocus Txt延迟打印
End Sub

Private Sub SetDispense()
'--------------------------------------
'设置配药控制的相关参数
'--------------------------------------
    Dim bln配药确认 As Boolean
    
    Me.chkIsDosage.Value = IIf(RecipeSendWork_DispensingMedi(Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex), bln配药确认) = True, 1, 0)
    
    Me.chkIsDosageOk.Value = IIf(bln配药确认 = True, 1, 0)
End Sub

Private Sub ReadWorkStationInf()
'*****************************************************
'读取站点信息
'*****************************************************

    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select 工作站 from zlClients where 禁止使用<>1 order by 工作站"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "读取站点信息")
    
    If rstemp.EOF Then Exit Sub
    
    While Not rstemp.EOF
        Call cboWorkStation.AddItem(rstemp("工作站"))
        rstemp.MoveNext
    Wend
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function NOCheck() As Boolean
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select 1 from 未发药品记录 where 库房id=[1] and (单据=8 or 单据=9 or 单据=10)"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "NOCheck", Val(Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex)))
    
    If rstemp.EOF Then
        NOCheck = True
    Else
        NOCheck = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDispensing()
'功能：初始化chkDispensing控件

    Dim objMachine As Object
    
    err.Clear
    On Error Resume Next
    If Val(zldatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set objMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '其次旧接口
            Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        '旧接口
        Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    On Error GoTo 0
    
    If objMachine Is Nothing Then
        '药品自动化设备接口不存在
        chkDispensing.Visible = False
        chkDispensing.Value = 0
    Else
        '药品自动化设备接口存在
        chkDispensing.Visible = True
        chkDispensing.Value = Val(zldatabase.GetPara("呼叫时通知开始发药", glngSys, 1341))
    End If
    
    lst卡类型.Height = cbo回车方式.Top - lst卡类型.Top - 180
End Sub
