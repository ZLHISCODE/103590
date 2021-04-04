VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统选项"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7110
      Index           =   5
      Left            =   195
      TabIndex        =   74
      Top             =   540
      Width           =   6990
      Begin VSFlex8Ctl.VSFlexGrid vsfLogList 
         Height          =   5295
         Left            =   300
         TabIndex        =   78
         Top             =   1305
         Width           =   6495
         _cx             =   11456
         _cy             =   9340
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOptions.frx":000C
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
      Begin VB.Label Label5 
         Caption         =   "保存天数为0时表示不清理，最大保存天数为100天"
         Height          =   360
         Left            =   300
         TabIndex        =   82
         Top             =   6735
         Width           =   5520
      End
      Begin VB.Label lblLogPath 
         Height          =   345
         Left            =   1485
         TabIndex        =   77
         Top             =   915
         Width           =   3870
      End
      Begin VB.Label lblLog 
         AutoSize        =   -1  'True
         Caption         =   "日志存储路径"
         Height          =   180
         Left            =   300
         TabIndex        =   76
         Top             =   900
         Width           =   1080
      End
      Begin VB.Image Image3 
         Height          =   510
         Left            =   225
         Picture         =   "frmOptions.frx":00A1
         Stretch         =   -1  'True
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "    选择是否启动跟踪日志以及设置跟踪日志的保留天数。"
         Height          =   480
         Left            =   840
         TabIndex        =   75
         Top             =   210
         Width           =   5685
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "1"
      Height          =   7125
      Index           =   3
      Left            =   180
      TabIndex        =   34
      Top             =   540
      Width           =   6885
      Begin VB.TextBox txtRemote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   81
         Top             =   6790
         Width           =   705
      End
      Begin VB.CheckBox chkAutoHide 
         Caption         =   "允许界面区域提供自动隐藏功能(&I)"
         Height          =   195
         Left            =   1860
         TabIndex        =   79
         Top             =   3735
         Width           =   3090
      End
      Begin VB.TextBox txtAutoLock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2610
         MaxLength       =   3
         TabIndex        =   73
         Text            =   "5"
         ToolTipText     =   "锁屏间隔可输入0-999"
         Top             =   6510
         Width           =   300
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   8
         Left            =   1590
         TabIndex        =   71
         Top             =   5955
         Width           =   5250
      End
      Begin VB.CheckBox chkLanJoin 
         Caption         =   "网络断开、恢复后自动重新连接服务器"
         Height          =   210
         Left            =   1860
         TabIndex        =   69
         Top             =   6195
         Width           =   3465
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   5
         Left            =   1605
         TabIndex        =   56
         Top             =   3405
         Width           =   5235
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   5
         Left            =   960
         Picture         =   "frmOptions.frx":18993
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox pic匹配方式 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   3495
         TabIndex        =   53
         Top             =   930
         Width           =   3495
         Begin VB.OptionButton opt 
            Caption         =   "双向匹配(&D)"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   60
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton opt 
            Caption         =   "从左匹配(&L)"
            Height          =   210
            Index           =   1
            Left            =   1815
            TabIndex        =   13
            Top             =   60
            Width           =   1320
         End
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   4
         Left            =   945
         Picture         =   "frmOptions.frx":1925D
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5310
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   3
         Left            =   960
         Picture         =   "frmOptions.frx":1BC4F
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4425
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   2
         Left            =   975
         Picture         =   "frmOptions.frx":1D5D1
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2715
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   1005
         Picture         =   "frmOptions.frx":1EF53
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1755
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   1005
         Picture         =   "frmOptions.frx":208D5
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   885
         Width           =   480
      End
      Begin VB.PictureBox pic2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Left            =   255
         Picture         =   "frmOptions.frx":22257
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "60"
         ToolTipText     =   "当检查周期设为0或空的时候，表示不检查。检查周期设置范围（10-300）"
         Top             =   5475
         Width           =   540
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   4
         Left            =   1590
         TabIndex        =   40
         Top             =   5100
         Width           =   5250
      End
      Begin VB.CheckBox chkShutDown 
         Caption         =   "退出程序时自动关闭 Windows (&S)"
         Height          =   210
         Left            =   1860
         TabIndex        =   19
         Top             =   4725
         Width           =   3045
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "在 Windows 启动时自动运行(&A)"
         Height          =   210
         Left            =   1860
         TabIndex        =   18
         Top             =   4515
         Width           =   2865
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   3
         Left            =   1590
         TabIndex        =   39
         Top             =   4245
         Width           =   5250
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   2
         Left            =   1590
         TabIndex        =   38
         Top             =   2430
         Width           =   5250
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   1
         Left            =   1590
         TabIndex        =   37
         Top             =   1530
         Width           =   5250
      End
      Begin VB.ComboBox cmbIME 
         Height          =   300
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1890
         Width           =   3135
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   0
         Left            =   1590
         TabIndex        =   36
         Top             =   645
         Width           =   5250
      End
      Begin VB.PictureBox pic简码 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   870
         Left            =   1830
         ScaleHeight     =   870
         ScaleWidth      =   3345
         TabIndex        =   35
         Top             =   2520
         Width           =   3345
         Begin VB.CheckBox chkIMETurn 
            Caption         =   "允许在窗口界面的工具栏切换简码方式 "
            Height          =   255
            Left            =   15
            TabIndex        =   17
            Top             =   560
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.OptionButton opt简码 
            Caption         =   "五笔，取每字的首字母构成简码(&W)"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   16
            Top             =   330
            Width           =   3150
         End
         Begin VB.OptionButton opt简码 
            Caption         =   "拼音，取每字的首字母构成简码(&P)"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   15
            Top             =   90
            Value           =   -1  'True
            Width           =   3150
         End
      End
      Begin VB.CheckBox chkAutoLock 
         Caption         =   "等待     分钟无操作后自动锁定系统"
         Height          =   210
         Left            =   1860
         TabIndex        =   72
         Top             =   6495
         Width           =   3345
      End
      Begin VB.Label lblRemote 
         AutoSize        =   -1  'True
         Caption         =   "远程连接监听端口________，值为-1时不允许远程连接"
         Height          =   180
         Left            =   1860
         TabIndex        =   83
         Top             =   6840
         Width           =   4320
      End
      Begin VB.Line lineTime 
         X1              =   2040
         X2              =   2655
         Y1              =   5670
         Y2              =   5670
      End
      Begin VB.Label lblTime 
         Caption         =   "每       秒检查消息通知"
         Height          =   255
         Left            =   1845
         TabIndex        =   80
         Top             =   5490
         Width           =   2145
      End
      Begin VB.Line Line1 
         X1              =   2490
         X2              =   2930
         Y1              =   6705
         Y2              =   6705
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他设置"
         Height          =   180
         Index           =   8
         Left            =   795
         TabIndex        =   70
         Top             =   5955
         Width           =   720
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "界面区域隐藏"
         Height          =   180
         Index           =   5
         Left            =   435
         TabIndex        =   57
         Top             =   3390
         Width           =   1080
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "消息通知检查周期"
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   46
         Top             =   5100
         Width           =   1440
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "与Windows的结合"
         Height          =   180
         Index           =   3
         Left            =   165
         TabIndex        =   45
         Top             =   4245
         Width           =   1350
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "简码处理方式"
         Height          =   180
         Index           =   2
         Left            =   435
         TabIndex        =   44
         Top             =   2430
         Width           =   1080
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "汉字输入法"
         Height          =   180
         Index           =   1
         Left            =   615
         TabIndex        =   43
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目输入匹配方式"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   42
         Top             =   630
         Width           =   1440
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":22F21
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "    用户可以根据自身的习惯来选择输入的匹配方式、简码类型、输入法等，以提高工作效率"
         Height          =   480
         Left            =   870
         TabIndex        =   41
         Top             =   120
         Width           =   6105
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   390
      TabIndex        =   23
      Top             =   7800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6075
      TabIndex        =   22
      Top             =   7800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4755
      TabIndex        =   21
      Top             =   7800
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   7635
      Left            =   90
      TabIndex        =   24
      Top             =   120
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   13467
      TabWidthStyle   =   2
      TabMinWidth     =   989
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "导航风格"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "菜单选择"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "使用习惯"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品名称"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "日志管理"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6975
      Index           =   1
      Left            =   180
      TabIndex        =   26
      Top             =   540
      Width           =   7095
      Begin VB.PictureBox pic导航风格 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1080
         ScaleHeight     =   420
         ScaleWidth      =   5340
         TabIndex        =   54
         Top             =   960
         Width           =   5340
         Begin VB.OptionButton optStyle 
            Caption         =   "Windows风格"
            Height          =   195
            Index           =   1
            Left            =   2053
            TabIndex        =   1
            Tag             =   "zlwin"
            Top             =   60
            Width           =   1428
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "传统风格"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   0
            Tag             =   "zlBrw"
            Top             =   60
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "MDI风格"
            Height          =   195
            Index           =   2
            Left            =   4245
            TabIndex        =   2
            Tag             =   "zlmdi"
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.PictureBox picPreview 
         Height          =   3795
         Left            =   1320
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   292
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2265
         Width           =   4440
      End
      Begin VB.PictureBox picBorder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   960
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   33
         Top             =   1905
         Width           =   5115
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":23536
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "    按照个人的兴趣爱好选择自己喜欢的导航风格，使你心情更愉快、工作更轻松。"
         Height          =   480
         Left            =   870
         TabIndex        =   25
         Top             =   120
         Width           =   5835
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6750
      Index           =   2
      Left            =   180
      TabIndex        =   27
      Top             =   540
      Width           =   6960
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   3
         Left            =   5985
         Picture         =   "frmOptions.frx":23978
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "向后移动"
         Top             =   6345
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   2
         Left            =   5985
         Picture         =   "frmOptions.frx":23AC5
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "向前移动"
         Top             =   5865
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   1
         Left            =   5985
         Picture         =   "frmOptions.frx":23C12
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "删除程序"
         Top             =   3915
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   0
         Left            =   5985
         Picture         =   "frmOptions.frx":23CB8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "新增程序"
         Top             =   3435
         Width           =   345
      End
      Begin MSComctlLib.ListView lvwFavorite 
         Height          =   3870
         Left            =   3735
         TabIndex        =   7
         Top             =   2820
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   6826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "程序"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "生成快捷方式"
         Height          =   1335
         Left            =   3765
         TabIndex        =   28
         Top             =   855
         Width           =   2625
         Begin VB.CommandButton cmdStartup 
            Caption         =   "到启动菜单(&S)"
            Height          =   350
            Left            =   630
            TabIndex        =   6
            Top             =   840
            Width           =   1725
         End
         Begin VB.CommandButton cmdDesktop 
            Caption         =   "到桌面(&D)"
            Height          =   350
            Left            =   630
            TabIndex        =   5
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   870
         Width           =   1620
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   2160
         Top             =   2910
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tvwMain 
         Height          =   5325
         Left            =   705
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   9393
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin VB.Label lblFavorite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "常用程序列表(&R)"
         Height          =   180
         Left            =   3705
         TabIndex        =   32
         Top             =   2445
         Width           =   1350
      End
      Begin VB.Label lblNote 
         Caption         =   $"frmOptions.frx":23D65
         Height          =   570
         Left            =   900
         TabIndex        =   30
         Top             =   30
         Width           =   4740
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "菜单体系"
         Height          =   180
         Left            =   720
         TabIndex        =   29
         Top             =   930
         Width           =   720
      End
      Begin VB.Image imgNote 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":23DED
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   7035
      Index           =   4
      Left            =   240
      TabIndex        =   58
      Top             =   540
      Width           =   6960
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   8
         Left            =   390
         Picture         =   "frmOptions.frx":23FF7
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.ComboBox cbo输入药品显示 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   3195
         Width           =   3015
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   7
         Left            =   1200
         TabIndex        =   65
         Top             =   2475
         Width           =   5610
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   7
         Left            =   360
         Picture         =   "frmOptions.frx":24CC1
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3075
         Width           =   480
      End
      Begin VB.ComboBox cbo药品名称显示 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1875
         Width           =   3015
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   6
         Left            =   1200
         TabIndex        =   60
         Top             =   1035
         Width           =   5610
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   6
         Left            =   390
         Picture         =   "frmOptions.frx":2598B
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1755
         Width           =   480
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入药品显示（通过输入简码方式进入选择器时药品名称的显示）"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   67
         Top             =   2715
         Width           =   5220
      End
      Begin VB.Label lblMedi 
         Caption         =   "    用户可以根据自身的习惯来选择药品名称的显示方式，支持显示通用名和商品名。"
         Height          =   585
         Left            =   1065
         TabIndex        =   63
         Top             =   240
         Width           =   5850
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSplit 
         BackStyle       =   0  'Transparent
         Caption         =   "药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）"
         Height          =   420
         Index           =   6
         Left            =   120
         TabIndex        =   62
         Top             =   1275
         Width           =   6690
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum constOpt
    opt_双向匹配 = 0
    opt_从左匹配 = 1
End Enum

Private Enum LogList
    LL_启用 = 0
    LL_系统名称 = 1
    LL_日志名称 = 2
    LL_保存天数 = 3
    LL_说明 = 4
    LL_最多保存天数 = 100
End Enum

Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function OSfRemoveShellLink Lib "vb6stkit.dll" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Private mintIndex As Integer

Private Sub chkAutoLock_Click()
        txtAutoLock.Enabled = chkAutoLock.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDesktop_Click()
    Dim strPath As String
    Dim StrName As String
    
    '创建快捷方式
    strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS+.exe")
    StrName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    
'    '到桌面的
    If Not OSfCreateShellLink("..\DeskTop", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
        If Not OSfCreateShellLink("..\桌面", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
            'Win7下路径
            If Not OSfCreateShellLink("..\..\..\..\..\DeskTop", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
                Call OSfCreateShellLink("..\..\..\..\..\桌面", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)")
            End If
        End If
    End If
End Sub

Private Sub cmdStartup_Click()
    Dim strPath As String
    Dim StrName As String
    
    '创建快捷方式
    strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS+.exe")
    StrName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")

    '到启动菜单
    If Not OSfCreateShellLink("\Startup", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)") Then
        Call OSfCreateShellLink("\启动", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)")
    End If
End Sub

Private Sub cmdFavorite_Click(Index As Integer)
    Dim lst As ListItem, lngIndex As Long
    
    If Index <> 0 And lvwFavorite.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Index
        Case 0 '新增
            If tvwMain.SelectedItem Is Nothing Then Exit Sub
            
            With tvwMain.SelectedItem
                If lvwFavorite.ListItems.Count >= 10 Then
                    MsgBox "最多只能添加10个常用模块！", vbInformation, gstrSysName
                    Exit Sub
                End If
                For Each lst In lvwFavorite.ListItems
                    If lst.Tag = .Tag Then
                        MsgBox "“" & .Text & "”已经存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
                If Right(.Tag, 1) = "_" Then Exit Sub
                
                Set lst = lvwFavorite.ListItems.Add(, , .Text, .Image, .Image)
                lst.Tag = .Tag
                lst.Selected = True
                lst.EnsureVisible
            End With
        Case 1 '删除
            lngIndex = lvwFavorite.SelectedItem.Index
            lvwFavorite.ListItems.Remove lngIndex
            
            If lvwFavorite.ListItems.Count = 0 Then Exit Sub
            If lngIndex > lvwFavorite.ListItems.Count Then
                lvwFavorite.ListItems.item(lngIndex - 1).Selected = True
            Else
                lvwFavorite.ListItems.item(lngIndex).Selected = True
            End If
        Case 2 '前移
            With lvwFavorite.SelectedItem
                If .Index = 1 Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index - 1, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
        Case 3 '后移
            With lvwFavorite.SelectedItem
                If .Index = lvwFavorite.ListItems.Count Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index + 2, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
    End Select
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim lst As ListItem, i As Integer
    Dim str系统 As String, str序号 As String, str图标 As String, str标题 As String
    Dim strOldStyle As String
    strOldStyle = zlDatabase.GetPara("导航台")
    '保存导航风格
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            Call zlDatabase.SetPara("导航台", optStyle(i).Tag)
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "导航台", optStyle(i).Tag
            Exit For
        End If
    Next
    
    '保存常用模块
    For Each lst In lvwFavorite.ListItems
        str系统 = str系统 & "," & Val(Mid(lst.Tag, 1, InStr(lst.Tag, "_") - 1))
        str序号 = str序号 & "," & Mid(lst.Tag, InStr(lst.Tag, "_") + 1)
        str图标 = str图标 & "," & Mid(lst.Icon, 2)
        str标题 = str标题 & "," & lst.Text
    Next
    If str系统 <> "" Then
        str系统 = Mid(str系统, 2)
        str序号 = Mid(str序号, 2)
        str图标 = Mid(str图标, 2)
        str标题 = Mid(str标题, 2)
    End If
    Call zlDatabase.SetPara("常用功能模块", str系统 & "|" & str序号 & "|" & str图标 & "|" & str标题)
    
    Call SaveRegister
    
    '保持药品名称显示方式
    Call zlDatabase.SetPara("药品名称显示", cbo药品名称显示.ListIndex)
    Call zlDatabase.SetPara("输入药品显示", cbo输入药品显示.ListIndex)
    
    '写入zl9Comlib.glngAutoConnect
    zl9ComLib.gblnAutoConnect = IIf(chkLanJoin.Value = 1, True, False)
    If UCase(strOldStyle) <> UCase(zlDatabase.GetPara("导航台")) Then
        If gclsAppTool Is Nothing Then
            MsgBox "新的风格需要重新启动导航台才生效。如果要立即体验新风格请重启导航台！", vbInformation, gstrSysName
        Else
            If MsgBox("新的风格需要重新启动导航台才生效。是否立即重启导航台？", vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                gclsAppTool.IsRestart = True
            End If
        End If
    End If
    
    '将日志配置信息保存到注册表
    For i = 1 To vsfLogList.Rows - 1
        If vsfLogList.Cell(flexcpChecked, i, LL_启用) = 1 Then
            SaveSetting "ZLSOFT", "公共全局\日志\" & vsfLogList.TextMatrix(i, LL_日志名称), "保存天数", vsfLogList.TextMatrix(i, LL_保存天数)
            SaveSetting "ZLSOFT", "公共全局\日志\" & vsfLogList.TextMatrix(i, LL_日志名称), "启用", 1
        Else
            On Error Resume Next
            DeleteSetting "ZLSOFT", "公共全局\日志\" & vsfLogList.TextMatrix(i, LL_日志名称)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next
    Call LogInit(True)
    Unload Me
End Sub

Private Sub Form_Activate()
    Call tabMain_Click
End Sub

Private Sub Form_Load()
    Dim intIndex As Integer
    Dim strStyle As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    Call InitIcon
    Call FillCommon
    
    gstrSQL = "select 组别 from ZLMENUS group by  组别 "
    
    rsTemp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cboGroup.AddItem rsTemp("组别")
        If rsTemp("组别") = gstrMenuSys Then intIndex = cboGroup.NewIndex
        rsTemp.MoveNext
    Loop
    mintIndex = -3
    cboGroup.ListIndex = intIndex
    
    strStyle = UCase(zlDatabase.GetPara("导航台"))
    
    For intIndex = optStyle.LBound To optStyle.UBound
        If UCase(optStyle(intIndex).Tag) = strStyle Then
            optStyle(intIndex).Value = True
            Exit For
        End If
    Next
    Call optStyle_Click(0)

    Call LoadCustom
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitIcon()
'功能：把图标装入到控件中
    Dim i As Long
    ils16.ListImages.Clear
    ils16.ImageWidth = 16
    ils16.ImageHeight = 16
    
    With ils16.ListImages
        For i = glngLBound To glngUBound
            .Add , "K" & i, LoadResPicture(i, vbResIcon)
        Next
    End With
End Sub

Private Sub FillTree(ByVal str组别 As String)
    Dim strSQL As String
    Dim strTemp As String
    Dim rsMenus As New ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "select * " & _
            " from zlMenus" & _
            " start with 上级ID is null and 组别=[1] " & _
            " connect by prior ID =上级ID  and 组别=[1] order by level,ID"
    rsMenus.CursorLocation = adUseClient
    Set rsMenus = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str组别)
    With rsMenus
        tvwMain.Nodes.Clear
        Do While Not .EOF
            If IsNull(rsMenus("图标")) Or rsMenus("图标") = 0 Then
                strTemp = IIf(IsNull(.Fields("模块").Value), "K99", "K100")
            Else
                strTemp = "K" & rsMenus("图标")
            End If
            If IsNull(.Fields("上级ID")) Then
                tvwMain.Nodes.Add , tvwChild, "C" & .Fields("ID").Value, .Fields("标题").Value, strTemp, strTemp
            Else
               tvwMain.Nodes.Add "C" & .Fields("上级ID").Value, tvwChild, "C" & .Fields("ID").Value, .Fields("标题").Value, strTemp, strTemp
            End If
            tvwMain.Nodes("C" & .Fields("ID").Value).Tag = IIf(IsNull(.Fields("系统")), "", .Fields("系统")) & "_" & IIf(IsNull(.Fields("模块")), "", .Fields("模块"))
            .MoveNext
        Loop
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboGroup_Click()
    If mintIndex <> cboGroup.ListIndex Then FillTree cboGroup.Text
    mintIndex = cboGroup.ListIndex
End Sub


Private Sub optStyle_Click(Index As Integer)
    Dim i As Integer
    
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            picPreview.Picture = LoadResPicture(optStyle(i).Tag, vbResBitmap)
            Exit For
        End If
    Next
End Sub

Private Sub picBorder_Paint()
    Dim rc As RECT
    
    With picBorder
        rc.Left = .ScaleLeft + 1
        rc.Right = .ScaleWidth - 2
        rc.Top = .ScaleTop + 1
        rc.Bottom = .ScaleHeight - 2
    End With
    DrawEdge picBorder.hDC, rc, EDGE_RAISED, BF_RECT
End Sub

Private Sub tabMain_Click()
    fra(1).Visible = False
    fra(2).Visible = False
    fra(3).Visible = False
    fra(4).Visible = False
    fra(5).Visible = False
    fra(tabMain.SelectedItem.Index).Visible = True
    fra(tabMain.SelectedItem.Index).ZOrder 0
End Sub

Private Sub FillCommon()
'功能：装入常用的程序
    Dim var系统 As Variant, var序号 As Variant, var图标 As Variant, var标题 As Variant
    Dim lngMax As Long, lngCount As Long, lst As ListItem, strValue As String
    
    strValue = zlDatabase.GetPara("常用功能模块")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    var系统 = Split(Split(strValue, "|")(0), ",")
    var序号 = Split(Split(strValue, "|")(1), ",")
    var图标 = Split(Split(strValue, "|")(2), ",")
    var标题 = Split(Split(strValue, "|")(3), ",")
    
    lngMax = IIf(UBound(var系统) > UBound(var序号), UBound(var系统), UBound(var序号))
    lngMax = IIf(lngMax > UBound(var图标), lngMax, UBound(var图标))
    lngMax = IIf(lngMax > UBound(var标题), lngMax, UBound(var标题))
    If lngMax = -1 Then Exit Sub
    
    For lngCount = 0 To lngMax
        Set lst = lvwFavorite.ListItems.Add(, , var标题(lngCount), "K" & var图标(lngCount), "K" & var图标(lngCount))
        lst.Tag = var系统(lngCount) & "_" & var序号(lngCount)
    Next
End Sub

Private Sub tvwMain_DblClick()
    Call cmdFavorite_Click(0)
End Sub

Private Sub LoadCustom()
'完成用户习惯的初始化工作
    Dim lng简码 As Long
    Dim strPath As String, strSQL As String, strExePath As String, strLogPath As String
    Dim int药品名称显示 As Integer
    Dim int输入药品显示 As Integer
    Dim rsLoglist As ADODB.Recordset
    Dim i As Long
    
    '输入匹配
    If Val(zlDatabase.GetPara("输入匹配")) = 0 Then
        opt(opt_双向匹配).Value = True
        opt(opt_从左匹配).Value = False
    Else
        opt(opt_双向匹配).Value = False
        opt(opt_从左匹配).Value = True
    End If
    
    '汉字输入法
    Call ChooseIME(cmbIME)
    
    '简码匹配方式切换
    chkIMETurn.Value = IIf(Val(zlDatabase.GetPara("简码匹配方式切换", , , 1)) = 1, 1, 0)
    
    '简码生成方式
    lng简码 = Val(zlDatabase.GetPara("简码方式"))
    If lng简码 = 0 Then
        opt简码(0).Value = True
    Else
        opt简码(1).Value = True
    End If
    
    '自动运行本程序
    Call OS.GetRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strPath)
    chkAutoStart.Value = IIf(Trim(strPath) <> "", 1, 0)
    
    '自动关闭Windows
    chkShutDown.Value = IIf(Val(zlDatabase.GetPara("关闭Windows")) = 1, 1, 0)
    
    '界面区域隐藏
    chkAutoHide.Value = IIf(Val(zlDatabase.GetPara("界面区域隐藏")) = 1, 1, 0)
    
    '通知检查周期
    txtTime.Text = zlDatabase.GetPara("邮件消息检查周期")
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
    
    '网络断网自动重连
    chkLanJoin.Value = IIf(Val(zlDatabase.GetPara("网络断网自动重连")) = 1, 1, 0)
    '自动锁屏
    txtAutoLock.Text = Val(zlDatabase.GetPara("自动锁屏"))
    If Val(txtAutoLock.Text) = 0 Then
        chkAutoLock.Value = 0
        txtAutoLock.Text = "5"
    Else
        chkAutoLock.Value = 1
    End If
    
    '远程设置
    txtRemote.Text = Val(zlDatabase.GetPara("允许远程控制"))
    If Val(txtRemote.Text) = 0 Then
        txtRemote.Text = "1001"
    End If
    '药品名称
    int药品名称显示 = Val(zlDatabase.GetPara("药品名称显示", , , 2))
    int输入药品显示 = Val(zlDatabase.GetPara("输入药品显示"))
    
    If int药品名称显示 < 0 Or int药品名称显示 > 2 Then int药品名称显示 = 2
    If int输入药品显示 < 0 Or int输入药品显示 > 1 Then int输入药品显示 = 0
    
    cbo药品名称显示.Clear
    cbo药品名称显示.AddItem "0-显示通用名"
    cbo药品名称显示.AddItem "1-显示商品名"
    cbo药品名称显示.AddItem "2-同时显示通用名和商品名"
    cbo药品名称显示.ListIndex = int药品名称显示
    
    cbo输入药品显示.Clear
    cbo输入药品显示.AddItem "0-按输入匹配显示"
    cbo输入药品显示.AddItem "1-固定显示通用名和商品名"
    cbo输入药品显示.ListIndex = int输入药品显示
    
    '日志管理
    If OS.IsDesinMode Then
        strExePath = "C:\APPSOFT\"
    Else
        strExePath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
    End If
    strLogPath = strExePath & "Log\日志跟踪"
    
    lblLogPath.Caption = strLogPath
    vsfLogList.Editable = flexEDKbdMouse '设置可编辑状态为键盘鼠标
    vsfLogList.ColDataType(LL_启用) = flexDTBoolean '设为复选框
    On Error GoTo ErrH
    strSQL = "Select Nvl(b.名称, '服务器管理工具') 系统, a.名称, a.说明 From Zllogconfig a, Zlsystems b Where a.系统 = b.编号(+) Order By Nvl(a.系统,0), a.编号"
    Set rsLoglist = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    vsfLogList.Rows = rsLoglist.RecordCount + 1
    With rsLoglist
        For i = 1 To .RecordCount
            vsfLogList.Cell(flexcpChecked, i, LL_启用, i, LL_启用) = GetSetting("ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "启用", 0)
            vsfLogList.TextMatrix(i, LL_系统名称) = rsLoglist!系统
            vsfLogList.TextMatrix(i, LL_日志名称) = rsLoglist!名称
            vsfLogList.TextMatrix(i, LL_保存天数) = GetSetting("ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "保存天数", 3)
            vsfLogList.TextMatrix(i, LL_说明) = rsLoglist!说明
            .MoveNext
        Next
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveRegister()
'保存到注册表中的信息
    Dim lng简码 As Long
    Dim strExeName As String
        
    '因为AppTools是公共使用，为兼容9版本系统，几个在各个模块中可能使用的参数仍同时保留在注册表中
    Call zlDatabase.SetPara("输入匹配", IIf(opt(opt_双向匹配).Value = True, "0", "1"))
    SaveSetting "ZLSOFT", "公共模块\操作", "输入匹配", IIf(opt(opt_双向匹配).Value = True, "0", "1")
    
    Call zlDatabase.SetPara("输入法", IIf(cmbIME.Text = "不自动开启", "", cmbIME.Text))
    SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "输入法", IIf(cmbIME.Text = "不自动开启", "", cmbIME.Text)
    
    Call zlDatabase.SetPara("简码匹配方式切换", chkIMETurn.Value)
    
    For lng简码 = opt简码.LBound To opt简码.UBound
        If opt简码(lng简码).Value = True Then
            Call zlDatabase.SetPara("简码方式", lng简码)
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "简码生成", lng简码
            Exit For
        End If
    Next

    '自动运行本程序
    If chkAutoStart.Value = 1 Then
        strExeName = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS+.exe")
        strExeName = Replace(strExeName, ":\\", ":\")
        Call OS.SetRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strExeName)
    Else
        Call OS.DeleteRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer")
    End If
    
    '界面区域隐藏
    Call zlDatabase.SetPara("界面区域隐藏", chkAutoHide.Value)
    
    '自动关闭Windows
    Call zlDatabase.SetPara("关闭Windows", chkShutDown.Value)
    
    '消息通知检查周期
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
    Call zlDatabase.SetPara("邮件消息检查周期", Val(txtTime.Text))
    
    '网络断网自动重连
    Call zlDatabase.SetPara("网络断网自动重连", chkLanJoin.Value)
    '自动锁屏
    Call zlDatabase.SetPara("自动锁屏", IIf(chkAutoLock.Value = 0, "", Val(txtAutoLock.Text)))
    '设置远程控制端口
    Call zlDatabase.SetPara("允许远程控制", IIf(Val(txtRemote.Text) = 0, "-1", Val(txtRemote.Text)))
End Sub

Private Sub txtAutoLock_GotFocus()
    Call zlControl.TxtSelAll(txtAutoLock)
End Sub

Private Sub txtAutoLock_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAutoLock_Validate(Cancel As Boolean)
    If Val(txtAutoLock.Text) < 0 Or Val(txtAutoLock.Text) > 999 Then txtAutoLock.Text = 5
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRemote_GotFocus()
    Call zlControl.TxtSelAll(txtRemote)
End Sub

Private Sub txtRemote_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
End Sub

Private Sub txtTime_GotFocus()
    Call zlControl.TxtSelAll(txtTime)
End Sub

Private Sub vsfLogList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = LL_保存天数 Then
        If Len(vsfLogList.TextMatrix(Row, Col)) > Len(LL_最多保存天数 & "") Then vsfLogList.TextMatrix(Row, Col) = LL_最多保存天数
        If Val(vsfLogList.TextMatrix(Row, Col)) > LL_最多保存天数 Then
            vsfLogList.TextMatrix(Row, Col) = LL_最多保存天数
        Else
            vsfLogList.TextMatrix(Row, Col) = Val(vsfLogList.TextMatrix(Row, Col))
        End If
    End If
End Sub

Private Sub vsfLogList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = LL_系统名称 Or Col = LL_日志名称 Or Col = LL_说明 Then
        Cancel = True
    End If
End Sub

Private Sub vsfLogList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub vsfLogList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    lngRow = Int(Y / 260)
    If lngRow > vsfLogList.Rows - 1 Or lngRow = 0 Then
        vsfLogList.ToolTipText = ""
        Exit Sub
    End If
    If X > vsfLogList.ColWidth(LL_启用) + vsfLogList.ColWidth(LL_系统名称) + vsfLogList.ColWidth(LL_日志名称) + vsfLogList.ColWidth(LL_保存天数) Then
        vsfLogList.ToolTipText = vsfLogList.TextMatrix(lngRow, LL_说明)
    Else
        vsfLogList.ToolTipText = ""
    End If
End Sub
