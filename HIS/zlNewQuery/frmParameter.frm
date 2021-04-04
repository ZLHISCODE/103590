VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParameter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "咨询参数设置"
   ClientHeight    =   7095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6540
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   6570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   76
      Top             =   6570
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   77
      Top             =   6570
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   6360
      Left            =   75
      TabIndex        =   79
      Top             =   90
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1.基本"
      TabPicture(0)   =   "frmParameter.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk(4)"
      Tab(0).Control(1)=   "pic(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2.费用"
      TabPicture(1)   =   "frmParameter.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pic(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3.挂号"
      TabPicture(2)   =   "frmParameter.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pic(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4.价格"
      TabPicture(3)   =   "frmParameter.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5.简易挂号"
      TabPicture(4)   =   "frmParameter.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "pic(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5895
         Index           =   4
         Left            =   240
         ScaleHeight     =   5895
         ScaleWidth      =   6135
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   360
         Width           =   6135
         Begin VB.Frame Fra背景色 
            Caption         =   "背景色"
            Height          =   735
            Left            =   120
            TabIndex        =   108
            Top             =   480
            Width           =   5775
            Begin VB.PictureBox picBgColor 
               Height          =   350
               Index           =   1
               Left            =   3720
               ScaleHeight     =   285
               ScaleWidth      =   1275
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.PictureBox picBgColor 
               Height          =   350
               Index           =   0
               Left            =   960
               ScaleHeight     =   285
               ScaleWidth      =   1275
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "下标题"
               Height          =   180
               Index           =   11
               Left            =   3000
               TabIndex        =   112
               Top             =   330
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "上标题"
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   111
               Top             =   330
               Width           =   540
            End
         End
         Begin VB.CommandButton cmdSelFont 
            Caption         =   "标题字体"
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   120
            Width           =   1100
         End
         Begin VB.TextBox txt简易挂号号别 
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "设置简易挂号号别 此号别免费"
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelReg 
            Caption         =   "…"
            Height          =   300
            Left            =   2280
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   120
            Width           =   300
         End
         Begin VB.CommandButton cmdSelFont 
            Caption         =   "提示信息字体"
            Height          =   300
            Index           =   0
            Left            =   4320
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   120
            Width           =   1300
         End
         Begin VB.Frame Fra 
            Caption         =   "下标题"
            Height          =   1455
            Index           =   18
            Left            =   120
            TabIndex        =   100
            Top             =   4440
            Width           =   5775
            Begin VB.Frame Fra 
               Height          =   1095
               Index           =   19
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   5535
               Begin VB.TextBox txt下标题 
                  BackColor       =   &H00FFFFFF&
                  Height          =   705
                  Left            =   120
                  MaxLength       =   1500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   103
                  Top             =   240
                  Width           =   5295
               End
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "上标题"
            Height          =   1455
            Index           =   16
            Left            =   120
            TabIndex        =   97
            Top             =   1320
            Width           =   5775
            Begin VB.Frame Fra 
               Height          =   1095
               Index           =   17
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   5535
               Begin VB.TextBox txt上标题 
                  BackColor       =   &H00FFFFFF&
                  Height          =   705
                  Left            =   120
                  MaxLength       =   1500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   99
                  Top             =   240
                  Width           =   5295
               End
            End
         End
         Begin VB.Frame Fra 
            Height          =   1095
            Index           =   15
            Left            =   240
            TabIndex        =   96
            Top             =   3120
            Width           =   5535
            Begin VB.TextBox txt挂号提示 
               BackColor       =   &H00FFFFFF&
               Height          =   705
               Left            =   120
               MaxLength       =   1500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   101
               Text            =   "frmParameter.frx":0098
               Top             =   240
               Width           =   5295
            End
         End
         Begin MSComDlg.CommonDialog dlgThis 
            Left            =   6000
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame FraFont 
            Caption         =   "提示信息"
            Height          =   1455
            Left            =   120
            TabIndex        =   113
            Top             =   2880
            Width           =   5775
         End
         Begin VB.Label lbl简易挂号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "挂号项目"
            Height          =   180
            Left            =   120
            TabIndex        =   114
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "关闭主页上的医院信息显示(&G)"
         Height          =   180
         Index           =   4
         Left            =   -74400
         TabIndex        =   12
         Top             =   2520
         Width           =   2730
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5910
         Index           =   3
         Left            =   -74925
         ScaleHeight     =   5910
         ScaleWidth      =   6240
         TabIndex        =   88
         Top             =   360
         Width           =   6240
         Begin VB.Frame Fra 
            Caption         =   "其他"
            Height          =   1065
            Index           =   9
            Left            =   0
            TabIndex        =   63
            Top             =   4830
            Width           =   4395
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   6
               Left            =   2070
               MaxLength       =   3
               TabIndex        =   65
               Text            =   "30"
               Top             =   300
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   7
               Left            =   2070
               MaxLength       =   2
               TabIndex        =   68
               Text            =   "10"
               Top             =   630
               Width           =   690
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   300
               Left            =   2775
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   630
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(7)"
               BuddyDispid     =   196626
               BuddyIndex      =   7
               OrigLeft        =   3375
               OrigTop         =   1230
               OrigRight       =   3615
               OrigBottom      =   1530
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDWait 
               Height          =   300
               Left            =   2775
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(6)"
               BuddyDispid     =   196626
               BuddyIndex      =   6
               OrigLeft        =   3375
               OrigTop         =   885
               OrigRight       =   3615
               OrigBottom      =   1185
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "价格查询滚动间隔(&4)            秒"
               Height          =   180
               Left            =   300
               TabIndex        =   67
               Top             =   690
               Width           =   2970
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "价格查询停留时间(&3)            秒"
               Height          =   180
               Left            =   300
               TabIndex        =   64
               Top             =   360
               Width           =   2970
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "显示收费类别"
            Height          =   2445
            Index           =   8
            Left            =   4395
            TabIndex        =   70
            Top             =   90
            Width           =   1830
            Begin VB.ListBox lstShow 
               Height          =   2160
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   71
               Top             =   210
               Width           =   1650
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "显示收费项目分类"
            Height          =   4710
            Index           =   7
            Left            =   0
            TabIndex        =   59
            Top             =   90
            Width           =   4395
            Begin VB.CommandButton cmdClsAll 
               Caption         =   "全清(&D)"
               Height          =   350
               Left            =   1230
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   4275
               Width           =   1100
            End
            Begin VB.CommandButton cmdSelAll 
               Caption         =   "全选(&A)"
               Height          =   350
               Left            =   90
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   4275
               Width           =   1100
            End
            Begin MSComctlLib.TreeView tvw 
               Height          =   4035
               Left            =   90
               TabIndex        =   60
               Top             =   210
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   7117
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   494
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Checkboxes      =   -1  'True
               ImageList       =   "ils16"
               Appearance      =   1
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "缺省选中类别"
            Height          =   1635
            Index           =   1
            Left            =   4395
            TabIndex        =   74
            Top             =   4260
            Width           =   1830
            Begin VB.ListBox lstClass 
               Height          =   1320
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   75
               Top             =   225
               Width           =   1665
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "显示隐藏项目"
            Height          =   1680
            Index           =   0
            Left            =   4395
            TabIndex        =   72
            Top             =   2550
            Width           =   1815
            Begin VB.ListBox lstPrice 
               Height          =   1320
               Left            =   75
               Style           =   1  'Checkbox
               TabIndex        =   73
               Top             =   255
               Width           =   1650
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5940
         Index           =   0
         Left            =   -74940
         ScaleHeight     =   5940
         ScaleWidth      =   6255
         TabIndex        =   80
         Top             =   315
         Width           =   6255
         Begin VB.CheckBox chkUnload 
            Caption         =   "关闭查询需输入登录口令(&F)"
            Height          =   180
            Left            =   540
            TabIndex        =   92
            Top             =   2490
            Width           =   2595
         End
         Begin VB.CheckBox chkShowWorkTime 
            Caption         =   "今日就诊可查询科室上班时间(&W)"
            Height          =   180
            Left            =   540
            TabIndex        =   13
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "设备配置(&S)"
            Height          =   350
            Left            =   4875
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   315
            Width           =   1245
         End
         Begin VB.CommandButton cmdYiBao 
            Caption         =   "医保设置(&B)"
            Height          =   350
            Left            =   4875
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "0"
            Top             =   390
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   3
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "5"
            Top             =   780
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   9
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "30"
            Top             =   1140
            Width           =   570
         End
         Begin VB.TextBox txt 
            Height          =   1185
            Index           =   2
            Left            =   540
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   4725
            Width           =   5595
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   4
            Left            =   870
            TabIndex        =   90
            Top             =   4005
            Width           =   5295
         End
         Begin VB.CheckBox chkusewww 
            Caption         =   "启用医院网站(打开的网页只能用CTRL+w或ALT+F4关闭)"
            Height          =   255
            Left            =   525
            TabIndex        =   9
            Top             =   1515
            Width           =   4935
         End
         Begin VB.TextBox txturl 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   11
            Text            =   "www.zlsoft.com"
            Top             =   1830
            Width           =   4845
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   16
            Text            =   "5"
            Top             =   3465
            Width           =   540
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   3
            Left            =   870
            TabIndex        =   81
            Top             =   105
            Width           =   5295
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   2
            Left            =   750
            TabIndex        =   82
            Top             =   3135
            Width           =   5295
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   1
            Left            =   1920
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3450
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txt(1)"
            BuddyDispid     =   196626
            BuddyIndex      =   1
            OrigLeft        =   2340
            OrigTop         =   165
            OrigRight       =   2580
            OrigBottom      =   465
            Max             =   300
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   0
            Left            =   2880
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(0)"
            BuddyDispid     =   196626
            BuddyIndex      =   0
            OrigLeft        =   3300
            OrigTop         =   195
            OrigRight       =   3540
            OrigBottom      =   495
            Max             =   300
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   2
            Left            =   2880
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(3)"
            BuddyDispid     =   196626
            BuddyIndex      =   3
            OrigLeft        =   3285
            OrigTop         =   570
            OrigRight       =   3525
            OrigBottom      =   870
            Max             =   600
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   4
            Left            =   2865
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1140
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(9)"
            BuddyDispid     =   196626
            BuddyIndex      =   9
            OrigLeft        =   3285
            OrigTop         =   570
            OrigRight       =   3525
            OrigBottom      =   870
            Max             =   600
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "自动返回主页间隔(&R)           秒"
            Height          =   180
            Index           =   0
            Left            =   510
            TabIndex        =   0
            Top             =   450
            Width           =   2880
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "今日就诊刷新间隔(&T)           秒"
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Width           =   2880
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "自动检查连接间隔(&E)          分钟"
            Height          =   180
            Index           =   2
            Left            =   495
            TabIndex        =   6
            Top             =   1200
            Width           =   2970
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "公告信息"
            Height          =   180
            Index           =   8
            Left            =   135
            TabIndex        =   19
            Top             =   4020
            Width           =   720
         End
         Begin VB.Label lbl 
            Caption         =   $"frmParameter.frx":00AF
            Height          =   390
            Index           =   9
            Left            =   510
            TabIndex        =   20
            Top             =   4290
            Width           =   5610
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "网址:"
            Enabled         =   0   'False
            Height          =   180
            Index           =   3
            Left            =   810
            TabIndex        =   10
            Top             =   1875
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "信息查询"
            Height          =   180
            Index           =   5
            Left            =   135
            TabIndex        =   89
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "播放间隔"
            Height          =   180
            Index           =   4
            Left            =   15
            TabIndex        =   14
            Top             =   3150
            Width           =   720
         End
         Begin VB.Label lbl 
            Caption         =   "(注:对Flash,实际上是它播放的最短时间)"
            Height          =   180
            Index           =   7
            Left            =   2520
            TabIndex        =   18
            Top             =   3510
            Width           =   3705
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "播放间隔(&5)          秒"
            Height          =   180
            Index           =   6
            Left            =   390
            TabIndex        =   15
            Top             =   3510
            Width           =   2070
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5925
         Index           =   1
         Left            =   -74895
         ScaleHeight     =   5925
         ScaleWidth      =   6150
         TabIndex        =   85
         Top             =   390
         Width           =   6150
         Begin VB.Frame Fra 
            Caption         =   "不显示明细"
            Height          =   2940
            Index           =   10
            Left            =   4020
            TabIndex        =   26
            Top             =   45
            Width           =   2130
            Begin VB.ListBox lst 
               Height          =   2580
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   27
               Top             =   240
               Width           =   1950
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "其它"
            Height          =   2835
            Index           =   6
            Left            =   30
            TabIndex        =   28
            Top             =   3015
            Width           =   6105
            Begin VB.OptionButton opt 
               Caption         =   "按登记时间查询病人费用"
               Height          =   180
               Index           =   1
               Left            =   300
               TabIndex        =   43
               Top             =   2460
               Width           =   2310
            End
            Begin VB.OptionButton opt 
               Caption         =   "按发生时间查询病人费用"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   42
               Top             =   2160
               Value           =   -1  'True
               Width           =   2310
            End
            Begin VB.CheckBox chkExit 
               Caption         =   "允许在费用查询处输入指令退出查询(&E)"
               Height          =   225
               Left            =   345
               TabIndex        =   41
               ToolTipText     =   "退出指令""AdminExitQuery""(不区分大小写)"
               Top             =   1755
               Width           =   3585
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   5
               Left            =   2070
               MaxLength       =   5
               TabIndex        =   33
               Text            =   "10"
               Top             =   660
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   4
               Left            =   2070
               MaxLength       =   5
               TabIndex        =   30
               Text            =   "30"
               Top             =   300
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   10
               Left            =   2070
               MaxLength       =   3
               TabIndex        =   36
               Text            =   "0"
               Top             =   990
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   8
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   39
               Text            =   "0"
               Top             =   1350
               Width           =   360
            End
            Begin MSComCtl2.UpDown udn 
               Height          =   300
               Index           =   3
               Left            =   1425
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1350
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   30
               BuddyControl    =   "txt(8)"
               BuddyDispid     =   196626
               BuddyIndex      =   8
               OrigLeft        =   3285
               OrigTop         =   570
               OrigRight       =   3525
               OrigBottom      =   870
               Max             =   365
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown udn 
               Height          =   300
               Index           =   5
               Left            =   2775
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   990
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   30
               BuddyControl    =   "txt(10)"
               BuddyDispid     =   196626
               BuddyIndex      =   10
               OrigLeft        =   3300
               OrigTop         =   195
               OrigRight       =   3540
               OrigBottom      =   495
               Max             =   300
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   300
               Left            =   2775
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(4)"
               BuddyDispid     =   196626
               BuddyIndex      =   4
               OrigLeft        =   3375
               OrigTop         =   165
               OrigRight       =   3615
               OrigBottom      =   465
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDown3 
               Height          =   300
               Left            =   2775
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   660
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(5)"
               BuddyDispid     =   196626
               BuddyIndex      =   5
               OrigLeft        =   3375
               OrigTop         =   525
               OrigRight       =   3615
               OrigBottom      =   825
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费用查询停留时间(&1)            秒"
               Height          =   180
               Left            =   300
               TabIndex        =   29
               Top             =   360
               Width           =   2970
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费用查询滚动间隔(&2)            秒"
               Height          =   180
               Left            =   300
               TabIndex        =   32
               Top             =   720
               Width           =   2970
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "自动返回费用间隔(&U)            秒(注:在返回主页为0时有效)"
               Height          =   180
               Left            =   300
               TabIndex        =   35
               Top             =   1065
               Width           =   5130
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "允许查前       天的门诊费用(0不限制)"
               Height          =   180
               Left            =   315
               TabIndex        =   38
               Top             =   1395
               Width           =   3240
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "身份验证"
            Height          =   2940
            Index           =   5
            Left            =   15
            TabIndex        =   24
            Top             =   45
            Width           =   4005
            Begin VB.ListBox lstID 
               Height          =   2580
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   25
               Top             =   255
               Width           =   3870
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "汉字输入法"
            Height          =   1530
            Left            =   -10500
            TabIndex        =   86
            Top             =   90
            Visible         =   0   'False
            Width           =   3915
            Begin VB.ComboBox cmbIME 
               Height          =   300
               Left            =   0
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   0
               Visible         =   0   'False
               Width           =   2730
            End
            Begin VB.Label Label10 
               Caption         =   "     请选择一种你所喜爱的输入法作为默认输入法。它会在可进行汉字录入的位置自动打开，然后在离开时自动关闭。"
               Height          =   750
               Left            =   525
               TabIndex        =   87
               Top             =   270
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Image Image6 
               Height          =   240
               Left            =   120
               Picture         =   "frmParameter.frx":011C
               Top             =   300
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5805
         Index           =   2
         Left            =   -74925
         ScaleHeight     =   5805
         ScaleWidth      =   6210
         TabIndex        =   84
         Top             =   420
         Width           =   6210
         Begin VB.Frame Fra 
            Caption         =   "挂号类别"
            Height          =   1635
            Index           =   12
            Left            =   2775
            TabIndex        =   55
            Top             =   2205
            Width           =   3420
            Begin MSComctlLib.ListView LvwClass 
               Height          =   1335
               Left            =   90
               TabIndex        =   56
               Top             =   225
               Width           =   3240
               _ExtentX        =   5715
               _ExtentY        =   2355
               View            =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "挂号方式"
            Height          =   1635
            Index           =   11
            Left            =   45
            TabIndex        =   53
            Top             =   2205
            Width           =   2730
            Begin VB.ListBox lstGh 
               Height          =   1320
               Left            =   105
               Style           =   1  'Checkbox
               TabIndex        =   54
               Top             =   225
               Width           =   2505
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "提示内容"
            Height          =   1935
            Index           =   14
            Left            =   45
            TabIndex        =   57
            Top             =   3870
            Width           =   6150
            Begin VB.TextBox TxtDisp 
               Height          =   1620
               Left            =   135
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               Text            =   "frmParameter.frx":29FE
               Top             =   240
               Width           =   5910
            End
         End
         Begin VB.Frame Fra 
            Height          =   2220
            Index           =   13
            Left            =   45
            TabIndex        =   91
            Top             =   -45
            Width           =   6150
            Begin VB.CheckBox chk 
               Caption         =   "允许显示自助挂号返回按钮"
               Height          =   255
               Index           =   1
               Left            =   3105
               TabIndex        =   94
               Top             =   1800
               Width           =   2895
            End
            Begin VB.CheckBox chk免费 
               Caption         =   "自助挂号时不显示免费号别"
               Height          =   255
               Left            =   3105
               TabIndex        =   93
               Top             =   1388
               Width           =   2895
            End
            Begin VB.CommandButton cmdSetup 
               Caption         =   "票据打印设置(&S)"
               Height          =   350
               Left            =   210
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1605
            End
            Begin VB.CheckBox ChkPWDDisp 
               Caption         =   "就诊卡卡号密文显示"
               Height          =   210
               Left            =   255
               TabIndex        =   50
               Top             =   1125
               Width           =   1980
            End
            Begin VB.VScrollBar VSstay 
               Height          =   300
               Left            =   2955
               Min             =   1
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   630
               Value           =   10
               Width           =   255
            End
            Begin VB.VScrollBar VSFresh 
               Height          =   300
               Left            =   2955
               Min             =   1
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   285
               Value           =   10
               Width           =   255
            End
            Begin VB.TextBox TxtFreshTime 
               Height          =   300
               Left            =   1905
               MaxLength       =   4
               TabIndex        =   45
               Text            =   "600"
               Top             =   285
               Width           =   1020
            End
            Begin VB.TextBox TXTPwdDelay 
               Height          =   300
               Left            =   1905
               MaxLength       =   4
               TabIndex        =   48
               Text            =   "60"
               Top             =   630
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号时生成划价单"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   51
               Top             =   1425
               Width           =   1800
            End
            Begin VB.Label LblReshTIme 
               AutoSize        =   -1  'True
               Caption         =   "挂号安排刷新周期"
               Height          =   180
               Left            =   255
               TabIndex        =   44
               Top             =   360
               Width           =   1440
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "挂号窗体可空闲时间"
               Height          =   180
               Left            =   255
               TabIndex        =   47
               Top             =   690
               Width           =   1620
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2340
      Top             =   6555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":2A25
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":2FBF
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":3359
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarFirst As Boolean
Private mstrClass As String
Private mstrPrivs As String

Public Function ShowDialog(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Function

Private Function IsPrivs(ByVal strPriv As String) As Boolean
    IsPrivs = (InStr(";" & mstrPrivs & ";", ";" & strPriv & ";") > 0)
End Function

Private Function GetDownAllKey(objNode As Node, ByRef blnCheck As Boolean) As Boolean
    Dim objChild As Node

    On Error GoTo errHand

    objNode.Checked = blnCheck
    If objNode.Children > 0 Then

        Set objChild = objNode.Child
        Do While Not (objChild Is Nothing)

            If GetDownAllKey(objChild, blnCheck) = False Then GoTo errHand

            Set objChild = objChild.Next
        Loop

    End If

    GetDownAllKey = True

    Exit Function

errHand:

End Function

Private Function SetParentCheck(objNode As Node, ByRef blnCheck As Boolean) As Boolean
    Dim objParent As Node

    On Error GoTo errHand

    If blnCheck = False Then Exit Function

    Set objParent = objNode.Parent


    If Not (objParent Is Nothing) Then

        objParent.Checked = blnCheck

        If SetParentCheck(objParent, blnCheck) = False Then GoTo errHand

    End If

    SetParentCheck = True

    Exit Function

errHand:

End Function

Private Sub Load分类()
    Dim strTmp As String
    Dim objNode As Node
    Dim lngLoop As Long



    '显示收费项目分类,药品单独的分类
    gstrSQL = "Select -1 As ID,'西成药' As 名称,Null+0 As 上级id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,名称,Decode(上级id,Null,-1,上级id) As 上级id,'K' As PrimaryKey  From 药品用途分类  Where 材质='西成药' Start with 上级id Is Null Connect By Prior ID=上级id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -2 As ID,'中成药' As 名称,Null+0 As 上级id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,名称,Decode(上级id,Null,-2,上级id) As 上级id,'K' As PrimaryKey  From 药品用途分类  Where 材质='中成药' Start with 上级id Is Null Connect By Prior ID=上级id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -3 As ID,'中草药' As 名称,Null+0 As 上级id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,名称,Decode(上级id,Null,-3,上级id) As 上级id,'K' As PrimaryKey From 药品用途分类  Where 材质='中草药' Start with 上级id Is Null Connect By Prior ID=上级id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -4 As ID,'非药疗' As 名称,Null+0 As 上级id,'P' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,名称,Decode(上级id,Null,-4,上级id) As 上级id,'P' As PrimaryKey From 收费分类目录  Start with 上级id Is Null Connect By Prior ID=上级id"

    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF

            If zlCommFun.Nvl(gRs("上级id").Value, 0) = 0 Then
                Set objNode = tvw.Nodes.Add(, , gRs("PrimaryKey").Value & gRs("ID").Value, gRs("名称").Value, 3, 3)
            Else
                Set objNode = tvw.Nodes.Add(gRs("PrimaryKey").Value & gRs("上级id").Value, tvwChild, gRs("PrimaryKey").Value & gRs("ID").Value, gRs("名称").Value, 1, 1)
            End If
            
            gRs.MoveNext
        Wend
    End If

    tbs.Tab = 3
    DoEvents
    
    Dim blnUnSelect As Boolean
    
    strTmp = zlDatabase.GetPara("允许显示的收费分类", glngSys, 1536, "", Array(tvw), IsPrivs("参数设置"))
    If strTmp <> "" Then
        If Left(strTmp, 1) = "-" Then blnUnSelect = True
        strTmp = "," & strTmp & ","
    End If
    
    For lngLoop = 1 To tvw.Nodes.Count
        Set objNode = tvw.Nodes(lngLoop)
        If strTmp = "" Then
            tvw.Nodes(lngLoop).Checked = True
        Else
            If blnUnSelect Then
                If InStr(strTmp, ",-" & objNode.Key & ",") = 0 Then objNode.Checked = True
                
            Else
                If InStr(strTmp, "," & objNode.Key & ",") > 0 Then objNode.Checked = True
                
            End If
        End If
    Next
    tbs.Tab = 0
End Sub

Private Sub CboClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Check1_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Index = 5 Then

            tbs.Tab = 1
            If cmbIME.Enabled And cmbIME.Visible Then
                cmbIME.SetFocus
            End If
            Exit Sub

        End If
        
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub chkExit_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkExit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub ChkPWDDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkShowWorkTime_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkShowWorkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkUnload_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkUnload_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'zyk add 200410
Private Sub chkusewww_Click()
    lbl(3).Enabled = Not lbl(3).Enabled
    txturl.Enabled = Not txturl.Enabled
    cmdOK.Tag = "1"
End Sub

Private Sub chkusewww_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbIME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClsAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To tvw.Nodes.Count
        tvw.Nodes(lngLoop).Checked = False
    Next
    cmdOK.Tag = "1"
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1536)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim v_Class As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    If DecideReg = False Then Exit Sub
        
        
    '查费用时的身份验证
    '------------------------------------------------------------------------------------------------------------------

    strTmp = ""
    
    With lstID

        For lngLoop = 0 To .ListCount - 1
            
            If .Selected(lngLoop) = True Then
                strTmp = strTmp & "1"
            Else
                strTmp = strTmp & "0"
            End If
            

        Next
    End With
    

    Call SetPara("查询费用方式", strTmp, IsPrivs("参数设置"))
    
    Call SetPara("费用时间类型", IIf(opt(0).Value, 0, 1), IsPrivs("参数设置"))
    Call SetPara("挂号不显示免费号别", IIf(chk免费.Value = 1, 1, 0), IsPrivs("参数设置"))
    '------------------------------------------------------------------------------------------------------------------
    v_Class = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            v_Class = v_Class & ",'" & Left(lst.List(i), 1) & "'"
        End If
    Next
    v_Class = IIf(v_Class <> "", Mid(v_Class, 2), "")
    
    Call SetPara("费用不显明细", v_Class, IsPrivs("参数设置"))
    
    strTmp = ""
    For lngLoop = 0 To lstPrice.ListCount - 1
        If lstPrice.Selected(lngLoop) Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call SetPara("价格显示信息", strTmp, IsPrivs("参数设置"))
    
    strTmp = ""
    For lngLoop = 0 To lstClass.ListCount - 1
        If lstClass.Selected(lngLoop) Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call SetPara("价格显示类别", strTmp, IsPrivs("参数设置"))
    
    '-----------------------------------------------------------------------------------------------------------------
    strTmp = ""
    For i = 0 To lstShow.ListCount - 1
        If lstShow.Selected(i) Then
            strTmp = strTmp & ",'" & Left(lstShow.List(i), 1) & "'"
        End If
    Next
    strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")

    Call SetPara("允许显示的收费类别", strTmp, IsPrivs("参数设置"))
    
    '------------------------------------------------------------------------------------------------------------------
    Dim blnUnSelect As Boolean
    
    strTmp = ""
    For i = 1 To tvw.Nodes.Count
        If tvw.Nodes(i).Checked Then
            strTmp = strTmp & "," & tvw.Nodes(i).Key
        Else
            blnUnSelect = True
        End If
    Next
    If blnUnSelect = False Then
        strTmp = ""
    Else
        strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")
    End If
    
    Dim lngMaxLength As Long
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "Select 参数值 From zlparameters Where 1=2"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    lngMaxLength = rs.Fields(0).DefinedSize
    
    If Len(strTmp) > lngMaxLength Then
        '大于了lngMaxLength，就取没有选中的，但所有的值前加上一个负号
        
        strTmp = ""
        For i = 1 To tvw.Nodes.Count
            If tvw.Nodes(i).Checked = False Then
                strTmp = strTmp & ",-" & tvw.Nodes(i).Key
            End If
        Next
        strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")
        
        If Len(strTmp) > lngMaxLength Then
            MsgBox "选择的分类数量过多，请重新选择！", vbInformation, gstrSysName
            Exit Sub
        End If
        
    End If
    Call SetPara("允许显示的收费分类", strTmp, IsPrivs("参数设置"))
    Call SetPara("挂号时生成划价单", chk(0).Value, IsPrivs("参数设置"))
    
    If chkusewww.Value = False Then
        Call SetPara("医院主页", "", IsPrivs("参数设置"))
    Else
        Call SetPara("医院主页", txturl.Text, IsPrivs("参数设置"))
    End If
    
    If cmdOK.Tag = "1" Then
        On Error GoTo errHand
        
        cmdOK.Tag = ""
        
        Call SetPara("广告播放间隔", Val(txt(1).Text), IsPrivs("参数设置"))
        Call SetPara("返回主页间隔", Val(txt(0).Text), IsPrivs("参数设置"))
        Call SetPara("今日就诊刷新间隔", Val(txt(3).Text), IsPrivs("参数设置"))
        Call SetPara("公告信息", txt(2).Text, IsPrivs("参数设置"))
        Call SetPara("费用查询停留时间", Val(txt(4).Text), IsPrivs("参数设置"))
        Call SetPara("费用查询滚动间隔", Val(txt(5).Text), IsPrivs("参数设置"))
        Call SetPara("价格查询停留时间", Val(txt(6).Text), IsPrivs("参数设置"))
        Call SetPara("价格查询滚动间隔", Val(txt(7).Text), IsPrivs("参数设置"))
        Call SetPara("允许查以前的门诊费用", Val(txt(8).Text), IsPrivs("参数设置"))
        Call SetPara("检查数据连接间隔时间", Val(txt(9).Text), IsPrivs("参数设置"))
        
        Call SetPara("允许显示自助挂号返回按钮", chk(1).Value, IsPrivs("参数设置"))
        
        If txt(10).Enabled Then
            Call SetPara("返回费用间隔", Val(txt(10).Text), IsPrivs("参数设置"))
        Else
            Call SetPara("返回费用间隔", 0, IsPrivs("参数设置"))
        End If
        
        Call SetPara("关闭主页上的医院信息显示", chk(4).Value, IsPrivs("参数设置"))

        Call SetPara("今日就诊可查询科室上班时间", Val(chkShowWorkTime.Value), IsPrivs("参数设置"))
        
        Call gfrmMain.FrameDefault.RefreshPage
        Call gfrmMain.RefreshParamer(Val(txt(0).Text), Val(txt(9).Text))
    End If
    '-------------------------------------------------------------------
    '设置挂免费号的 Id
    '------------------------------------------------------------------
    Call SetPara("简单挂号号别", txt简易挂号号别.Tag)
    Call SaveFreeRegist
    Unload Me
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
End Sub

Private Sub cmdSelAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To tvw.Nodes.Count
        tvw.Nodes(lngLoop).Checked = True
    Next
    cmdOK.Tag = "1"
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1111", Me)
End Sub

Private Sub cmdYiBao_Click()
    gclsInsure.InsureSupport
End Sub



Private Sub Form_Activate()
    If mvarFirst = False Then Exit Sub
    mvarFirst = False
    
    Call Load分类
    
    If txt(4).Enabled And txt(4).Visible Then
        txt(4).SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim lngLoop As Long
    
'    txt(1).Text = GetInterval
    txt(1).Text = Val(zlDatabase.GetPara("广告播放间隔", glngSys, 1536, "5", Array(txt(1), udn(1)), IsPrivs("参数设置")))
    
    TxtDisp.Text = zlDatabase.GetPara("显示的提示信息", glngSys, 1536, "", Array(TxtDisp), IsPrivs("参数设置"))
    chkExit.Value = Val(zlDatabase.GetPara("允许指令退出查询", glngSys, 1536, "0", Array(chkExit), IsPrivs("参数设置")))
    chkUnload.Value = Val(zlDatabase.GetPara("关闭查询需输入登录口令", glngSys, 1536, "0", Array(chkUnload), IsPrivs("参数设置")))
    
    On Error Resume Next
    opt(CLng(zlDatabase.GetPara("费用时间类型", glngSys, 1536, "0", Array(opt(0), opt(1)), IsPrivs("参数设置")))).Value = True
    On Error GoTo 0
    chk(0).Value = Val(zlDatabase.GetPara("挂号时生成划价单", glngSys, 1536, "1", Array(chk(0)), IsPrivs("参数设置")))
    chk免费.Value = Val(zlDatabase.GetPara("挂号不显示免费号别", glngSys, 1536, "0", Array(chk免费), IsPrivs("参数设置")))
        
    txt(0).Text = Val(zlDatabase.GetPara("返回主页间隔", glngSys, 1536, 0, Array(txt(0), udn(0)), IsPrivs("参数设置")))
    txt(3).Text = Val(zlDatabase.GetPara("今日就诊刷新间隔", glngSys, 1536, 5, Array(txt(3), udn(2)), IsPrivs("参数设置")))
    txt(2).Text = zlDatabase.GetPara("公告信息", glngSys, 1536, "", Array(txt(2)), IsPrivs("参数设置"))
    txt(4).Text = Val(zlDatabase.GetPara("费用查询停留时间", glngSys, 1536, 30, Array(txt(4), UpDown1), IsPrivs("参数设置")))
    txt(5).Text = Val(zlDatabase.GetPara("费用查询滚动间隔", glngSys, 1536, 10, Array(txt(5), UpDown3), IsPrivs("参数设置")))
    txt(6).Text = Val(zlDatabase.GetPara("价格查询停留时间", glngSys, 1536, 30, Array(txt(6), UDWait), IsPrivs("参数设置")))
    txt(7).Text = Val(zlDatabase.GetPara("价格查询滚动间隔", glngSys, 1536, 10, Array(txt(7), UpDown2), IsPrivs("参数设置")))
    
    txt(8).Text = Val(zlDatabase.GetPara("允许查以前的门诊费用", glngSys, 1536, 0, Array(txt(8), udn(3)), IsPrivs("参数设置")))
    txt(9).Text = Val(zlDatabase.GetPara("检查数据连接间隔时间", glngSys, 1536, 30, Array(txt(9), udn(4)), IsPrivs("参数设置")))
    txt(10).Text = Val(zlDatabase.GetPara("返回费用间隔", glngSys, 1536, 0, Array(txt(10), udn(5)), IsPrivs("参数设置")))
        
    chk(4).Value = Val(zlDatabase.GetPara("关闭主页上的医院信息显示", glngSys, 1536, 0, Array(chk(4)), IsPrivs("参数设置")))
    chk(1).Value = Val(zlDatabase.GetPara("允许显示自助挂号返回按钮", glngSys, 1536, 0, Array(chk(1)), IsPrivs("参数设置")))
    
    
    chkShowWorkTime.Value = Val(zlDatabase.GetPara("今日就诊可查询科室上班时间", glngSys, 1536, 0, Array(chkShowWorkTime), IsPrivs("参数设置")))
    
    Dim v_Class As String
    Dim strTmp As String

    strTmp = zlDatabase.GetPara("允许显示的收费类别", glngSys, 1536, "", Array(lstShow), IsPrivs("参数设置"))
    If strTmp <> "" Then strTmp = "," & strTmp & ","
    
    v_Class = zlDatabase.GetPara("费用不显明细", glngSys, 1536, "", Array(lst), IsPrivs("参数设置"))
    v_Class = "," & v_Class & ","
    
    Set gRs = zlDatabase.OpenSQLRecord("select 编码,编码||'-'||名称 as 类别 from 收费项目类别", Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            lst.AddItem IIf(IsNull(gRs!类别), "", gRs!类别)
            If InStr(v_Class, ",'" & gRs!编码 & "',") > 0 Then lst.Selected(lst.NewIndex) = True
            
            lstShow.AddItem IIf(IsNull(gRs!类别), "", gRs!类别)

            If strTmp = "" Then
                lstShow.Selected(lstShow.NewIndex) = True
            Else
                If InStr(strTmp, ",'" & gRs!编码 & "',") > 0 Then lstShow.Selected(lstShow.NewIndex) = True
            End If
            
            gRs.MoveNext
        Wend
    End If
    
    
    '查费用时的身份验证,000000000
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = Trim(zlDatabase.GetPara("查询费用方式", glngSys, 1536, "100000000", Array(lstID), IsPrivs("参数设置")))
    If strTmp = "000000000" Then strTmp = "100000000"
    
    With lstID
        .Clear
        .AddItem "允许通过就诊卡查费用"
        .AddItem "允许通过门诊号查费用"
        .AddItem "允许通过住院号查费用"
        .AddItem "允许通过病人ID号查费用"
        .AddItem "允许通过医保卡查费用"
        .AddItem "允许通过身份证查费用"
        .AddItem "允许通过票据号查费用"
        .AddItem "允许通过单据号查费用"
        .AddItem "允许通过ＩＣ卡查费用"
        
        For lngLoop = 0 To .ListCount - 1
'            If Val(zlDatabase.GetPara(.List(lngLoop), glngSys, 1536, IIf(lngLoop = 0, 1, 0), Array(lstID), IsPrivs("参数设置"))) = 1 Then
'                .Selected(lngLoop) = True
'            End If
            
            If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
                .Selected(lngLoop) = True
            End If
            
        Next
    End With
    '
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = Trim(zlDatabase.GetPara("价格显示信息", glngSys, 1536, "0000011", Array(lstPrice), IsPrivs("参数设置")))
    If Len(strTmp) = 6 Then strTmp = strTmp & "1"
    
    lstPrice.Clear
    lstPrice.AddItem "费用类型"
    lstPrice.AddItem "编码"
    lstPrice.AddItem "产地"
    lstPrice.AddItem "标识主码"
    lstPrice.AddItem "标识子码"
    lstPrice.AddItem "指导售价"
    lstPrice.AddItem "剂型"
    
    For lngLoop = 0 To lstPrice.ListCount - 1
        
        If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
            lstPrice.Selected(lngLoop) = True
        End If
        
    Next
    
    Dim blnHave As Boolean
    
    '挂号方式初始
    '------------------------------------------------------------------------------------------------------------------
    
    With lstGh
        .Clear
        .AddItem "就诊卡"
        .AddItem "医保卡"
        .AddItem "身份证"
        .AddItem "ＩＣ卡"
        
        '挂号类别
        strTmp = "," & zlDatabase.GetPara("挂号类别", glngSys, 1536, "", Array(lstGh), IsPrivs("参数设置")) & ","
        If strTmp = ",两者都可以," Then strTmp = ",就诊卡,医保卡,"
        
        blnHave = False
        For lngLoop = 0 To .ListCount - 1
        
            If InStr(strTmp, "," & .List(lngLoop) & ",") > 0 Then
                .Selected(lngLoop) = True
                blnHave = True
            End If

        Next
        
        If blnHave = False Then
            For lngLoop = 0 To .ListCount - 1
                .Selected(lngLoop) = True
            Next
        End If
    End With
    
    
    strTmp = Trim(zlDatabase.GetPara("价格显示类别", glngSys, 1536, "000000", Array(lstClass), IsPrivs("参数设置")))
    
    lstClass.Clear
    lstClass.AddItem "药疗"
    lstClass.AddItem "检验"
    lstClass.AddItem "检查"
    lstClass.AddItem "治疗"
    lstClass.AddItem "手术"
    lstClass.AddItem "其他所有"
    
    blnHave = False
    For lngLoop = 0 To lstClass.ListCount - 1
'        If Val(zlDatabase.GetPara(lstClass.List(lngLoop), glngSys, 1536, 0, Array(lstClass), IsPrivs("参数设置"))) = 1 Then
'            lstClass.Selected(lngLoop) = True
'            blnHave = True
'        End If

        If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
            lstClass.Selected(lngLoop) = True
            blnHave = True
        End If
        
    Next
    If blnHave = False Then
        For lngLoop = 0 To lstClass.ListCount - 1
            lstClass.Selected(lngLoop) = True
        Next
    End If
    
    cmdOK.Tag = ""
    
    mvarFirst = True
    '将自助挂号信息进行初始化
    LoadRegSelef
    
    Dim wwwurl As String
    
    wwwurl = zlDatabase.GetPara("医院主页", glngSys, 1536, "", Array(chkusewww, txturl), IsPrivs("参数设置"))
    If wwwurl <> "" Then
        chkusewww.Value = 1
        lbl(3).Enabled = True
        txturl.Enabled = True
        txturl.Text = wwwurl
    End If
     Call LoadFreeRegist
     Call InitFreeRegist
    
    
    
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstClass_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub lstClass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
    
End Sub

Private Sub lstGh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstPrice_ItemCheck(Item As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub lstPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstShow_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub lstShow_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub LvwClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        tbs.Tab = 2
        TxtFreshTime.SetFocus
    End If
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    Dim i As Long
    
    tbs.ZOrder 0
    For i = 0 To pic.UBound
        pic(i).Enabled = False
    Next
    pic(tbs.Tab).Enabled = True
    
    Select Case tbs.Tab
        Case 0
            If txt(0).Enabled Then txt(0).SetFocus
        Case 1
            If lstID.Enabled Then lstID.SetFocus
        Case 2
            If TxtFreshTime.Enabled Then TxtFreshTime.SetFocus
        Case 3
            If tvw.Enabled Then tvw.SetFocus
    End Select
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)

    Dim lngLoop As Long
    Dim blnCheck As Boolean

    blnCheck = Node.Checked
    '下级
    Call GetDownAllKey(Node, Node.Checked)

    '向上
    Call SetParentCheck(Node, Node.Checked)
    cmdOK.Tag = "1"
End Sub


Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "1"
    
    If Index = 0 Then
        udn(5).Enabled = (Val(txt(Index).Text) = 0)
        txt(10).Enabled = (Val(txt(Index).Text) = 0)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
        
   If KeyAscii = 13 Then
        KeyAscii = 0
        
        Select Case Index
        Case 2
            tbs.Tab = 1
            'tbs.Tabs(1).Select = True
            
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    
    Else
        Select Case Index
        Case 2
        Case Else
            If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
        End Select
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Cancel = False Then
        Select Case Index
        Case 1
            If Val(txt(Index).Text) < 5 Then
                MsgBox "广告播放的时间间隔至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 300 Then
                MsgBox "广告播放的时间间隔至多5分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 0
            If Val(txt(Index).Text) > 300 Then
                MsgBox "返回主页的时间间隔至多300秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 3
            If Val(txt(Index).Text) < 5 Then
                MsgBox "返回主页的时间间隔至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "返回主页的时间间隔至多600秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 4
            If Val(txt(Index).Text) < 5 Then
                MsgBox "费用查询停留时间至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "费用查询停留时间至多10分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 5
            If Val(txt(Index).Text) < 5 Then
                MsgBox "费用查询滚动间隔至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "费用查询滚动间隔至多10分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 6
            If Val(txt(Index).Text) < 5 Then
                MsgBox "价格查询停留时间至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "价格查询停留时间至多10分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 7
            If Val(txt(Index).Text) < 5 Then
                MsgBox "价格查询滚动间隔至少5秒钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "价格查询滚动间隔至多10分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 8
            If Val(txt(Index).Text) < 0 Then
                MsgBox "天数不能为负数！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 365 Then
                MsgBox "查询门诊费用不能超过365天！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 9
            If Val(txt(Index).Text) < 1 Then
                MsgBox "检查连接间隔不能小于1分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 365 Then
                MsgBox "检查连接间隔不能大于600(即10小时)分钟！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        End Select
        
    End If
End Sub


Private Sub TxtDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tbs.Tab = 3
        tvw.SetFocus
        
    End If
End Sub

Private Sub TxtFreshTime_GotFocus()
    SelAll TxtFreshTime
End Sub

Private Sub LoadRegSelef()
    Dim rsTmp As New ADODB.Recordset
    Dim Itmx As ListItem
    Dim i As Integer
    
    
    On Error GoTo ErrHandle
    
    Call ReadRegest                             '从注册表之中读取初始化数据
    
    '将该系统不处理的号类显示
    gstrSQL = "select 编码,名称 from 号类"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    i = 1
    LvwClass.ListItems.Clear
    Do While Not rsTmp.EOF
        Set Itmx = LvwClass.ListItems.Add(, "K" + CStr(i), CStr(rsTmp("名称")))
        If InStr(mstrClass, CStr(rsTmp("名称"))) > 0 Then Itmx.Checked = True
    rsTmp.MoveNext
    i = i + 1
    Loop
    rsTmp.Close
    '将系统处理的号类显示
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReadRegest()
    '将各个参数从注册表之中读取并且显示在界面
    If zlDatabase.GetPara("主窗体刷新周期", glngSys, 1536, "", Array(TxtFreshTime, VSFresh), IsPrivs("参数设置")) = "" Then
        TxtFreshTime.Text = 600
        TXTPwdDelay.Text = 60
'        CboClass.Text = "就诊卡挂号"
        ChkPWDDisp.Value = 0
    Else
        TxtFreshTime.Text = zlDatabase.GetPara("主窗体刷新周期", glngSys, 1536, "", Array(TxtFreshTime, VSFresh), IsPrivs("参数设置"))
        TXTPwdDelay.Text = zlDatabase.GetPara("密码验证窗体停留时间", glngSys, 1536, "", Array(TXTPwdDelay, VSstay), IsPrivs("参数设置"))

        mstrClass = zlDatabase.GetPara("挂号的号类", glngSys, 1536, "", Array(LvwClass), IsPrivs("参数设置"))
        ChkPWDDisp.Value = Val(zlDatabase.GetPara("密文显示卡号", glngSys, 1536, 0, Array(ChkPWDDisp), IsPrivs("参数设置")))
        
    End If
End Sub

Private Sub WriteRegedit()
    Dim i As Long
    Dim strTmp As String
    
    '将变量从界面抄写进入注册表
    mstrClass = ""
    For i = 1 To CLng(LvwClass.ListItems.Count)
        If LvwClass.ListItems("K" + CStr(i)).Checked = True Then
          mstrClass = mstrClass + "'" + LvwClass.ListItems("K" + CStr(i)).Text + "',"
        End If
    Next
    If Trim(mstrClass) <> "" Then mstrClass = Mid(mstrClass, 1, Len(mstrClass) - 1)
    
    strTmp = ""
    With lstGh
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = strTmp & "," & .List(i)
            End If
        Next
    End With
    
    Call SetPara("主窗体刷新周期", Val(TxtFreshTime.Text), IsPrivs("参数设置"))
    Call SetPara("密码验证窗体停留时间", Val(TXTPwdDelay.Text), IsPrivs("参数设置"))
    Call SetPara("挂号类别", strTmp, IsPrivs("参数设置"))
    Call SetPara("挂号的号类", mstrClass, IsPrivs("参数设置"))
    Call SetPara("密文显示卡号", ChkPWDDisp.Value, IsPrivs("参数设置"))
    Call SetPara("显示的提示信息", TxtDisp.Text, IsPrivs("参数设置"))
    Call SetPara("允许指令退出查询", chkExit.Value, IsPrivs("参数设置"))
    Call SetPara("关闭查询需输入登录口令", chkUnload.Value, IsPrivs("参数设置"))

End Sub

Private Sub TxtFreshTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
End Sub

Private Sub TxtFreshTime_LostFocus()
    If Not IsNumeric(TxtFreshTime.Text) Then
        MsgBox "请将刷新周期设置为数字信息", vbInformation, gstrSysName
        tbs.Tab = 2
        TxtFreshTime.SetFocus
    End If
End Sub

Private Sub TXTPwdDelay_GotFocus()
    SelAll TXTPwdDelay
End Sub

Private Sub TXTPwdDelay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
End Sub

Private Sub TXTPwdDelay_LostFocus()
    If Not IsNumeric(TXTPwdDelay.Text) Then
        MsgBox "请将挂号窗体可空闲时间设置为数字信息", vbInformation, gstrSysName
        tbs.Tab = 2
        TXTPwdDelay.SetFocus
    End If
End Sub
'zyk add 200410
Private Sub txturl_Change()
        cmdOK.Tag = "1"
End Sub

Private Sub VSFresh_Change()
    If Not IsNumeric(TxtFreshTime.Text) Then
        MsgBox "请将刷新时间设置为数字信息", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TxtFreshTime.Text < 1 Then TxtFreshTime.Text = 1
    TxtFreshTime.Text = TxtFreshTime.Text + 10 - VSFresh.Value
    VSFresh.Value = 10
End Sub

Private Sub VSstay_Change()
    If Not IsNumeric(TXTPwdDelay.Text) Then
        MsgBox "请将刷新时间设置为数字信息", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TXTPwdDelay.Text < 1 Then TXTPwdDelay.Text = 1
    TXTPwdDelay.Text = TXTPwdDelay.Text + 10 - VSstay.Value
    VSstay.Value = 10
End Sub

Private Function DecideReg() As Boolean
    Dim i As Integer

    '判断挂号的号类信息
    mstrClass = ""
    For i = 1 To CLng(LvwClass.ListItems.Count)
        If LvwClass.ListItems("K" + CStr(i)).Checked = True Then
          mstrClass = mstrClass + "'" + LvwClass.ListItems("K" + CStr(i)).Text + "',"
        End If
    Next
'    If mstrClass = "" Then
'       MsgBox "请至少选择一项挂号项目", vbInformation, gstrSysName
'       DecideReg = False: Exit Function
'    End If
     '判断刷新时间的
    If Not IsNumeric(TxtFreshTime.Text) Then
         MsgBox "请将刷新时间设置为正数信息", vbInformation, gstrSysName
         If TxtFreshTime.Enabled And TxtFreshTime.Visible Then TxtFreshTime.SetFocus
         DecideReg = False: Exit Function
    End If

    If TxtFreshTime.Text < 0 Or TxtFreshTime.Text > 9999 Then
         MsgBox "请将刷新时间设置为0到9999的正数信息", vbInformation, gstrSysName
         If TxtFreshTime.Enabled And TxtFreshTime.Visible Then TxtFreshTime.SetFocus
         DecideReg = False: Exit Function
    End If
   '判断密码延迟窗体
    If Not IsNumeric(TXTPwdDelay.Text) Then
         MsgBox "请将密码验证窗体的延迟时间设置为1到9999的正数信息", vbInformation, gstrSysName
         If TXTPwdDelay.Enabled And TXTPwdDelay.Visible Then TXTPwdDelay.SetFocus
         DecideReg = False: Exit Function
    End If
    If (TXTPwdDelay.Text > 9999) Or (TXTPwdDelay.Text < 0) Then
         MsgBox "请将密码验证窗体的延迟时间设置为0到9999的正数信息", vbInformation, gstrSysName
         If TXTPwdDelay.Enabled And TXTPwdDelay.Visible Then TXTPwdDelay.SetFocus
         DecideReg = False: Exit Function
    End If

    On Error GoTo ErrHandle
    
    Call WriteRegedit
    
    DecideReg = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadFreeRegist()
'---------------------------
'加载已经设置得 简易挂号类别
'---------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "select a.号码 as 号别,'['|| a.号码||']'||b.名称 as 名称 from 挂号安排 a,收费项目目录 b  where a.项目id =b.id and  a.号码=[1]"
    On Error GoTo hErr
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, GetPara("简单挂号号别", -1, True))
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Me.txt简易挂号号别.Text = rsTmp!名称
    Me.txt简易挂号号别.Tag = Nvl(rsTmp!号别, -1)
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub selFreeRegist()
'------------------------------------------------
'功能:加载当前可挂的挂号
'------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    Dim strTime As String
    Dim i As Long
   vRect = GetControlRect(Me.txt简易挂号号别.hwnd)
   On Error GoTo ErrHandle
            '求出当前时间属于具体的时间段
            strTime = _
                  "Select 时间段 From 时间段 Where" & _
                  " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                  " Between" & _
                  " Decode(Sign(开始时间 - 终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & _
                  " And" & _
                  " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
                  " Or" & _
                  " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                  " Between" & _
                  " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & _
                  " And" & _
                  " Decode(Sign(开始时间 - 终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
   
   
            strSQL = "" & _
            "   Select  distinct M.ID as ID,M.号码 as 号别,M.科室ID as 科室ID,M.号类 as 号类,M.项目ID as 项目ID,C.名称 as 科室, " & _
            "             N.名称 as 名称,Nvl(M.医生姓名, ' ') as 医生姓名,M.医生id,Decode(To_Char(SysDate,'D'),'1',M.周日," & _
            "             '2',M.周一,'3',M.周二,'4',M.周三,'5',M.周四,'6',M.周五,'7',M.周六)  as 时间" & _
            "   From 挂号安排 M,收费项目目录 N,部门表 C " & _
            "   Where M.ID not in (  Select  A.ID from 挂号安排 A,病人挂号汇总 B,挂号安排限制 C " & _
            "                                   Where  a.科室ID = B.科室ID And a.项目ID = B.项目ID And a.id=c.安排ID(+) And " & _
            "                                          Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =C.限制项目(+) " & vbNewLine & _
            "                                          And Nvl(A.医生ID,0)=Nvl(B.医生ID,0)  And   a.科室ID = B.科室ID And a.项目ID = B.项目ID   And " & GetNodeCheckSQL("N.站点") & " And " & GetNodeCheckSQL("C.站点") & " And " & _
            "                                               B.日期=Trunc(Sysdate)  and c.限号数<= B.已挂数 and C.限号数<>0 ) " & _
            "               And  Decode(To_Char(SysDate,'D'),'1',M.周日,'2',M.周一,'3',M.周二,'4', M.周三,'5',M.周四,'6',M.周五,'7',M.周六) in (" + strTime + ")  " & _
            "               And M.项目ID=N.ID  and M.科室ID=C.ID   " & _
            "               And M.停用日期 is NULL And (M.医生id Is Null Or Exists (Select 1 From 人员表 y Where y.ID=M.医生id And " & GetNodeCheckSQL("y.站点") & _
            "               And (y.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.撤档时间 Is Null)) ) " & _
            "               And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=M.ID and Sysdate between 开始停止时间 and 结束停止时间 )"
            
            strSQL = "" & _
            "  Select 号别 as ID ,'['||号别||']'||名称 as 名称,科室,医生姓名,号类,时间,sum(nvl(价格,0)) as 价格 " & _
            "  From ( With A1 as (" & strSQL & ") " & _
            "           Select  A1.*,D.现价 as 价格  From A1,收费价目 D " & _
            "           Where A1.项目ID=D.收费细目ID And     D.执行日期<=sysdate and (D.终止日期> sysdate or D.终止日期 is null)  " & _
            "           Union all " & _
            "           Select  A1.*,D.现价 as 价格  From A1,收费从属项目 A,收费价目 D " & _
            "           Where A1.项目ID=A.主项ID and A.从项ID=D.收费细目ID  And  D.执行日期<=sysdate and (D.终止日期> sysdate or D.终止日期 is null)  " & _
            "       )" & _
            " Group by ID,号类,号别,科室ID,项目ID,科室,名称,医生姓名,医生id,时间   Having sum(nvl(价格,0))=0" & _
              vbNewLine & "  union all  " & vbNewLine & _
             " select '-1' as id ,'[不设置简易挂号项目]' as 名称, null as 科室,null as 医生姓名,null as 号类,null as 时间,null as 价格 from Dual " & _
             "   Order by 科室,价格"
            '科室ID,项目ID,医生id,
            Set rsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "挂号项目选择", False, _
            Nvl(Me.txt简易挂号号别.Tag, -1), "", False, False, True, vRect.Left, vRect.Top, Me.txt简易挂号号别.Height, blnCancel, True, True)
            If blnCancel Then Exit Sub
            If rsInfo Is Nothing Then
                MsgBox "当前没有可用的挂号项目！请到挂号安排中设置！", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
            Me.txt简易挂号号别.Text = IIf(Nvl(rsInfo!ID, -1) = -1, "", Nvl(rsInfo!名称))
            Me.txt简易挂号号别.Tag = Nvl(rsInfo!ID, -1)
            
            Exit Sub
ErrHandle:
            If ErrCenter() = 1 Then Resume
            SaveErrLog
End Sub


Private Sub InitFreeRegist()
    Dim strFontName As String, strMsg As String, dblColor As Double, dblSize As Double
    Dim dblUpBgColor As Double, dblDownBgColor As Double
    Dim blnBold As Boolean, blnItalic As Boolean
    '提示信息
    If GetRegistParaFont("简单挂号提示信息", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With txt挂号提示
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
'            .Font.Size = dblSize
            .ForeColor = dblColor
            If dblColor = vbWhite Then
              .BackColor = &HE0E0E0
            Else
               .BackColor = vbWhite
            End If
        End With
    End If
    If GetRegistParaFont("简单挂号上标题", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.txt上标题
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
'            .SelStart=0:.SelLength=1:.setfo
            .ForeColor = dblColor
            If dblColor = vbWhite Then
                .BackColor = &HE0E0E0
             Else
               .BackColor = vbWhite
            End If
        End With
    End If
    If GetRegistParaFont("简单挂号下标题", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
      With Me.txt下标题
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
            '          .Font.Size = dblSize
            .ForeColor = dblColor
            If dblColor = vbWhite Then
              .BackColor = &HE0E0E0
            Else
               .BackColor = vbWhite
            End If
      End With
    End If
    If GetFreeRegistBGColor(dblUpBgColor, dblDownBgColor) Then
       Me.picBgColor(0).BackColor = dblUpBgColor
       Me.picBgColor(1).BackColor = dblDownBgColor
    End If
    
End Sub

Private Function SaveFreeRegist()
    With Me.txt挂号提示
        SetRegistParaFont "简单挂号提示信息", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                          .FontBold, .FontItalic
        
    End With
    With Me.txt上标题
       SetRegistParaFont "简单挂号上标题", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                          .FontBold, .FontItalic
    End With
    With Me.txt下标题
       SetRegistParaFont "简单挂号下标题", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                         .FontBold, .FontItalic
    End With
    SetFreeRegistBGColor CDbl(picBgColor(0).BackColor), CDbl(Me.picBgColor(1).BackColor)
End Function

Private Sub cmdSelFont_Click(Index As Integer)
    '---------------------
    '设置简易挂号 相关字体颜色等
    '----------------------
    With Me.dlgThis
        Select Case Index
    
        Case 0
            .DialogTitle = "设置简易挂号提示信息字体"
            .flags = &H2 + &H1 + &H400 + &H800 + &H100  '&H100000 +
            .Color = Me.txt挂号提示.ForeColor
            .FontBold = txt挂号提示.Font.Bold
            .FontItalic = txt挂号提示.FontItalic
            .FontName = txt挂号提示.Font.Name
            .FontSize = IIf(Val(txt挂号提示.Tag) > 0, Val(txt挂号提示.Tag), txt挂号提示.Font.Size)
             Err.Clear: On Error Resume Next:
             .CancelError = True
            .ShowFont
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
             txt挂号提示.Tag = Val(.FontSize)
             txt挂号提示.Font.Name = .FontName
             txt挂号提示.Font.Bold = .FontBold
             txt挂号提示.Font.Italic = .FontItalic
             txt挂号提示.ForeColor = .Color
             If .Color = vbWhite Then
                 txt挂号提示.BackColor = &HE0E0E0
             Else
                txt挂号提示.BackColor = vbWhite
             End If
             
        Case 1
             .DialogTitle = "设置简易挂号标题字体"
            .flags = &H2 + &H1 + &H400 + &H800 + &H100
            .Color = Me.txt上标题.ForeColor
            .FontBold = txt上标题.Font.Bold
            .FontItalic = txt上标题.FontItalic
            .FontName = txt上标题.Font.Name
            .FontSize = IIf(Val(txt上标题.Tag) > 0, Val(txt上标题.Tag), txt上标题.Font.Size)
             Err.Clear: On Error Resume Next:
             .CancelError = True
            .ShowFont
            If Err.Number <> 0 Then Err.Clear:  On Error GoTo 0: Exit Sub
             txt上标题.Tag = Val(.FontSize)
             txt上标题.Font.Name = .FontName
             txt上标题.Font.Bold = .FontBold
             txt上标题.Font.Italic = .FontItalic
             txt上标题.ForeColor = .Color
             txt下标题.Tag = Val(.FontSize)
             txt下标题.Font.Name = .FontName
             txt下标题.Font.Bold = .FontBold
             txt下标题.Font.Italic = .FontItalic
             txt下标题.ForeColor = .Color
             If .Color = vbWhite Then
                 txt上标题.BackColor = &HE0E0E0
                 txt下标题.BackColor = &HE0E0E0
             Else
                txt上标题.BackColor = vbWhite
                txt下标题.BackColor = vbWhite
             End If
        End Select
    End With
  
     
End Sub

Private Sub cmdSelReg_Click()
    Call selFreeRegist
End Sub


Private Sub picBgColor_Click(Index As Integer)
        With Me.dlgThis
        
        .DialogTitle = "设置简易挂号上标题颜色"
        .flags = &H2 + &H1
        .Color = Me.picBgColor(Index).BackColor
        Err.Clear: On Error Resume Next:
        .CancelError = True
        .ShowColor
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
        Me.picBgColor(Index).BackColor = .Color: picBgColor(Index).Tag = 1
'        If Index = 0 Then
'            Me.txt上标题.BackColor = .Color
'        Else
'            Me.txt下标题.BackColor = .Color
'        End If
    End With
End Sub
