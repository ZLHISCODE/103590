VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLisStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5970
   Icon            =   "frmLisStationPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5160
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   375
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6495
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   6315
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&1.报告"
      TabPicture(0)   =   "frmLisStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2.文件提取"
      TabPicture(1)   =   "frmLisStationPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFile"
      Tab(1).Control(1)=   "cmdFile"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboDevice"
      Tab(1).Control(3)=   "optRange(0)"
      Tab(1).Control(4)=   "optRange(1)"
      Tab(1).Control(5)=   "cardSet"
      Tab(1).Control(6)=   "dtpStart"
      Tab(1).Control(7)=   "dtpEnd"
      Tab(1).Control(8)=   "lblNotify"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(11)=   "Label13"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&3.科室打印"
      TabPicture(2)   =   "frmLisStationPara.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSelectAll"
      Tab(2).Control(1)=   "cmdClearAll"
      Tab(2).Control(2)=   "lvwDept"
      Tab(2).Control(3)=   "Label1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&4.其他"
      TabPicture(3)   =   "frmLisStationPara.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmsign"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame9"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame8"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame frmsign 
         Caption         =   "签名流程"
         Height          =   855
         Left            =   -74820
         TabIndex        =   76
         Top             =   2640
         Width           =   5280
         Begin VB.CheckBox checkSaveReprotSign 
            Caption         =   "报告单保存时签名"
            Height          =   315
            Left            =   2970
            TabIndex        =   78
            Top             =   300
            Width           =   2025
         End
         Begin VB.CheckBox checkSaveInfoSign 
            Caption         =   "核收登记保存时签名"
            Height          =   315
            Left            =   210
            TabIndex        =   77
            Top             =   300
            Width           =   2025
         End
      End
      Begin VB.Frame Frame11 
         Height          =   525
         Left            =   -74820
         TabIndex        =   66
         Top             =   510
         Width           =   5280
         Begin VB.OptionButton opt门诊处理 
            Caption         =   "提示修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   69
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt门诊处理 
            Caption         =   "自动修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   68
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton opt门诊处理 
            Caption         =   "不修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   67
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "门诊病人信息不一致时"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   70
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.Frame Frame10 
         Height          =   525
         Left            =   -74820
         TabIndex        =   61
         Top             =   1035
         Width           =   5280
         Begin VB.OptionButton opt住院处理 
            Caption         =   "不修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   64
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton opt住院处理 
            Caption         =   "自动修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   63
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton opt住院处理 
            Caption         =   "提示修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   62
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.Label Label11 
            Caption         =   "住院病人信息不一致时"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   65
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.Frame Frame9 
         Height          =   525
         Left            =   -74820
         TabIndex        =   56
         Top             =   2085
         Width           =   5280
         Begin VB.OptionButton opt体检处理 
            Caption         =   "提示修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   59
            Top             =   165
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt体检处理 
            Caption         =   "自动修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3045
            TabIndex        =   58
            Top             =   165
            Width           =   1065
         End
         Begin VB.OptionButton opt体检处理 
            Caption         =   "不修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4260
            TabIndex        =   57
            Top             =   165
            Width           =   960
         End
         Begin VB.Label lbl病人信息 
            Caption         =   "体检病人信息不一致时"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   60
            TabIndex        =   60
            Top             =   195
            Width           =   2160
         End
      End
      Begin VB.Frame Frame8 
         Height          =   525
         Left            =   -74820
         TabIndex        =   51
         Top             =   1575
         Width           =   5280
         Begin VB.OptionButton opt院外处理 
            Caption         =   "提示修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   54
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt院外处理 
            Caption         =   "自动修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   53
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton opt院外处理 
            Caption         =   "不修正"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   52
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "院外病人信息不一致时"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   55
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "全选(&A)"
         Height          =   375
         Left            =   -70680
         TabIndex        =   48
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&L)"
         Height          =   375
         Left            =   -70680
         TabIndex        =   47
         Top             =   1440
         Width           =   1100
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   -74700
         TabIndex        =   41
         Top             =   1560
         Width           =   4725
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&S"
         Height          =   300
         Left            =   -69990
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   300
      End
      Begin VB.ComboBox cboDevice 
         Height          =   300
         Left            =   -73530
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1980
         Width           =   3855
      End
      Begin VB.OptionButton optRange 
         Caption         =   "只提取当天数据(&T)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   38
         Top             =   2460
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton optRange 
         Caption         =   "提取指定时段数据(&R)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   -72600
         TabIndex        =   37
         Top             =   2460
         Width           =   2025
      End
      Begin VB.Frame cardSet 
         Caption         =   "读卡接口设置"
         Height          =   810
         Left            =   -74670
         TabIndex        =   33
         Top             =   3240
         Width           =   4875
         Begin VB.CommandButton cmdIC 
            Caption         =   "IC卡设置(I)"
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   285
            Width           =   1215
         End
         Begin VB.CommandButton cmdIdent 
            Caption         =   "设备配置(&S)"
            Height          =   390
            Left            =   330
            TabIndex        =   34
            Top             =   285
            Width           =   1260
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "其他本地参数"
         Height          =   3810
         Left            =   120
         TabIndex        =   9
         Top             =   2370
         Width           =   5445
         Begin VB.CheckBox chkSampleType 
            Caption         =   "上次结果不参照标本类型"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   79
            ToolTipText     =   "上次结果不参照标本类型"
            Top             =   2685
            Width           =   2640
         End
         Begin VB.CheckBox chkAutoAddItem 
            Caption         =   "自动增加计算项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   75
            ToolTipText     =   "自动增加未申请的计算项目"
            Top             =   2430
            Width           =   1965
         End
         Begin VB.CheckBox chkLoadLast 
            Caption         =   "登记时保留上一次申请项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   74
            Top             =   2430
            Width           =   2745
         End
         Begin VB.CheckBox chkLast 
            Caption         =   "核收时提示上次超标结果"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3030
            TabIndex        =   73
            Top             =   3450
            Width           =   2355
         End
         Begin VB.CheckBox chkOnlyMachine 
            Caption         =   "只核收当前仪器项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   72
            Top             =   2160
            Width           =   2265
         End
         Begin VB.CheckBox chkSkipRule 
            Caption         =   "审核后自动跳到下一个可审标本"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   71
            Top             =   2160
            Width           =   2835
         End
         Begin VB.CheckBox chkNotSend 
            Caption         =   "使用二级报告审核"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   32
            Top             =   1890
            Width           =   2265
         End
         Begin VB.CheckBox chkItemNumber 
            Caption         =   "手工项目按项目累加标本号"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   1890
            Width           =   2685
         End
         Begin VB.CheckBox ChkCheckInNoItem 
            Caption         =   "登记时不需要输入项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   28
            Top             =   1620
            Width           =   2265
         End
         Begin VB.CheckBox chkShowOption 
            Caption         =   "只在核收登记时显示登记窗口"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   1620
            Width           =   2715
         End
         Begin VB.CheckBox chkNO 
            Caption         =   "按上次输入的标本号累加"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   26
            Top             =   1350
            Width           =   2355
         End
         Begin VB.CheckBox chkShowType 
            Caption         =   "自适应高度自动分列显示结果"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   1350
            Width           =   2685
         End
         Begin VB.CheckBox chkPatientType 
            Caption         =   "所有登记病人标识为外来"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   23
            Top             =   1089
            Width           =   2355
         End
         Begin VB.CheckBox chkShowAll 
            Caption         =   "不区分仪器显示核收项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1089
            Width           =   2355
         End
         Begin VB.CheckBox ChkPrivacy 
            Caption         =   "报告单是否显示隐私项目"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   21
            Top             =   816
            Width           =   2355
         End
         Begin VB.CheckBox chkNoRange 
            Caption         =   "核收时忽略指定的时间范围(&I)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   3180
            Width           =   2745
         End
         Begin VB.CheckBox chkCheck 
            Caption         =   "核收时显示是否收费(&N)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3030
            TabIndex        =   16
            Top             =   3180
            Width           =   2355
         End
         Begin VB.CheckBox chkAutoRefresh 
            Caption         =   "收到仪器数据自动刷新(&A)"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   15
            Top             =   270
            Width           =   2745
         End
         Begin VB.CheckBox chkComm 
            Caption         =   "核收时允许双向通信(&D)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   3450
            Width           =   2355
         End
         Begin VB.CheckBox chkSample 
            Caption         =   "登记时可直接输入病人信息"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   13
            Top             =   816
            Width           =   2835
         End
         Begin VB.CheckBox chkEmerge 
            Caption         =   "标本区分急诊/常规(&E)"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   12
            Top             =   543
            Width           =   2475
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "审核后自动打印(&P)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   11
            Top             =   270
            Width           =   1935
         End
         Begin VB.CheckBox chkCheckAll 
            Caption         =   "按仪器项目核收(&C)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   10
            Top             =   543
            Width           =   2355
         End
         Begin VB.Label lblNotice 
            Caption         =   "选中以下选项可能会使核收过程变慢："
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   90
            TabIndex        =   18
            Top             =   2940
            Width           =   4395
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "时间范围"
         Height          =   1110
         Left            =   120
         TabIndex        =   3
         Top             =   375
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "快速找到当前病人的历次检验的时间范围"
            Top             =   630
            Width           =   1920
         End
         Begin MSComCtl2.DTPicker DTPHisTory 
            Height          =   285
            Left            =   4080
            TabIndex        =   24
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   68812801
            CurrentDate     =   39475
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "快速找到当前病人的历次检验的时间范围"
            Top             =   270
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "标本序号生成规则(&2)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   345
            TabIndex        =   30
            Top             =   690
            Width           =   1710
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "历次检验范围(&1)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   705
            TabIndex        =   20
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label9 
            Caption         =   "在检验技师工作站中的待核收、在检验以及已完成的时间范围分别按如下设置进行搜索。"
            Height          =   15
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   4065
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   135
            Picture         =   "frmLisStationPara.frx":007C
            Top             =   105
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "历史比较病人识别方式"
         Height          =   765
         Left            =   105
         TabIndex        =   5
         Top             =   1530
         Width           =   5475
         Begin VB.OptionButton OptHistoryName 
            Caption         =   "病人姓名"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2370
            TabIndex        =   8
            Top             =   330
            Width           =   1845
         End
         Begin VB.OptionButton optHistoryID 
            Caption         =   "病人ID"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   330
            Width           =   1455
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   105
            Picture         =   "frmLisStationPara.frx":0946
            Top             =   210
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   285
         Left            =   -74670
         TabIndex        =   36
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68812803
         CurrentDate     =   38792
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   -72900
         TabIndex        =   42
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68812803
         CurrentDate     =   38792
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   5625
         Left            =   -74850
         TabIndex        =   49
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9922
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "编码"
            Object.Width           =   2277
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   4235
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "只打印指定申请科室的报告单"
         Height          =   195
         Left            =   -74820
         TabIndex        =   50
         Top             =   420
         Width           =   3525
      End
      Begin VB.Label lblNotify 
         AutoSize        =   -1  'True
         Caption         =   "    某些仪器不能直接从串口读取数据，必须从其产生的特定文件中提取数据。"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   -74730
         TabIndex        =   46
         Top             =   660
         Width           =   5205
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "仪器数据文件(&F)"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -74730
         TabIndex        =   45
         Top             =   1290
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "检验仪器(&Y)"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -74700
         TabIndex        =   44
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label13 
         Caption         =   "～"
         Height          =   165
         Left            =   -73140
         TabIndex        =   43
         Top             =   2850
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   6495
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   6495
      Width           =   1100
   End
End
Attribute VB_Name = "frmLisStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object
Private mstrPrivs As String                                         '权限

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lng药房ID As Long
    Dim strsql As String, strPar As String, i As Long
    Dim bln参数设置 As Boolean
    Dim strMachine As String
    Dim strDepts As String
    Dim lItem As ListItem
    Dim int病人信息处理 As Integer
    
    mblnOK = False
    mstrPrivs = gstrPrivs
    
'    If InStr(mstrPrivs, "参数设置") <= 0 Then
'        Me.chkSample.Enabled = False
'        Me.chkPatientType.Enabled = False
'        Me.chkNotSend.Enabled = False
'    End If
    bln参数设置 = InStr(";" & mstrPrivs & ";", ";参数设置;")
    Set mfrmMain = frmMain
    '初始化
    
    For mlngLoop = 2 To 2
        cbo(mlngLoop).AddItem "今  天"
        cbo(mlngLoop).AddItem "昨  天"
        cbo(mlngLoop).AddItem "本  周"
        cbo(mlngLoop).AddItem "本  月"
        cbo(mlngLoop).AddItem "本  季"
        cbo(mlngLoop).AddItem "本半年"
        cbo(mlngLoop).AddItem "本  年"
        cbo(mlngLoop).AddItem "前三天"
        cbo(mlngLoop).AddItem "前一周"
        cbo(mlngLoop).AddItem "前半月"
        cbo(mlngLoop).AddItem "前一月"
        cbo(mlngLoop).AddItem "前二月"
        cbo(mlngLoop).AddItem "前三月"
        cbo(mlngLoop).AddItem "前半年"
        cbo(mlngLoop).AddItem "自定义"
    Next
    
    cbo(2).AddItem "指定开始日期"
    
    cbo(3).AddItem "今  天"
    cbo(3).AddItem "本  周"
    cbo(3).AddItem "本  月"
    cbo(3).AddItem "本  年"
    cbo(3).AddItem "不重复"
    
    On Error Resume Next
    chkSample.Value = zlDatabase.GetPara("登记时可直接输入病人信息", 100, 1208, 0, Array(chkSample), bln参数设置)
    chkPrint.Value = zlDatabase.GetPara("审核打印", 100, 1208, 0, Array(chkPrint), bln参数设置)
    chkShowAll.Value = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0, Array(chkShowAll), bln参数设置)
    ChkPrivacy.Value = zlDatabase.GetPara("报告单是否显示隐私项目", 100, 1208, 0, Array(ChkPrivacy), bln参数设置)
    chkPatientType.Value = zlDatabase.GetPara("所有登记病人标识为外来", 100, 1208, 0, Array(chkPatientType), bln参数设置)
    chkNotSend.Value = zlDatabase.GetPara("使用二级报告审核", 100, 1208, 0, Array(chkNotSend), bln参数设置)
    
    checkSaveInfoSign.Value = zlDatabase.GetPara("核收登记保存时签名", 100, 1208, 0, Array(checkSaveInfoSign), bln参数设置)
    checkSaveReprotSign.Value = zlDatabase.GetPara("报告单保存时签名", 100, 1208, 0, Array(checkSaveReprotSign), bln参数设置)
    
    
    
    
    int病人信息处理 = Val(zlDatabase.GetPara("门诊病人信息不一致的处理方式", 100, 1208, 1, Array(opt门诊处理(0), opt门诊处理(1), opt门诊处理(2)), bln参数设置))
    opt门诊处理(0).Value = int病人信息处理 = 1
    opt门诊处理(1).Value = int病人信息处理 = 2
    opt门诊处理(2).Value = int病人信息处理 = 3
    
    int病人信息处理 = Val(zlDatabase.GetPara("住院病人信息不一致的处理方式", 100, 1208, 1, Array(opt住院处理(0), opt住院处理(1), opt住院处理(2)), bln参数设置))
    opt住院处理(0).Value = int病人信息处理 = 1
    opt住院处理(1).Value = int病人信息处理 = 2
    opt住院处理(2).Value = int病人信息处理 = 3
    
    int病人信息处理 = Val(zlDatabase.GetPara("院外病人信息不一致的处理方式", 100, 1208, 1, Array(opt院外处理(0), opt院外处理(1), opt院外处理(2)), bln参数设置))
    opt院外处理(0).Value = int病人信息处理 = 1
    opt院外处理(1).Value = int病人信息处理 = 2
    opt院外处理(2).Value = int病人信息处理 = 3
    
    int病人信息处理 = Val(zlDatabase.GetPara("体检病人信息不一致的处理方式", 100, 1208, 1, Array(opt体检处理(0), opt体检处理(1), opt体检处理(2)), bln参数设置))
    opt体检处理(0).Value = int病人信息处理 = 1
    opt体检处理(1).Value = int病人信息处理 = 2
    opt体检处理(2).Value = int病人信息处理 = 3
    
    
    cbo(2).Text = zlDatabase.GetPara("历次检验范围", 100, 1208, "本  月", Array(cbo(2)), bln参数设置)
    cbo(3).Text = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天", Array(cbo(3)), bln参数设置)
    Me.DTPHisTory.Value = zlDatabase.GetPara("历次检验范围指定开始日期", 100, 1208, Format(Now - 30, "yyyy-mm-dd"), Array(Me.DTPHisTory), bln参数设置)
    Me.DTPHisTory.Visible = (cbo(2).Text = "指定开始日期")
    chkAutoRefresh.Value = zlDatabase.GetPara("自动刷新", 100, 1208, 1, Array(chkAutoRefresh), bln参数设置)
    chkNoRange.Value = zlDatabase.GetPara("核收忽略时间", 100, 1208, 1, Array(chkNoRange), bln参数设置)
    chkCheck.Value = zlDatabase.GetPara("核收显示收费", 100, 1208, 1, Array(chkCheck), bln参数设置)
    chkComm.Value = zlDatabase.GetPara("核收允许双向", 100, 1208, 0, Array(chkComm), bln参数设置)
    chkEmerge.Value = zlDatabase.GetPara("急诊标本", 100, 1208, 0, Array(chkEmerge), bln参数设置)
    chkCheckAll.Value = zlDatabase.GetPara("按仪器项目核收", 100, 1208, 0, Array(chkCheckAll), bln参数设置)
    chkShowType.Value = zlDatabase.GetPara("自适应显示结果", 100, 1208, 0, Array(chkShowType), bln参数设置)
    chkNO.Value = zlDatabase.GetPara("按上次输入的标本号累加", 100, 1208, 0, Array(chkNO), bln参数设置)
    chkShowOption.Value = zlDatabase.GetPara("只在核收登记时显示登记窗口", 100, 1208, 0, Array(chkShowOption), bln参数设置)
    ChkCheckInNoItem.Value = zlDatabase.GetPara("登记时不需要输入项目", 100, 1208, 0, Array(ChkCheckInNoItem), bln参数设置)
    chkItemNumber.Value = zlDatabase.GetPara("手工项目按项目累加标本号", 100, 1208, 0, Array(chkItemNumber), bln参数设置)
    chkSkipRule.Value = zlDatabase.GetPara("审核后跳到下一个可审标本", 100, 1208, 0, Array(chkSkipRule), bln参数设置)
    chkOnlyMachine.Value = zlDatabase.GetPara("只核收当前仪器项目", 100, 1208, 0, Array(chkOnlyMachine), bln参数设置)
    chkLast.Value = zlDatabase.GetPara("核收时提示上次超标结果", 100, 1208, 0, Array(chkLast), bln参数设置)
    chkLoadLast.Value = zlDatabase.GetPara("登记时保留上一次申请项目", 100, 1208, 0, Array(chkLoadLast), bln参数设置)
    chkAutoAddItem.Value = zlDatabase.GetPara("自动增加计算项目", 100, 1208, 1, Array(chkAutoAddItem), bln参数设置)
    chkSampleType.Value = zlDatabase.GetPara("上次结果不参照标本类型", 100, 1208, 0, Array(chkSampleType), bln参数设置)
    
    i = zlDatabase.GetPara("历史病人识别", 100, 1208, 0, Array(optHistoryID), bln参数设置)
    i = zlDatabase.GetPara("历史病人识别", 100, 1208, 0, Array(OptHistoryName), bln参数设置)
    If i = 0 Then
        Me.optHistoryID.Value = True
    Else
        Me.OptHistoryName.Value = True
    End If
    
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    If cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    
    '初始文件提取参数
    On Error GoTo DBError
    
    strsql = "Select " & gConst_检验仪器_列名 & " From 检验仪器 a"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem "(" & rsTmp("编码") & ")" & rsTmp("名称")
        cboDevice.ItemData(cboDevice.ListCount - 1) = rsTmp("ID")
        
        rsTmp.MoveNext
    Loop
    
    txtFile = zlDatabase.GetPara("仪器数据文件", 100, 1208, "", Array(txtFile, cmdFile), bln参数设置)
    strMachine = zlDatabase.GetPara("文件提取仪器", 100, 1208, "", Array(cboDevice), bln参数设置)

    On Error Resume Next
    If strMachine <> "" And cboDevice.Enabled = True Then
        cboDevice.ListIndex = GetComboxIndex(cboDevice, strMachine)
    End If
    i = zlDatabase.GetPara("文件提取范围", 100, 1208, 0, Array(optRange), bln参数设置)
    optRange(i).Value = True
    If i = 0 Then '提取当天
        dtpStart = zlDatabase.Currentdate: dtpEnd = zlDatabase.Currentdate
    Else
        dtpStart = CDate(zlDatabase.GetPara("文件提取开始日期", 100, 1208, zlDatabase.Currentdate, Array(dtpStart), bln参数设置))
        dtpEnd = CDate(zlDatabase.GetPara("文件提取结束日期", 100, 1208, zlDatabase.Currentdate, Array(dtpEnd), bln参数设置))
    End If
    
    '处理那些申请科室的报告单可以打印
    strDepts = zlDatabase.GetPara("只打指定科室报告单", 100, 1208, "", Array(lvwDept), bln参数设置)
    gstrSql = "Select a.id,a.编码,a.名称 From 部门表 A, 部门性质说明 B Where A.ID = B.部门id And B.工作性质 In ('临床', '护理','检验','体检') " & _
            " order by a.编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.lvwDept
        Do While Not rsTmp.EOF
            Set lItem = .ListItems.Add(1, "A" & rsTmp("id"), rsTmp("编码"))
            lItem.SubItems(1) = rsTmp("名称")
            If InStr("," & strDepts & ",", "," & rsTmp("id") & ",") > 0 Then
                lItem.Checked = True
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    If strDepts = "" Then
        Call cmdSelectAll_Click
    End If
    
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo_Click(Index As Integer)
    
    Me.DTPHisTory.Visible = (cbo(2).Text = "指定开始日期")
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDevice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门成药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门西药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门中药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住成药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住西药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住中药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub chkActLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkFinish_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkAutoRefresh_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkComm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkNoRange_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkShort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk药库_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim intLoop As Integer
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            .ListItems(intLoop).Checked = False
        Next
    End With
End Sub

Private Sub cmdFile_Click()
    On Error GoTo OpenError
    With dlgFile
        .CancelError = True
        .DialogTitle = "请选择仪器数据文件"
        .ShowOpen
        txtFile = .FileName
    End With
    zlCommFun.PressKey vbKeyTab
    Exit Sub
OpenError:
    txtFile.SetFocus
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdIC_Click()
    Dim objIC As Object
    Set objIC = CreateObject("zlICCard.clsICCard")
    If Not objIC Is Nothing Then
        Call objIC.Set_Card
        Set objIC = Nothing
    End If
    
End Sub

Private Sub cmdIdent_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1101)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    Dim intLoop As Integer
    Dim strDepts As String
    
    zlDatabase.SetPara "登记时可直接输入病人信息", chkSample.Value, 100, 1208
    zlDatabase.SetPara "审核打印", chkPrint.Value, 100, 1208
    zlDatabase.SetPara "不区分仪器显示核收项目", chkShowAll.Value, 100, 1208
    zlDatabase.SetPara "报告单是否显示隐私项目", ChkPrivacy.Value, 100, 1208
    zlDatabase.SetPara "所有登记病人标识为外来", chkPatientType.Value, 100, 1208
    zlDatabase.SetPara "使用二级报告审核", chkNotSend.Value, 100, 1208
    zlDatabase.SetPara "审核后跳到下一个可审标本", Me.chkSkipRule.Value, 100, 1208
    
    zlDatabase.SetPara "核收登记保存时签名", checkSaveInfoSign.Value, 100, 1208
    zlDatabase.SetPara "报告单保存时签名", checkSaveReprotSign.Value, 100, 1208
        
    If opt门诊处理(0).Value = True Then zlDatabase.SetPara "门诊病人信息不一致的处理方式", 1, 100, 1208
    If opt门诊处理(1).Value = True Then zlDatabase.SetPara "门诊病人信息不一致的处理方式", 2, 100, 1208
    If opt门诊处理(2).Value = True Then zlDatabase.SetPara "门诊病人信息不一致的处理方式", 3, 100, 1208
    
    If opt住院处理(0).Value = True Then zlDatabase.SetPara "住院病人信息不一致的处理方式", 1, 100, 1208
    If opt住院处理(1).Value = True Then zlDatabase.SetPara "住院病人信息不一致的处理方式", 2, 100, 1208
    If opt住院处理(2).Value = True Then zlDatabase.SetPara "住院病人信息不一致的处理方式", 3, 100, 1208
    
    If opt院外处理(0).Value = True Then zlDatabase.SetPara "院外病人信息不一致的处理方式", 1, 100, 1208
    If opt院外处理(1).Value = True Then zlDatabase.SetPara "院外病人信息不一致的处理方式", 2, 100, 1208
    If opt院外处理(2).Value = True Then zlDatabase.SetPara "院外病人信息不一致的处理方式", 3, 100, 1208
    
    If opt体检处理(0).Value = True Then zlDatabase.SetPara "体检病人信息不一致的处理方式", 1, 100, 1208
    If opt体检处理(1).Value = True Then zlDatabase.SetPara "体检病人信息不一致的处理方式", 2, 100, 1208
    If opt体检处理(2).Value = True Then zlDatabase.SetPara "体检病人信息不一致的处理方式", 3, 100, 1208
    
    '----------------------------------------------------------------------------------------
    zlDatabase.SetPara "历次检验范围", cbo(2).Text, 100, 1208
    zlDatabase.SetPara "标本序号生成规则", cbo(3).Text, 100, 1208
    zlDatabase.SetPara "历次检验范围指定开始日期", Format(Me.DTPHisTory.Value, "yyyy-mm-dd 00:00:00"), 100, 1208
    zlDatabase.SetPara "自动刷新", chkAutoRefresh.Value, 100, 1208
    zlDatabase.SetPara "核收忽略时间", chkNoRange.Value, 100, 1208
    zlDatabase.SetPara "核收显示收费", chkCheck.Value, 100, 1208
    zlDatabase.SetPara "核收允许双向", chkComm.Value, 100, 1208
    zlDatabase.SetPara "急诊标本", chkEmerge.Value, 100, 1208
    zlDatabase.SetPara "按仪器项目核收", chkCheckAll.Value, 100, 1208
    zlDatabase.SetPara "历史病人识别", IIf(Me.optHistoryID.Value, 0, 1), 100, 1208
    zlDatabase.SetPara "自适应显示结果", chkShowType.Value, 100, 1208
    zlDatabase.SetPara "按上次输入的标本号累加", chkNO.Value, 100, 1208
    zlDatabase.SetPara "只在核收登记时显示登记窗口", chkShowOption.Value, 100, 1208
    zlDatabase.SetPara "登记时不需要输入项目", ChkCheckInNoItem.Value, 100, 1208
    zlDatabase.SetPara "手工项目按项目累加标本号", Me.chkItemNumber.Value, 100, 1208
    zlDatabase.SetPara "只核收当前仪器项目", Me.chkOnlyMachine.Value, 100, 1208
    zlDatabase.SetPara "核收时提示上次超标结果", Me.chkLast.Value, 100, 1208
    zlDatabase.SetPara "登记时保留上一次申请项目", Me.chkLoadLast.Value, 100, 1208
    zlDatabase.SetPara "自动增加计算项目", Me.chkAutoAddItem.Value, 100, 1208
    zlDatabase.SetPara "上次结果不参照标本类型", Me.chkSampleType.Value, 100, 1208
    
    If Len(txtFile) > 0 Then
        zlDatabase.SetPara "仪器数据文件", txtFile, 100, 1208
    End If
    If cboDevice.ListIndex > -1 Then
        zlDatabase.SetPara "文件提取仪器", cboDevice.ItemData(cboDevice.ListIndex), 100, 1208
    End If
    zlDatabase.SetPara "文件提取范围", IIf(optRange(0).Value, 0, 1), 100, 1208
    zlDatabase.SetPara "文件提取开始日期", Format(dtpStart, "yyyy-MM-dd"), 100, 1208
    zlDatabase.SetPara "文件提取结束日期", Format(dtpEnd, "yyyy-MM-dd"), 100, 1208
    '----------------------------------------------------------------------------------------
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            If .ListItems(intLoop).Checked = True Then
                strDepts = strDepts & "," & Mid(.ListItems(intLoop).Key, 2)
            End If
        Next
        Call zlDatabase.SetPara("只打指定科室报告单", strDepts, 100, 1208)
    End With
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim intLoop As Integer
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            .ListItems(intLoop).Checked = True
        Next
    End With
End Sub

Private Sub dtpEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_LostFocus()
    If dtpEnd < dtpStart Then dtpEnd = dtpStart
End Sub

Private Sub dtpStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpStart_LostFocus()
    If dtpStart > dtpEnd Then dtpStart = dtpEnd
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
'    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
'        lst收费类别.Selected(Item) = True
'    End If
End Sub

Private Sub lst收费类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRange_Click(Index As Integer)
    If Index = 0 Then
        dtpStart.Enabled = False
        dtpEnd.Enabled = False
    Else
        dtpStart.Enabled = True
        dtpEnd.Enabled = True
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub optRange_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt药品单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
'    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tbs.Tab = 1
'        cbo门西药.SetFocus
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtFile_GotFocus()
    With txtFile
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

