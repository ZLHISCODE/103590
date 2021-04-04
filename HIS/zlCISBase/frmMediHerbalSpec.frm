VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediHerbalSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "草药规格编辑"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "frmMediHerbalSpec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "保存后新增规格(&B)"
      Height          =   350
      Left            =   4800
      TabIndex        =   110
      Top             =   5805
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "保存后新增品种(&A)"
      Height          =   350
      Left            =   2880
      TabIndex        =   109
      Top             =   5805
      Width           =   1695
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frmMediHerbalSpec.frx":08CA
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5805
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8160
      TabIndex        =   87
      Top             =   5805
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存退出(&O)"
      Height          =   350
      Left            =   6720
      TabIndex        =   86
      Top             =   5805
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   90
      Top             =   330
      Width           =   9405
   End
   Begin VB.PictureBox picFound 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3900
      ScaleHeight     =   210
      ScaleWidth      =   5505
      TabIndex        =   88
      Top             =   480
      Width           =   5505
      Begin VB.Label lblFound 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "注：该规格建立于2002年12月20日，于2003年8月10日停用。"
         Height          =   180
         Left            =   630
         TabIndex        =   89
         Top             =   0
         Width           =   4770
      End
   End
   Begin TabDlg.SSTab stbSpec 
      Height          =   5115
      Left            =   120
      TabIndex        =   106
      Top             =   525
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "规格信息(&1)"
      TabPicture(0)   =   "frmMediHerbalSpec.frx":0A14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStationNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl备选码"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl发药类型"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl说明"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblComment"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl合同单位"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "申领单位"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl申领单位"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl注册商标"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl批准文号"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl药品来源"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl药库单位Child"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl药房单位Child"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl售价单位Child"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl数字码"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl药库包装"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl药库单位"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl药房包装"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl药房单位"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl剂量系数"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl售价单位"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl编码"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl规格"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl产地"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl标识码"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl中药形态"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl送货单位child"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl送货包装"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl送货单位"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl本位码"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl原产地"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txt合同单位"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmbStationNo"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt备选码"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cbo发药类型"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt说明"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmd合同单位"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cbo申领单位"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt申领阀值"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt注册商标"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt批准文号"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cbo药品来源"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt编码"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt数字码"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt售价单位"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt药库包装"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt药库单位"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt药房包装"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt药房单位"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt剂量系数"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt规格"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txt产地"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt标识码"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "opt中药形态(0)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "opt中药形态(1)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "opt中药形态(2)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cmd产地"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt送货包装"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt送货单位"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt本位码"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmd原产地"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt原产地"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).ControlCount=   63
      TabCaption(1)   =   "药价信息(&2)"
      TabPicture(1)   =   "frmMediHerbalSpec.frx":0A30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk零差价"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chk摆药"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbo门诊可否分零"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd病案"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chk非常备药"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chk屏蔽费别"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt指导售价"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt指导批价"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cbo药价级别"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cbo收入记入"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txt当前售价"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cbo费用类型"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cbo服务对象"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cbo住院可否分零"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txt扣率"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txt结算价"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cbo药价属性"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fra分批核算"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkGMP认证"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txt加成率"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txt成本价格"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chk住院动态分零"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txt病案费目"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lbl门诊可否分零"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lbl病案费目"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lbl指导售价"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lbl指导批价"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "lblPercent(0)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "lbl扣率"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lbl结算价"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lbl药价属性"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lbl药价级别"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lbl当前售价"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "lbl收入记入"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lbl加成率"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lbl成本价格"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lbl费用类型"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lbl住院可否分零"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "lbl服务对象"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "lbl批价单位(0)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "lbl批价单位(1)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).ControlCount=   41
      TabCaption(2)   =   "扩展属性(&3)"
      TabPicture(2)   =   "frmMediHerbalSpec.frx":0A4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsfItem"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt原产地 
         Height          =   300
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1882
         Width           =   3045
      End
      Begin VB.CommandButton cmd原产地 
         Caption         =   "…"
         Height          =   285
         Left            =   4150
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1890
         Width           =   285
      End
      Begin VB.TextBox txt本位码 
         Height          =   300
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   2
         Top             =   750
         Width           =   1995
      End
      Begin VB.CheckBox chk零差价 
         Caption         =   "启用零差价管理模式"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chk摆药 
         Caption         =   "摆药"
         Height          =   180
         Left            =   -67560
         TabIndex        =   82
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.TextBox txt送货单位 
         Height          =   300
         Left            =   5910
         MaxLength       =   8
         TabIndex        =   30
         Text            =   "袋"
         Top             =   2640
         Width           =   585
      End
      Begin VB.TextBox txt送货包装 
         Height          =   300
         Left            =   7125
         MaxLength       =   10
         TabIndex        =   31
         Text            =   "1000"
         Top             =   2640
         Width           =   945
      End
      Begin VB.ComboBox cbo门诊可否分零 
         Height          =   300
         Left            =   -67545
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   832
         Width           =   1725
      End
      Begin VB.CommandButton cmd产地 
         Caption         =   "…"
         Height          =   285
         Left            =   4150
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   1508
         Width           =   285
      End
      Begin VB.CommandButton cmd病案 
         Caption         =   "…"
         Height          =   240
         Left            =   -69585
         TabIndex        =   107
         TabStop         =   0   'False
         Tag             =   "分类"
         ToolTipText     =   "按*打开选择器"
         Top             =   862
         Width           =   255
      End
      Begin VB.CheckBox chk非常备药 
         Caption         =   "非常备药"
         Height          =   180
         Left            =   -69000
         TabIndex        =   81
         Top             =   1680
         Width           =   1080
      End
      Begin VB.CheckBox chk屏蔽费别 
         Alignment       =   1  'Right Justify
         Caption         =   "屏蔽费别(&M)"
         Height          =   285
         Left            =   -71820
         TabIndex        =   74
         Top             =   2456
         Width           =   1290
      End
      Begin VB.OptionButton opt中药形态 
         Caption         =   "免煎剂"
         Height          =   180
         Index           =   2
         Left            =   7850
         TabIndex        =   34
         Top             =   3075
         Width           =   855
      End
      Begin VB.OptionButton opt中药形态 
         Caption         =   "中药饮片"
         Height          =   180
         Index           =   1
         Left            =   6700
         TabIndex        =   33
         Top             =   3075
         Width           =   1020
      End
      Begin VB.OptionButton opt中药形态 
         Caption         =   "散装"
         Height          =   180
         Index           =   0
         Left            =   5900
         TabIndex        =   32
         Top             =   3075
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txt标识码 
         Height          =   300
         Left            =   3150
         MaxLength       =   29
         TabIndex        =   11
         Top             =   2265
         Width           =   1275
      End
      Begin VB.TextBox txt产地 
         Height          =   300
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1503
         Width           =   3045
      End
      Begin VB.TextBox txt规格 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1125
         Width           =   3285
      End
      Begin VB.TextBox txt剂量系数 
         Height          =   300
         Left            =   7125
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "1"
         Top             =   1155
         Width           =   945
      End
      Begin VB.TextBox txt药房单位 
         Height          =   300
         Left            =   5910
         MaxLength       =   8
         TabIndex        =   24
         Text            =   "g"
         Top             =   1515
         Width           =   585
      End
      Begin VB.TextBox txt药房包装 
         Height          =   300
         Left            =   7125
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "1"
         Top             =   1515
         Width           =   945
      End
      Begin VB.TextBox txt药库单位 
         Height          =   300
         Left            =   5910
         MaxLength       =   8
         TabIndex        =   26
         Text            =   "g"
         Top             =   1890
         Width           =   585
      End
      Begin VB.TextBox txt药库包装 
         Height          =   300
         Left            =   7125
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "1"
         Top             =   1890
         Width           =   945
      End
      Begin VB.TextBox txt指导售价 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   61
         Top             =   3252
         Width           =   1470
      End
      Begin VB.TextBox txt指导批价 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   55
         Top             =   2046
         Width           =   1470
      End
      Begin VB.ComboBox cbo药价级别 
         Height          =   300
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1242
         Width           =   1485
      End
      Begin VB.ComboBox cbo收入记入 
         Height          =   300
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox txt当前售价 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   53
         Top             =   1644
         Width           =   1485
      End
      Begin VB.ComboBox cbo费用类型 
         Height          =   300
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1644
         Width           =   1485
      End
      Begin VB.ComboBox cbo服务对象 
         Height          =   300
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   2046
         Width           =   1485
      End
      Begin VB.TextBox txt售价单位 
         Height          =   300
         Left            =   5910
         MaxLength       =   8
         TabIndex        =   22
         Text            =   "g"
         Top             =   1170
         Width           =   585
      End
      Begin VB.TextBox txt数字码 
         Height          =   300
         Left            =   1140
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2265
         Width           =   1020
      End
      Begin VB.ComboBox cbo住院可否分零 
         Height          =   300
         Left            =   -67560
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   450
         Width           =   1725
      End
      Begin VB.TextBox txt扣率 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   57
         Text            =   "100"
         Top             =   2448
         Width           =   1470
      End
      Begin VB.TextBox txt结算价 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   59
         Top             =   2850
         Width           =   1470
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   1140
         MaxLength       =   13
         TabIndex        =   1
         Top             =   375
         Width           =   1995
      End
      Begin VB.ComboBox cbo药价属性 
         Height          =   300
         Left            =   -73860
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   450
         Width           =   1470
      End
      Begin VB.Frame fra分批核算 
         Caption         =   "分批管理(&K)"
         Height          =   1065
         Left            =   -69000
         TabIndex        =   83
         Top             =   2160
         Width           =   2520
         Begin VB.CheckBox chk药房 
            Caption         =   "药房"
            Enabled         =   0   'False
            Height          =   210
            Left            =   1470
            TabIndex        =   85
            Top             =   300
            Width           =   675
         End
         Begin VB.CheckBox chk药库 
            Caption         =   "药库"
            Height          =   210
            Left            =   330
            TabIndex        =   84
            Top             =   300
            Width           =   675
         End
      End
      Begin VB.ComboBox cbo药品来源 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2655
         Width           =   3300
      End
      Begin VB.TextBox txt批准文号 
         Height          =   300
         Left            =   5910
         MaxLength       =   40
         TabIndex        =   20
         Top             =   375
         Width           =   2400
      End
      Begin VB.CheckBox chkGMP认证 
         Caption         =   "GMP认证(&Z)"
         Height          =   180
         Left            =   -67560
         TabIndex        =   80
         Top             =   1320
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txt加成率 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   63
         Text            =   "15.00"
         Top             =   3654
         Width           =   1470
      End
      Begin VB.TextBox txt成本价格 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   51
         Top             =   1242
         Width           =   1485
      End
      Begin VB.TextBox txt注册商标 
         Height          =   300
         Left            =   5910
         MaxLength       =   50
         TabIndex        =   21
         Top             =   765
         Width           =   2400
      End
      Begin VB.TextBox txt申领阀值 
         Height          =   300
         Left            =   7470
         MaxLength       =   8
         TabIndex        =   29
         Top             =   2265
         Width           =   585
      End
      Begin VB.ComboBox cbo申领单位 
         Height          =   300
         Left            =   5910
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2265
         Width           =   1320
      End
      Begin VB.CommandButton cmd合同单位 
         Caption         =   "…"
         Height          =   285
         Left            =   4140
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3045
         Width           =   285
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   17
         Top             =   3765
         Width           =   3285
      End
      Begin VB.CheckBox chk住院动态分零 
         Caption         =   "住院动态分零"
         Height          =   180
         Left            =   -69000
         TabIndex        =   79
         Top             =   1320
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox cbo发药类型 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4125
         Width           =   3300
      End
      Begin VB.TextBox txt备选码 
         Height          =   300
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   35
         Top             =   3765
         Width           =   2400
      End
      Begin VB.ComboBox cmbStationNo 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4500
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.TextBox txt病案费目 
         Height          =   300
         Left            =   -70800
         MaxLength       =   40
         TabIndex        =   67
         ToolTipText     =   "按*打开选择器"
         Top             =   832
         Width           =   1485
      End
      Begin VB.TextBox txt合同单位 
         Height          =   300
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   15
         Top             =   3045
         Width           =   3045
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfItem 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   114
         Top             =   480
         Width           =   9075
         _cx             =   16007
         _cy             =   6800
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMediHerbalSpec.frx":0A68
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
         Editable        =   2
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
      Begin VB.Label lbl原产地 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "原产地"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   117
         Top             =   1942
         Width           =   540
      End
      Begin VB.Label lbl本位码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "本位码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   115
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lbl送货单位 
         AutoSize        =   -1  'True
         Caption         =   "送货单位(&V)"
         Height          =   180
         Left            =   4800
         TabIndex        =   113
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label lbl送货包装 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1袋="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6675
         TabIndex        =   112
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label lbl送货单位child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   8100
         TabIndex        =   111
         Top             =   2700
         Width           =   180
      End
      Begin VB.Label lbl门诊可否分零 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "门诊分零使用(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -69000
         TabIndex        =   77
         Top             =   892
         Width           =   1350
      End
      Begin VB.Label lbl病案费目 
         Caption         =   "病案费目(&O)"
         Height          =   255
         Left            =   -71820
         TabIndex        =   66
         Top             =   855
         Width           =   990
      End
      Begin VB.Label lbl中药形态 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中药形态(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   45
         Top             =   3075
         Width           =   990
      End
      Begin VB.Label lbl标识码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标识码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2520
         TabIndex        =   10
         Top             =   2325
         Width           =   540
      End
      Begin VB.Label lbl产地 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "生产商(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   5
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lbl规格 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品规格(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   3
         Top             =   1185
         Width           =   990
      End
      Begin VB.Label lbl编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "规格编码(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   0
         Top             =   435
         Width           =   990
      End
      Begin VB.Label lbl售价单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "售价单位(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   41
         Top             =   1215
         Width           =   990
      End
      Begin VB.Label lbl剂量系数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1g="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6765
         TabIndex        =   103
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label lbl药房单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药房单位(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   42
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl药房包装 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1g="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6765
         TabIndex        =   102
         Top             =   1575
         Width           =   360
      End
      Begin VB.Label lbl药库单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药库单位(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   43
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl药库包装 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1g="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6765
         TabIndex        =   101
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label lbl指导售价 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "指导售价(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   60
         Top             =   3312
         Width           =   990
      End
      Begin VB.Label lbl指导批价 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "采购限价"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   54
         Top             =   2106
         Width           =   720
      End
      Begin VB.Label lbl数字码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "数字码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   8
         Top             =   2325
         Width           =   540
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
         Left            =   -72315
         TabIndex        =   100
         Top             =   2508
         Width           =   90
      End
      Begin VB.Label lbl扣率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "采购扣率(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   56
         Top             =   2508
         Width           =   990
      End
      Begin VB.Label lbl结算价 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结算价(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   58
         Top             =   2865
         Width           =   810
      End
      Begin VB.Label lbl售价单位Child 
         AutoSize        =   -1  'True
         Caption         =   "mg)"
         Height          =   180
         Left            =   8100
         TabIndex        =   97
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label lbl药房单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   8100
         TabIndex        =   96
         Top             =   1575
         Width           =   180
      End
      Begin VB.Label lbl药库单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   8100
         TabIndex        =   95
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label lbl药价属性 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药价属性(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   47
         Top             =   525
         Width           =   990
      End
      Begin VB.Label lbl药品来源 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "来源分类(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   2715
         Width           =   990
      End
      Begin VB.Label lbl批准文号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "批准文号(&W)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   39
         Top             =   435
         Width           =   990
      End
      Begin VB.Label lbl药价级别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药价级别(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   68
         Top             =   1302
         Width           =   990
      End
      Begin VB.Label lbl当前售价 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "当前售价(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   52
         Top             =   1704
         Width           =   990
      End
      Begin VB.Label lbl收入记入 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收入项目(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   64
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lbl加成率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "加成率"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   62
         Top             =   3714
         Width           =   540
      End
      Begin VB.Label lbl成本价格 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "成本价格(&C)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   50
         Top             =   1302
         Width           =   990
      End
      Begin VB.Label lbl注册商标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "注册商标(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   40
         Top             =   825
         Width           =   990
      End
      Begin VB.Label lbl费用类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保类型(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   70
         Top             =   1704
         Width           =   990
      End
      Begin VB.Label lbl住院可否分零 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院分零使用(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -69000
         TabIndex        =   75
         Top             =   510
         Width           =   1350
      End
      Begin VB.Label lbl服务对象 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "服务对象(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   72
         Top             =   2106
         Width           =   990
      End
      Begin VB.Label lbl申领单位 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   8100
         TabIndex        =   94
         Top             =   2325
         Width           =   180
      End
      Begin VB.Label 申领单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申领单位(&W)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4800
         TabIndex        =   44
         Top             =   2325
         Width           =   990
      End
      Begin VB.Label lbl合同单位 
         AutoSize        =   -1  'True
         Caption         =   "合同单位(&C)"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   3090
         Width           =   990
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "(指定了合同单位，药品就只能按合同单位入库。)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   93
         Top             =   3435
         Width           =   3960
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         Caption         =   "标识说明(&X)"
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   3810
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(请填写适当的说明，来表示限用、适用症药品。)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4800
         TabIndex        =   92
         Top             =   3450
         Width           =   3960
      End
      Begin VB.Label lbl发药类型 
         AutoSize        =   -1  'True
         Caption         =   "发药类型(&H)"
         Height          =   180
         Left            =   105
         TabIndex        =   37
         Top             =   4185
         Width           =   990
      End
      Begin VB.Label lbl备选码 
         AutoSize        =   -1  'True
         Caption         =   "备选码(&F)"
         Height          =   180
         Left            =   4800
         TabIndex        =   46
         Top             =   3825
         Width           =   810
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "院区编号(&Z)"
         Height          =   180
         Left            =   105
         TabIndex        =   38
         Top             =   4560
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl批价单位 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元/片"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   -72375
         TabIndex        =   99
         Top             =   2106
         Width           =   645
      End
      Begin VB.Label lbl批价单位 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元/片"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   -72375
         TabIndex        =   98
         Top             =   2985
         Width           =   645
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf合同单位 
      Height          =   1845
      Left            =   3720
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   6480
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
   Begin VB.Label lbl品种 
      AutoSize        =   -1  'True
      Caption         =   "药品编码：2010303   通用名称：头孢呋辛钠   剂型：片剂   剂量单位：mg"
      Height          =   180
      Left            =   165
      TabIndex        =   105
      Top             =   120
      Width           =   6120
   End
End
Attribute VB_Name = "frmMediHerbalSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前类别：由Me.tag存放，分别为"5"-西成药，"6"-中成药，根据lng药名ID查询确定
'   2、编辑状态：由Me.stbSpec.Tag存放，分别为"增加"、"修改"、"查阅"，由上级程序传递进入
'---------------------------------------------------
Public lng药名id As Long        '当前规格所属药品品种，由外部程序传递进入；根据品种确定类别等
Public lng药品ID As Long        '修改和、查询时由外部程序传递进入；增加时若不为0，表示根据该规格复制增加新的规格
Public strPrivs As String       '当前用户具有的本程序权限
Public mlng分类id As Long      '记录传过来的分类id

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mblnUsed As Boolean         '是否已使用
Dim blnActivate As Boolean
Private mblnLoad As Boolean      '只让active一次
Private mstr所有记录 As String      '记录所有界面的值
Private mblnOK As Boolean       '记录确定按钮是否被点击了
Private mblnCancel As Boolean   '记录取消按钮是否被点击了
Private mint分段加成 As Integer '用来获取系统参数中，是否勾选了时价药品按分段加成入库 0-未勾选，1-勾选
Private mrs分段加成 As ADODB.Recordset '用来记录分段加成率入库
Private mblnOtherSave As Boolean    '其他保存按钮被点击了
Private mintSet分批 As Integer  '库房分批设置 0-手工设置分批属性（默认值）；1-仅药库分批；2-药库和药房分批；3-药库和药房都不分批
Private mbln病案费目 As Boolean     '病案费目是否被点击了
Private mdbl加成率 As Double
Private mdbl差价额 As Double

'--协定药品与自制药品列常量--
Private mint招标药品 As Integer
Private Const col药品名称 As Integer = 1
Private Const col售价单位 As Integer = 2
Private Const col规格 As Integer = 3
Private Const col产地 As Integer = 4
Private Const col采用量 As Integer = 5
Private Const col剂量单位 As Integer = 6

'--储备限额列常量--
Private Const col库房 As Integer = 1
Private Const col上限 As Integer = 2
Private Const col下限 As Integer = 3
Private Const col日盘 As Integer = 4
Private Const col周盘 As Integer = 5
Private Const col月盘 As Integer = 6
Private Const col季盘 As Integer = 7
Private Const col货位 As Integer = 8

Private mlng编码长度 As Long
Private mlng规格长度 As Long
Private mlng产地长度 As Long
Private mlng原产地长度 As Long
Private mlng说明长度 As Long
Private mint简码长度 As Integer
Private mint备选码长度 As Integer

'从参数表中取药品价格小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数

Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer

Private mlngExpItemMaxLength As Long    '扩展项目内容的最大长度
Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSql = "Select A.编码, A.规格, A.说明, A.产地, B.简码, A.备选码, C.原产地 " & _
        " From 收费项目目录 A, 收费项目别名 B,药品规格 C " & _
        " Where A.ID = C.药品ID And A.ID = B.收费细目id And A.ID = 0 And B.码类 = 1 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng编码长度 = rsTmp.Fields("编码").DefinedSize
    mlng规格长度 = rsTmp.Fields("规格").DefinedSize
    mlng产地长度 = rsTmp.Fields("产地").DefinedSize
    mlng原产地长度 = rsTmp.Fields("原产地").DefinedSize
    mlng说明长度 = rsTmp.Fields("说明").DefinedSize
    mint简码长度 = rsTmp.Fields("简码").DefinedSize
    mint备选码长度 = rsTmp.Fields("备选码").DefinedSize
    
    txt编码.MaxLength = mlng编码长度
    txt规格.MaxLength = mlng规格长度
    txt产地.MaxLength = mlng产地长度
    txt原产地.MaxLength = mlng原产地长度
    txt说明.MaxLength = mlng说明长度
'    txt拼音.MaxLength = mint简码长度
'    txt五笔.MaxLength = mint简码长度
    txt备选码.MaxLength = mint备选码长度
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo发药类型_KeyPress(KeyAscii As Integer)
    'KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo费用类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo服务对象_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

'Private Sub cbo住院可否分零_Click()
'    If cbo住院可否分零.ListIndex = 0 Then
'        chk住院动态分零.Enabled = False
'    Else
'        chk住院动态分零.Enabled = True
'    End If
'End Sub

Private Sub cbo住院可否分零_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo门诊可否分零_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo申领单位_Click()
    Select Case cbo申领单位.ListIndex
    Case 0
        lbl申领单位.Caption = txt售价单位.Text & ")"
    Case 1
        lbl申领单位.Caption = txt药房单位.Text & ")"
    Case 2
        lbl申领单位.Caption = txt药库单位.Text & ")"
    End Select
End Sub

Private Sub cbo申领单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo收入记入_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo药价级别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

'Private Sub cbo药价属性_Click()
'    txt差价让利.Enabled = Not (cbo药价属性.Text = "定价")
'End Sub

Private Sub cbo药价属性_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo药品来源_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk摆药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk非常备药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk零差价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkGMP认证_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk基本药物_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub


Private Sub chk屏蔽费别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

'Private Sub chk效期_Click()
'    On Error Resume Next
'    Me.txt效期.Enabled = (chk效期.Value = 1)
'    If Me.txt效期.Enabled = False Then
'        Me.txt效期.Text = 0
'    Else
'        If Val(Me.txt效期.Text) = 0 Then Me.txt效期.Text = 24
'    End If
'    If Me.chk效期.Value = 1 Then Me.txt效期.SetFocus
'End Sub
'
'Private Sub chk效期_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        If Me.txt效期.Enabled = True Then
'            Me.txt效期.SetFocus
'        Else
'            Call OS.PressKey(vbKeyTab)
'        End If
'    End If
'End Sub

Private Sub chk药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk药库_Click()
    Dim blnEnable As Boolean
    
    '在药库分批的前提下，如果药房没有库存，则可设置其是否分批
    On Error GoTo ErrHandle
    gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
             " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
    
    With rsTemp
        blnEnable = True
        If .Fields(0).Value <> 0 Then
            blnEnable = False
        End If
    End With
    If Me.chk药库.Value = 0 Then
        Me.chk药房.Value = 0: Me.chk药房.Enabled = False
'        Me.chk效期.Value = 0: Me.chk效期.Enabled = False
'        Me.txt效期.Text = 0: Me.txt效期.Enabled = False
    Else
        Me.chk药房.Enabled = True
'        Me.chk效期.Enabled = True
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk药库_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk住院动态分零_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cmbStationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Dim strTemp As String
    
    If mblnOtherSave = False Then
        strTemp = txt编码.Text & "|" & txt本位码.Text & "|" & txt规格.Text & "|" & txt产地.Text & "|" & txt原产地.Text & "|" & _
                        txt数字码.Text & "|" & txt标识码.Text & "|" & cbo药品来源.Text & "|" & txt合同单位.Text & "|" & txt说明.Text & "|" & cbo发药类型.Text & "|" & _
                        cmbStationNo.Text & "|" & txt批准文号.Text & "|" & txt注册商标.Text & "|" & txt售价单位.Text & "|" & txt剂量系数.Text & "|" & _
                        txt药房单位.Text & "|" & txt药房包装.Text & "|" & txt药库单位.Text & "|" & txt药库包装.Text & "|" & cbo申领单位.Text & "|" & txt申领阀值.Text & "|" & _
                        txt备选码.Text & "|" & cbo药价属性.Text & "|" & txt成本价格.Text & "|" & txt当前售价.Text & "|" & txt指导批价.Text & "|" & txt扣率.Text & "|" & txt结算价.Text & "|" & _
                        txt指导售价.Text & "|" & txt加成率.Text & "|" & cbo收入记入.Text & "|" & txt病案费目.Text & "|" & cbo药价级别.Text & "|" & _
                        chk屏蔽费别.Value & "|" & cbo费用类型.Text & "|" & cbo服务对象.Text & "|" & cbo住院可否分零.Text & "|" & chk住院动态分零.Value & "|" & _
                        chkGMP认证.Value & "|" & chk非常备药.Value & "|" & chk药库.Value & "|" & chk药房.Value & "|" & cbo门诊可否分零.Text
        If strTemp <> mstr所有记录 Then
            mblnCancel = True
            If MsgBox("有数据被修改了确定退出？", vbYesNo, gstrSysName) = vbYes Then
                Unload Me
            Else
                mblnCancel = False
            End If
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim dbl当前售价 As Double, dbl指导售价 As Double, dbl成本价格 As Double
    Dim str站点 As String
    Dim blnTran As Boolean
    Dim strItems As String
    Dim n As Integer
    Dim rsPrice As New ADODB.Recordset
    
    mblnOK = True
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
        
    '检查规格页面的输入项是否正确
    strTemp = IIf(glngSys \ 100 <> 8, "药库", "采购")
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt编码.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt编码.Text, vbFromUnicode)) > mlng编码长度 Then MsgBox "编码超长(最多" & mlng编码长度 & "个字符)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt规格.Text) = "" Then MsgBox "请输入规格！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt规格.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt规格.Text, vbFromUnicode)) > mlng规格长度 Then MsgBox "规格超长(最多" & mlng规格长度 & "个字符或" & Int(mlng规格长度 / 2) & "个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt规格.SetFocus: Exit Sub
'    If LenB(StrConv(Me.txt商品名.Text, vbFromUnicode)) > 40 Then MsgBox "商品名超长(最多40个字符或20个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt商品名.SetFocus: Exit Sub
    
    If Trim(Me.txt售价单位.Text) = "" Then MsgBox "请输入售价单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt售价单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt售价单位.Text, vbFromUnicode)) > 8 Then MsgBox "售价单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt售价单位.SetFocus: Exit Sub
    If Val(Me.txt剂量系数.Text) = 0 Then MsgBox "剂量系数错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt剂量系数.SetFocus: Exit Sub
    If Val(Me.txt剂量系数.Text) >= 100000 Then MsgBox "剂量系数超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt剂量系数.SetFocus: Exit Sub
    
    If Trim(Me.txt药房单位.Text) = "" Then MsgBox "请输入药房单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt药房单位.Text, vbFromUnicode)) > 8 Then MsgBox "药房单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房单位.SetFocus: Exit Sub
    If Val(Me.txt药房包装.Text) = 0 Then MsgBox "药房包装错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房包装.SetFocus: Exit Sub
    If Val(Me.txt药房包装.Text) >= 100000 Then MsgBox "药房包装超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房包装.SetFocus: Exit Sub
    
'    If Trim(Me.txt住院单位.Text) = "" Then MsgBox "请输入住院单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt住院单位.SetFocus: Exit Sub
'    If LenB(StrConv(Me.txt住院单位.Text, vbFromUnicode)) > 8 Then MsgBox "住院单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt住院单位.SetFocus: Exit Sub
'    If Val(Me.txt住院包装.Text) = 0 Then MsgBox "住院包装错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt住院包装.SetFocus: Exit Sub
'    If Val(Me.txt住院包装.Text) >= 100000 Then MsgBox "住院包装超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt住院包装.SetFocus: Exit Sub
    
    If Trim(Me.txt药库单位.Text) = "" Then MsgBox "请输入" & strTemp & "单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt药库单位.Text, vbFromUnicode)) > 8 Then MsgBox strTemp & "单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库单位.SetFocus: Exit Sub
    If Val(Me.txt药库包装.Text) = 0 Then MsgBox strTemp & "包装错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库包装.SetFocus: Exit Sub
    If Val(Me.txt药库包装.Text) >= 100000 Then MsgBox strTemp & "包装超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库包装.SetFocus: Exit Sub
    If Trim(txt送货单位.Text) <> "" And Trim(txt送货包装.Text) = "" Then
        MsgBox "有送货单位情况下，送货包装不能为空！", vbInformation, gstrSysName
        txt送货包装.SetFocus
        txt送货包装.SelStart = 0
        txt送货包装.SelLength = 100
        Exit Sub
    End If
    If Trim(txt送货包装.Text) <> "" And IsNumeric(txt送货包装.Text) = False Then
        MsgBox "送货包装只能是数字，请重新输入！", vbInformation, gstrSysName
        txt送货包装.SetFocus
        txt送货包装.SelStart = 0
        txt送货包装.SelLength = 100
        Exit Sub
    End If
    
    If LenB(StrConv(Me.txt注册商标.Text, vbFromUnicode)) > 50 Then
        MsgBox "注册商标超长，最多50个字符或25个汉字！", vbInformation, gstrSysName
        Me.stbSpec.Tab = 0
        txt注册商标.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(Me.txt备选码.Text, vbFromUnicode)) > mint备选码长度 Then
        MsgBox "备选码超长(最多" & mint备选码长度 & "个字符)！", vbInformation, gstrSysName
        Me.stbSpec.Tab = 0
        txt备选码.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(Me.txt产地.Text, vbFromUnicode)) > mlng产地长度 Then MsgBox "生产商超长(最多" & mlng产地长度 & "个字符或" & Int(mlng产地长度 / 2) & "个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt产地.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt原产地.Text, vbFromUnicode)) > mlng原产地长度 Then MsgBox "原产地超长(最多" & mlng原产地长度 & "个字符或" & Int(mlng原产地长度 / 2) & "个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt原产地.SetFocus: Exit Sub
    
    If Val(Me.txt申领阀值.Text) < 0 Then MsgBox strTemp & "申领阀值不能小于零！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt申领阀值.SetFocus: Exit Sub
    If Val(Me.txt申领阀值.Text) >= 100000 Then MsgBox strTemp & "申领阀值超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt申领阀值.SetFocus: Exit Sub
    
    If Val(Me.txt指导批价.Text) = 0 And mblnUsed = True Then
        MsgBox "请输入" & IIf(mint招标药品 = 1, "中标价格", "指导批价") & "！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导批价.Enabled Then Me.txt指导批价.SetFocus: Exit Sub
    End If
    If Val(Me.txt指导批价.Text) > 1000000 Then
        MsgBox IIf(mint招标药品 = 1, "中标价格", "指导批价") & "超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导批价.Enabled Then Me.txt指导批价.SetFocus: Exit Sub
    End If
    If Val(Me.txt指导售价.Text) = 0 And mblnUsed = True Then
        MsgBox "请输入指导售价！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导售价.Enabled Then Me.txt指导售价.SetFocus: Exit Sub
    End If
    If Val(Me.txt指导售价.Text) > 1000000 Then
        MsgBox "指导售价超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导售价.Enabled Then Me.txt指导售价.SetFocus: Exit Sub
    End If
        
    If Val(Me.txt扣率.Text) = 0 Then MsgBox "请输入扣率！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt扣率.SetFocus: Exit Sub
    If Val(Me.txt扣率.Text) > 100 Then MsgBox "扣率超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt扣率.SetFocus: Exit Sub
    
    If Val(Me.txt加成率.Text) > 1000000 Then
        MsgBox "当前加成率超过最大值！", vbInformation, gstrSysName
        Me.stbSpec.Tab = 1
        If Me.txt加成率.Enabled Then Me.txt加成率.SetFocus
        Exit Sub
    End If
    
    If Me.cbo药价属性.ItemData(cbo药价属性.ListIndex) = 0 Then

        If Val(Me.txt当前售价.Text) > Val(Me.txt指导售价.Text) Then
            If MsgBox("售价高于指导零售价。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Me.stbSpec.Tab = 1
                If Me.txt当前售价.Enabled Then Me.txt当前售价.SetFocus
                Exit Sub
            End If
        End If
        If Val(Me.txt当前售价.Text) > 1000000 Then
            MsgBox "当前售价超过最大值！", vbInformation, gstrSysName
            Me.stbSpec.Tab = 1
            If Me.txt当前售价.Enabled Then Me.txt当前售价.SetFocus
            Exit Sub
        End If
    End If
    
    '检查扩展项目内容长度
    If stbSpec.TabVisible(2) = True Then
        With vsfItem
            For n = 1 To .Rows - 1
                If LenB(StrConv(Trim(.TextMatrix(n, .ColIndex("内容"))), vbFromUnicode)) > mlngExpItemMaxLength Then
                    MsgBox "扩展项目内容超长(最多" & mlngExpItemMaxLength & "个字符或" & Int(mlngExpItemMaxLength) / 2 & "个汉字)！", vbInformation, gstrSysName
                    Me.stbSpec.Tab = 2
                    
                    .Row = n
                    .Col = .ColIndex("内容")
                    Exit Sub
                End If
            Next
        End With
    End If
    
    '零差价管理模式检查价格
    If chk零差价.Enabled = True And chk零差价.Value = 1 Then
        If Me.stbSpec.Tag = "增加" Then
            If Val(Me.txt当前售价.Text) <> Val(Me.txt成本价格.Text) Then
                MsgBox "启用零差价管理模式时，售价和成本价要一致！", vbInformation, gstrSysName
                Me.stbSpec.Tab = 0
                If Me.txt当前售价.Enabled Then Me.txt当前售价.SetFocus
                Exit Sub
            End If
        ElseIf txt当前售价.Enabled = True And txt成本价格.Enabled = True Then
            If Val(Me.txt当前售价.Text) <> Val(Me.txt成本价格.Text) Then
                MsgBox "启用零差价管理模式时，售价和成本价要一致！", vbInformation, gstrSysName
                Me.stbSpec.Tab = 0
                If Me.txt当前售价.Enabled Then Me.txt当前售价.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Not CheckUnit Then Exit Sub
    If Not CheckRequest Then Exit Sub
    
    If Me.stbSpec.Tag = "修改" Then
        If cbo药价属性.Tag = 0 And Me.cbo药价属性.ItemData(Me.cbo药价属性.ListIndex) = 1 Then
            If MsgBox("药品价格属性由【定价】变为了【时价】，是否继续保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        ElseIf cbo药价属性.Tag = 1 And Me.cbo药价属性.ItemData(Me.cbo药价属性.ListIndex) = 0 Then
            If txt当前售价.Enabled = False Then
                gstrSql = "Select b.上次售价 as 价格, b.药库包装" & vbNewLine & _
                                "From 药品规格 B" & vbNewLine & _
                                "Where b.药品id=[1]"
        
                Set rsPrice = zlDatabase.OpenSQLRecord(gstrSql, "时价转定价", lng药品ID)
                If IsNull(rsPrice!价格) Then
                    gstrSql = "Select a.现价 as 价格, b.药库包装" & vbNewLine & _
                                "From 收费价目 A, 药品规格 B" & vbNewLine & _
                                "Where a.收费细目id = b.药品id And a.收费细目id =[1] And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And" & vbNewLine & _
                                "      变动原因 = 1"
                    
                    Set rsPrice = zlDatabase.OpenSQLRecord(gstrSql, "时价转定价", lng药品ID)
                End If
                If MsgBox("药品价格属性由【时价】变为了【定价】新售价为" & zlStr.FormatEx(rsPrice!价格 * rsPrice!药库包装, mintPriceDigit, , True) & "，是否继续保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("药品价格属性由【时价】变为了【定价】新售价为" & txt当前售价.Text & "，是否继续保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    
    '------------------------------------------
    '数据保存
    gstrSql = "'" & Me.txt编码.Text & "','" & MoveSpecialChar(Me.txt规格.Text) & "','" & MoveSpecialChar(Me.txt产地.Text, False) & "'"
    gstrSql = gstrSql & ",null ,null ,null ,'" & MoveSpecialChar(Me.txt数字码.Text) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Me.txt标识码.Text) & "','" & Mid(Me.cbo药品来源.Text, InStr(1, Me.cbo药品来源.Text, "-") + 1) & "','" & MoveSpecialChar(Me.txt批准文号.Text) & "','" & MoveSpecialChar(Me.txt注册商标.Text) & "'"
    gstrSql = gstrSql & ",'" & Me.txt售价单位.Text & "'," & Val(Me.txt剂量系数.Text)
    gstrSql = gstrSql & ",'" & Me.txt药房单位.Text & "'," & Val(Me.txt药房包装.Text)
    gstrSql = gstrSql & ",'" & Me.txt药库单位.Text & "'," & Val(Me.txt药库包装.Text)
    gstrSql = gstrSql & "," & IIf(cbo申领单位.ListIndex = 0, 1, IIf(cbo申领单位.ListIndex = 1, 3, 4)) '申领单位（1-零售单位;3-药房单位;4-药库单位）
    gstrSql = gstrSql & "," & Val(txt申领阀值.Tag)       '始终以零售单位保存
    gstrSql = gstrSql & "," & Me.cbo药价属性.ItemData(Me.cbo药价属性.ListIndex)
    If Val(Me.lbl批价单位(0).Tag) <> 0 Then
        dbl指导售价 = FormatEx(Val(txt指导售价.Text) / Val(txt药库包装.Text), gtype_MaxDigits.dig_零售价)
        dbl当前售价 = FormatEx(Val(txt当前售价.Text) / Val(txt药库包装.Text), gtype_MaxDigits.dig_零售价)
        dbl成本价格 = FormatEx(Val(txt成本价格.Text) / Val(txt药库包装.Text), gtype_MaxDigits.dig_成本价)
        gstrSql = gstrSql & "," & FormatEx(Val(Me.txt指导批价.Text) / Val(Me.txt药库包装), gtype_MaxDigits.dig_成本价)
    Else
        dbl当前售价 = FormatEx(Val(txt当前售价.Text), gtype_MaxDigits.dig_零售价)
        dbl指导售价 = FormatEx(Val(txt指导售价.Text), gtype_MaxDigits.dig_零售价)
        dbl成本价格 = FormatEx(Val(txt成本价格.Text), gtype_MaxDigits.dig_成本价)
        gstrSql = gstrSql & "," & FormatEx(Val(Me.txt指导批价.Text), gtype_MaxDigits.dig_成本价)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt扣率.Text) & "," & dbl指导售价 & "," & Val(Me.txt加成率.Text) & "," & 0
    gstrSql = gstrSql & ",'" & Mid(Me.cbo药价级别.Text, InStr(1, Me.cbo药价级别.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo费用类型.Text, InStr(1, Me.cbo费用类型.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Me.cbo服务对象.ItemData(Me.cbo服务对象.ListIndex) & "," & Me.chkGMP认证.Value & "," & mint招标药品 & "," & Me.chk屏蔽费别.Value
    gstrSql = gstrSql & "," & Me.cbo住院可否分零.ItemData(Me.cbo住院可否分零.ListIndex)
    gstrSql = gstrSql & "," & Me.chk药库 & "," & Me.chk药房 & ",0"
    gstrSql = gstrSql & ",null "
    
    If Me.stbSpec.Tag = "增加" Then
        lng药品ID = Sys.NextId("收费项目目录")
        gstrSql = "zl_草药规格_Insert(" & lng药名id & "," & lng药品ID & "," & gstrSql
        gstrSql = gstrSql & "," & dbl成本价格 & "," & dbl当前售价 & "," & Me.cbo收入记入.ItemData(Me.cbo收入记入.ListIndex) & ""
    Else
        gstrSql = "zl_草药规格_Update(" & lng药品ID & "," & gstrSql
        gstrSql = gstrSql & "," & dbl成本价格 & "," & dbl当前售价 & "," & Me.cbo收入记入.ItemData(Me.cbo收入记入.ListIndex) & ""
    End If
    
    gstrSql = gstrSql & "," & ZVal(Split(Me.txt合同单位.Tag, "|")(0)) & ",'"
    gstrSql = gstrSql & MoveSpecialChar(Me.txt说明.Text) & "'" & ","
    gstrSql = gstrSql & IIf(Me.chk住院动态分零.Enabled = False, 0, chk住院动态分零.Value) & ",'"
    gstrSql = gstrSql & cbo发药类型.Text & "','"
    gstrSql = gstrSql & MoveSpecialChar(txt备选码.Text) & "',"
    gstrSql = gstrSql & 0 & ","
    gstrSql = gstrSql & "null,"
    gstrSql = gstrSql & IIf(opt中药形态(0).Value, 0, IIf(opt中药形态(1).Value, 1, 2)) & ","
    gstrSql = gstrSql & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点) & ","
    gstrSql = gstrSql & chk非常备药.Value & ",'"
    gstrSql = gstrSql & txt病案费目.Text & "'"
    gstrSql = gstrSql & "," & Me.cbo门诊可否分零.ItemData(Me.cbo门诊可否分零.ListIndex)
    gstrSql = gstrSql & ",'" & Trim(txt送货单位.Text) & "'"
    gstrSql = gstrSql & "," & IIf(Trim(txt送货包装.Text) = "", "Null", Val(Trim(txt送货包装.Text)) * Val(Trim(txt药库包装.Text)))
    gstrSql = gstrSql & "," & chk摆药.Value & ","
    gstrSql = gstrSql & chk零差价.Value & ","
    gstrSql = gstrSql & "'" & MoveSpecialChar(Me.txt本位码.Text) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Me.txt原产地.Text, False) & "'"
    gstrSql = gstrSql & " )"
  
    err = 0: On Error GoTo ErrHand
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    '保存扩展项目
    If stbSpec.TabVisible(2) = True Then
        With vsfItem
            For n = 1 To .Rows - 1
                If Trim(.TextMatrix(n, .ColIndex("内容"))) <> "" Then
                    strItems = IIf(strItems = "", "", strItems & "|") & .TextMatrix(n, .ColIndex("项目")) & "," & Trim(.TextMatrix(n, .ColIndex("内容")))
                End If
            Next
        End With
        
        If strItems <> "" Then
            gstrSql = "Zl_药品规格扩展信息_Update("
            '药品ID
            gstrSql = gstrSql & lng药品ID
            '项目内容串
            gstrSql = gstrSql & "," & "'" & strItems & "'"
            gstrSql = gstrSql & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSql, "保存药品规格扩展信息")
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    '零差价管理
    If chk零差价.Enabled = True And chk零差价.Value = 1 Then
        If CheckPriceAdjust(lng药品ID, 0, -1) = False Then
            MsgBox "该药品已启用零差价管理，但售价和成本价不一致，请注意调价！", vbInformation, gstrSysName
        End If
    End If
    
    If Me.stbSpec.Tag = "增加" Then
        'Val(zldatabase.GetPara("规格增加模式", glngSys, 1023, 0)) = 0
        If ActiveControl Is cmdOK Then '一般保存
            Unload Me
        ElseIf ActiveControl Is cmdSaveAddSpec Then    '保存后增加规格
            Call frmMediLists.zlRefRecords(lng药名id)
            Call Form_Activate
            Me.stbSpec.Tab = 0: Me.txt规格.SetFocus
        ElseIf ActiveControl Is cmdSaveAddItem Then '保存后新增品种
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "增加"
                .lng分类id = mlng分类id
                .lng药名id = 0
                .strPrivs = gstrPrivs
                Unload Me
                .Show 1, frmMediLists
            End With
        End If
    Else
        Unload Me
    End If
    Exit Sub

ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub IniStationNo()
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset

    lblStationNo.Visible = False
    cmbStationNo.Visible = False
    
    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select 编号,名称 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
    End If
End Sub
Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub cmdSaveAddItem_Click()
    mblnOtherSave = True
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    mblnOtherSave = True
    Call cmdOK_Click
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmd病案_Click()
    On Error GoTo ErrHandle
    Dim strSql As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    
    mbln病案费目 = True
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
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo ErrHandle
    
    vRect = zlControl.GetControlRect(txt产地.hwnd)
    
    gstrSql = "Select 编码 as id,名称,简码 From 药品生产商  Order By 编码 "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "cmd产地_Click", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
        
    If blnCancel = True Then txt产地.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    If rsTemp.State = 0 Then
        MsgBox "请初始化药品生产商（字典管理）！", vbInformation, gstrSysName
        Me.txt产地.Tag = "": Me.txt产地.SetFocus: Exit Sub
        Exit Sub
    End If
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    End If

    txt产地.SetFocus
    txt产地.Text = rsTemp!名称
    txt产地.Tag = txt产地.Text
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd合同单位_Click()
    Dim vRect As RECT, blnCancel As Boolean

    vRect = zlControl.GetControlRect(txt合同单位.hwnd)
    
    On Error GoTo ErrHandle
    gstrSql = "Select 编码,名称,简码,id" & _
              " From 供应商" & _
              " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
              " Order By 编码 "
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd合同单位_Click")
'        Call SQLTest
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "cmd合同单位_Click", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
    
    If blnCancel = True Then txt合同单位.SetFocus: Exit Sub  '打开选择器时，点Esc不做以下处理
    
    If rsTemp.State = 0 Then
        MsgBox "请初始化药品供应商（字典管理）！", vbInformation, gstrSysName
        Me.txt合同单位.Tag = "": Me.txt合同单位.SetFocus: Exit Sub
        Exit Sub
    End If
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    End If

    txt合同单位.SetFocus
    Me.txt合同单位.Text = rsTemp!名称
    Me.txt合同单位.Tag = rsTemp!ID & "|" & rsTemp!名称
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd原产地_Click()
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo ErrHandle
    
    vRect = zlControl.GetControlRect(txt原产地.hwnd)
    
    gstrSql = "Select 编码 as id,名称,简码 From 药品生产商  Order By 编码 "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "cmd产地_Click", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
        
    If blnCancel = True Then txt原产地.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    If rsTemp.State = 0 Then
        MsgBox "请初始化药品生产商（字典管理）！", vbInformation, gstrSysName
        Me.txt原产地.Tag = "": Me.txt原产地.SetFocus: Exit Sub
        Exit Sub
    End If
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    End If

    txt原产地.SetFocus
    txt原产地.Text = rsTemp!名称
    txt原产地.Tag = txt原产地.Text
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Dim blnExit As Boolean
    Dim strMsg As String
    Dim i As Integer
    Dim rs差价率 As ADODB.Recordset
    Dim str送货单位 As String
    Dim dbl送货包装 As Double
    Dim rsTmp As ADODB.Recordset
    
    If mbln病案费目 = True Then Exit Sub
    mintSet分批 = Val(zlDatabase.GetPara("药品分批属性自动设置", glngSys, 1023, 0))
    '----------依赖关系判断-------------------------------------
    If Me.cbo药品来源.ListCount = 0 Then
        strMsg = "未设置药品来源分类（字典管理）！"
        blnExit = True
    End If
    If Me.cbo费用类型.ListCount = 0 And Not blnExit Then
        strMsg = "未设置用于药品的医保类型（字典管理）！"
        blnExit = True
    End If
    If Me.cbo收入记入.ListCount = 0 And Not blnExit Then
        strMsg = "未设置明细的收入项目！"
        blnExit = True
    End If
    If Me.cbo药价级别.ListCount = 0 And Not blnExit Then
        strMsg = "未设置药价管理级别（字典管理）！"
        blnExit = True
    End If
    If blnExit Then
        MsgBox strMsg, vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    txt本位码.MaxLength = Val(zlDatabase.GetPara("本位码", glngSys, 1023, 20))
    txt数字码.MaxLength = Val(zlDatabase.GetPara("数字码", glngSys, 1023, 7))
    If Me.stbSpec.Tag <> "增加" Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
'    If mblnLoad = True Then Exit Sub
    '----------药品品种识别-------------------------------------
    On Error GoTo ErrHandle
    gstrSql = "select I.类别,I.编码,I.名称,I.计算单位,T.药品剂型" & _
            " from 诊疗项目目录 I,药品特性 T" & _
            " where I.ID=T.药名ID and I.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
    
    With rsTemp
'        If !类别 = "5" Then
            Me.Tag = "7": Me.Caption = "中草药规格编辑"
            Me.lbl收入记入.Tag = zlDatabase.GetPara("中草药收入项目", glngSys, 1023, False)
'        Else
'            Me.Tag = "6": Me.Caption = "中成药规格编辑"
'            Me.lbl收入记入.Tag = zlDatabase.GetPara("中成药收入项目", glngSys, 1023, False)
'        End If
        If Me.stbSpec.Tag = "增加" And Val(Me.lbl收入记入.Tag) = 0 Then
            MsgBox "没有设置“中草药”对应的收入项目（本地参数设置）！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        For intCount = 0 To Me.cbo收入记入.ListCount - 1
            If Me.cbo收入记入.ItemData(intCount) = Val(Me.lbl收入记入.Tag) Then
                Me.cbo收入记入.ListIndex = intCount: Exit For
            End If
        Next
        
        Me.lbl品种.Caption = "药品编码：" & !编码 & _
                "   中药名称：" & !名称 & _
                "   剂型：" & IIf(IsNull(!药品剂型), "", !药品剂型) & _
                "   剂量单位：" & IIf(IsNull(!计算单位), "", !计算单位)
        Me.lbl品种.Tag = !编码
        Me.lbl售价单位Child.Caption = IIf(IsNull(!计算单位), "", !计算单位)
    
        Me.lbl批价单位(0).Tag = Val(zlDatabase.GetPara(29, glngSys))
        
        mintCostDigit = GetDigit(1, 1, IIf(Me.lbl批价单位(0).Tag = 0, 1, 4))
        mintPriceDigit = GetDigit(1, 2, IIf(Me.lbl批价单位(0).Tag = 0, 1, 4))
        
        mintSaleCostDigit = GetDigit(1, 1, 1)
        mintSalePriceDigit = GetDigit(1, 2, 1)

    End With
    
    '----------数据装载-------------------------------------
    '只要存在lng药品ID，则无论什么状态都读该规格信息
    gstrSql = "select I.编码,S.本位码,I.规格,I.产地,S.标识码,S.药品来源,S.批准文号,S.注册商标,S.中药形态," & _
            "        I.计算单位,S.剂量系数,S.门诊单位,S.门诊包装,S.住院单位,S.住院包装,S.药库单位,S.药库包装,s.送货单位,s.送货包装," & _
            "        I.是否变价,S.指导批发价,S.扣率,S.指导零售价,i.病案费目,S.加成率,S.管理费比例,S.成本价,S.招标药品,S.GMP认证,S.基本药物, " & _
            "        S.药价级别,I.费用类型,I.服务对象,I.屏蔽费别,S.申领单位,S.申领阀值," & _
            "        S.住院可否分零,S.动态分零 as 住院动态分零,S.门诊可否分零,S.药库分批,S.药房分批,S.最大效期,S.发药类型,I.备选码," & _
            "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,S.合同单位id,G.名称 合同单位,I.说明,I.站点,S.增值税率,S.是否常备,s.是否摆药,s.是否零差价管理,s.原产地 " & _
            " from 收费项目目录 I,药品规格 S,(Select Id,名称 From 供应商 Where 末级 = 1 And substr(类型,1,1) = '1' And " & _
            " 撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) G " & _
            " where I.ID=S.药品ID and G.id(+)=S.合同单位id and I.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt编码.Text = !编码
            Me.txt本位码.Text = Nvl(!本位码)
            Me.txt规格.Text = IIf(IsNull(!规格), "", !规格)
            Me.txt产地.Text = IIf(IsNull(!产地), "", !产地)
            Me.txt原产地.Text = IIf(IsNull(!原产地), "", !原产地)
            Me.txt合同单位.Text = IIf(IsNull(!合同单位), "", !合同单位)
            Me.txt合同单位.Tag = IIf(IsNull(!合同单位id), "|", !合同单位id & "|" & !合同单位)
            Me.txt标识码.Text = IIf(IsNull(!标识码), "", !标识码)
            Me.txt说明.Text = IIf(IsNull(!说明), "", !说明)
            Me.txt备选码.Text = IIf(IsNull(!备选码), "", !备选码)

            For intCount = 0 To Me.cbo药品来源.ListCount - 1
                If Mid(Me.cbo药品来源.List(intCount), InStr(1, Me.cbo药品来源.List(intCount), "-") + 1) = IIf(IsNull(!药品来源), "", !药品来源) Then
                    Me.cbo药品来源.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt批准文号.Text = IIf(IsNull(!批准文号), "", !批准文号)
            Me.txt注册商标.Text = IIf(IsNull(!注册商标), "", !注册商标)
            Me.txt售价单位.Text = IIf(IsNull(!计算单位), "", !计算单位)
            Me.lbl药房单位Child.Caption = Me.txt售价单位 & ")"
'            Me.lbl住院单位Child.Caption = Me.txt售价单位 & ")"
            Me.lbl药库单位Child.Caption = Me.txt售价单位 & ")"
            Me.lbl剂量系数.Caption = "(1" & Me.txt售价单位 & "="
            Me.txt剂量系数.Text = FormatEx(IIf(IsNull(!剂量系数), 1, !剂量系数), 5, , False)
            Me.txt药房单位.Text = IIf(IsNull(!门诊单位), "", !门诊单位)
            Me.lbl药房包装.Caption = "(1" & Me.txt药房单位.Text & "="
            Me.txt药房包装.Text = FormatEx(IIf(IsNull(!门诊包装), 1, !门诊包装), 5, , False)
'            Me.txt住院单位.Text = IIf(IsNull(!住院单位), "", !住院单位)
'            Me.lbl住院包装.Caption = "(1" & Me.txt住院单位.Text & "="
'            Me.txt住院包装.Text = IIf(IsNull(!住院包装), 1, !住院包装)
            Me.txt药库单位.Text = IIf(IsNull(!药库单位), "", !药库单位)
            Me.lbl药库包装.Caption = "(1" & Me.txt药库单位.Text & "="
            Me.txt药库包装.Text = FormatEx(IIf(IsNull(!药库包装), 1, !药库包装), 5, , False)
            str送货单位 = IIf(IsNull(!送货单位), "", !送货单位)
            dbl送货包装 = IIf(IsNull(!送货单位), 0, !送货包装)
            Me.txt送货单位.Text = str送货单位
            Me.txt送货包装.Text = IIf(dbl送货包装 = 0, "", FormatEx(dbl送货包装 / !药库包装, 1, , True))
            lbl送货单位child.Caption = txt药库单位.Text
            
            Me.cbo申领单位.ListIndex = IIf(Nvl(!申领单位, 1) = 1, 0, IIf(Nvl(!申领单位, 1) = 4, 2, 1))
            For i = 0 To cbo发药类型.ListCount
                If cbo发药类型.List(i) = !发药类型 Then
                    Me.cbo发药类型.ListIndex = i
                    Exit For
                ElseIf IsNull(!发药类型) Then
                    Me.cbo发药类型.ListIndex = 0
                End If
            Next
            SetStationNo IIf(IsNull(!站点), "", !站点)
            
            Select Case Nvl(!申领单位, 1)
            Case 1 '零售
                Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0), "#0.00;-#0.00; ;")
            Case 3 '药房
                Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0) / Nvl(!门诊包装, 1), "#0.00;-#0.00; ;")
            Case Else '药库
                Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0) / Nvl(!药库包装, 1), "#0.00;-#0.00; ;")
            End Select
            
            Me.cbo药价属性.ListIndex = IIf(IsNull(!是否变价), 0, !是否变价)
            cbo药价属性.Tag = Me.cbo药价属性.ListIndex
            Me.txt扣率.Text = IIf(IsNull(!扣率), 100, !扣率)
            If Me.stbSpec.Tag = "增加" Then
                Me.txt指导批价.Text = ""
                Me.txt指导售价.Text = ""
                Me.txt成本价格.Text = ""
                txt当前售价.Text = ""
            Else
                If Val(Me.lbl批价单位(0).Tag) <> 0 Then
                    Me.txt指导批价.Text = FormatEx(IIf(IsNull(!指导批发价), 0, !指导批发价) * Me.txt药库包装.Text, mintCostDigit, , True)
                    Me.txt指导售价.Text = FormatEx(IIf(IsNull(!指导零售价), 0, !指导零售价) * Me.txt药库包装.Text, mintPriceDigit, , True)
                    Me.txt成本价格.Text = FormatEx(IIf(IsNull(!成本价), 0, !成本价) * Me.txt药库包装.Text, mintCostDigit, , True)
                Else
                    Me.txt指导批价.Text = FormatEx(IIf(IsNull(!指导批发价), 0, !指导批发价), mintCostDigit, , True)
                    Me.txt指导售价.Text = FormatEx(IIf(IsNull(!指导零售价), 0, !指导零售价), mintPriceDigit, , True)
                    Me.txt成本价格.Text = FormatEx(IIf(IsNull(!成本价), 0, !成本价), mintCostDigit, , True)
                End If
            End If
            Me.txt结算价 = FormatEx(Val(Me.txt指导批价.Text) * Me.txt扣率.Text / 100, mintPriceDigit, , True)
            Me.txt加成率.Text = Format(IIf(IsNull(!加成率), 0, !加成率), "0.00")
            
            For intCount = 0 To Me.cbo药价级别.ListCount - 1
                If Mid(Me.cbo药价级别.List(intCount), InStr(1, Me.cbo药价级别.List(intCount), "-") + 1) = IIf(IsNull(!药价级别), "", !药价级别) Then
                    Me.cbo药价级别.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo费用类型.ListCount - 1
                If Mid(Me.cbo费用类型.List(intCount), InStr(1, Me.cbo费用类型.List(intCount), "-") + 1) = IIf(IsNull(!费用类型), "", !费用类型) Then
                    Me.cbo费用类型.ListIndex = intCount: Exit For
                End If
            Next
            Me.cbo服务对象.ListIndex = IIf(IsNull(!服务对象), 0, !服务对象)
            Me.chk屏蔽费别.Value = IIf(IsNull(!屏蔽费别), 0, !屏蔽费别)
            Me.chk住院动态分零.Value = IIf(IsNull(!住院动态分零), 0, !住院动态分零)
            Me.chk非常备药.Value = IIf(IsNull(!是否常备), 0, !是否常备)
            Me.chk摆药.Value = IIf(IsNull(!是否摆药), 0, !是否摆药)
            Me.chk零差价.Value = IIf(IsNull(!是否零差价管理), 0, !是否零差价管理)
            
'            Select Case IIf(IsNull(!住院可否分零), 0, !住院可否分零)
'            Case 0, 1
'                Me.cbo住院可否分零.ListIndex = IIf(IsNull(!住院可否分零), 0, !住院可否分零)
'            Case Else
'                Me.cbo住院可否分零.ListIndex = 0
'            End Select
            If IIf(IsNull(!住院可否分零), 0, !住院可否分零) = 0 Then
                Me.cbo住院可否分零.ListIndex = 0
            Else
                Me.cbo住院可否分零.ListIndex = 1
            End If
            
            If IIf(IsNull(!门诊可否分零), 0, !门诊可否分零) = 0 Then
                Me.cbo门诊可否分零.ListIndex = 0
            Else
                Me.cbo门诊可否分零.ListIndex = 1
            End If
            
            Me.chkGMP认证.Value = IIf(IsNull(!GMP认证), 0, !GMP认证)
            
            If IsNull(!中药形态) Then
                Me.opt中药形态(0).Tag = 0
                Me.opt中药形态(0).Value = True
            ElseIf !中药形态 = 1 Then
                Me.opt中药形态(0).Tag = 1
                Me.opt中药形态(1).Value = True
            ElseIf !中药形态 = 2 Then
                Me.opt中药形态(0).Tag = 2
                Me.opt中药形态(2).Value = True
            Else
                Me.opt中药形态(0).Tag = 0
                Me.opt中药形态(0).Value = True
            End If
            Call opt中药形态_Click(opt中药形态(0).Tag)
            
            If Me.stbSpec.Tag <> "增加" Then mint招标药品 = IIf(IsNull(!招标药品), 0, !招标药品)
            If mint招标药品 = 1 Then Me.lbl指导批价.Caption = "中标价格(&B)"
            
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "注：该规格于" & Format(!建档时间, "YYYY年MM月DD日") & "建立"
            Else
                Me.lblFound.Caption = "注：该规格于" & Format(!建档时间, "YYYY年MM月DD日") & "建立，" & Format(!撤档时间, "YYYY年MM月DD日") & "停用"
            End If
            txt病案费目.Text = IIf(IsNull(!病案费目), "", !病案费目)
            
            Me.chk药房.Tag = IIf(IsNull(!药房分批), 0, !药房分批)
'            Me.txt效期.Tag = IIf(IsNull(!最大效期), 0, !最大效期)
            
            Me.chk药库.Value = IIf(IsNull(!药库分批), 0, Abs(!药库分批))
            If Me.chk药库.Value = 0 Then
                Me.chk药房.Enabled = False: Me.chk药房.Value = 0
'                Me.chk效期.Enabled = False: Me.chk效期.Value = 0
'                Me.txt效期.Enabled = False: Me.chk效期.Value = 0
            Else
                Me.chk药房.Enabled = True
'                Me.chk效期.Enabled = True
                Me.chk药房.Value = Me.chk药房.Tag
'                Me.txt效期.Text = Me.txt效期.Tag
'                If Val(Me.txt效期.Text) = 0 Then
'                    Me.txt效期.Enabled = False: Me.chk效期.Value = 0
'                Else
'                    Me.txt效期.Enabled = True: Me.chk效期.Value = 1
'                End If
            End If
        End If
        If Trim(Me.txt合同单位.Tag) = "" Then
            Me.txt合同单位.Tag = "|"
        End If
        If Val(Me.lbl批价单位(0).Tag) <> 0 Then
            Me.lbl批价单位(0).Caption = "元/" & Me.txt药库单位
            Me.lbl批价单位(1).Caption = "元/" & Me.txt药库单位
        Else
            Me.lbl批价单位(0).Caption = "元/" & Me.txt售价单位
            Me.lbl批价单位(1).Caption = "元/" & Me.txt售价单位
        End If
    End With
        
    If Me.stbSpec.Tag = "增加" Then
        gstrSql = "Select 加成率" & vbNewLine & _
                "From 药品规格" & vbNewLine & _
                "Where 药品id = (Select Max(药品id) From 药品规格 A, 收费项目目录 B Where a.药品id = b.Id And b.类别 = [1])"

        Set rs差价率 = zlDatabase.OpenSQLRecord(gstrSql, "加成率查询", Me.Tag)
        
        If rs差价率.RecordCount > 0 Then
            Me.txt加成率.Text = Format(IIf(IsNull(rs差价率!加成率), 0, rs差价率!加成率), "0.00")
        End If
        
        '增加时，重新提取编码号，清空规格和厂牌
        Me.txt编码.Text = "": Me.txt规格.Text = "": Me.txt产地.Text = "": Me.txt原产地.Text = "": Me.lblFound.Caption = ""
        gstrSql = "select max(I.编码) as 最大编码 from 收费项目目录 I,药品规格 S where I.ID=S.药品ID and  S.药名ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
        With rsTemp
            If .BOF Or .EOF Then
                Me.txt编码.Text = Me.lbl品种.Tag & "01"
            ElseIf IsNull(!最大编码) Then
                Me.txt编码.Text = Me.lbl品种.Tag & "01"
            ElseIf Len(!最大编码) <= Len(Me.lbl品种.Tag) Then
                Me.txt编码.Text = Me.lbl品种.Tag & "01"
            Else
                Me.txt编码.Text = zlStr.Increase(!最大编码)
            End If
        End With
        
        If mintSet分批 = 0 Then
            gstrSql = "Select b.药库分批, b.药房分批" & _
                       " From 药品规格 B, (Select Max(a.Id) As ID From 收费项目目录 A, 药品规格 B Where a.Id = b.药品id And b.药名id = [1]) C" & _
                       " Where b.药品id = c.Id"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
            
            If rsTmp.RecordCount <> 0 Then
                chk药库.Value = IIf(IsNull(rsTmp!药库分批), "0", rsTmp!药库分批)
                chk药房.Value = IIf(IsNull(rsTmp!药房分批), "0", rsTmp!药房分批)
            End If
        ElseIf mintSet分批 = 1 Then
            chk药库.Value = 1
            chk药房.Value = 0
            chk药库.Enabled = False
            chk药房.Enabled = False
        ElseIf mintSet分批 = 2 Then
            chk药库.Value = 1
            chk药房.Value = 1
            chk药库.Enabled = False
            chk药房.Enabled = False
        ElseIf mintSet分批 = 3 Then
            chk药库.Value = 0
            chk药房.Value = 0
            chk药库.Enabled = False
            chk药房.Enabled = False
        End If
    Else
        '提取商品名和简码、数字码
        gstrSql = "select 名称,性质,简码,码类 from 收费项目别名 where 收费细目id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        With rsTemp
            Do While Not .EOF
                If !性质 = 1 And !码类 = 3 Then
                    Me.txt数字码.Text = IIf(IsNull(!简码), "", !简码)
                End If
'                If !性质 = 3 And !码类 = 1 Then
'                    Me.txt商品名.Text = IIf(IsNull(!名称), "", !名称)
'                    Me.txt拼音.Text = IIf(IsNull(!简码), "", !简码)
'                End If
'                If !性质 = 3 And !码类 = 2 Then
'                    Me.txt商品名.Text = IIf(IsNull(!名称), "", !名称)
'                    Me.txt五笔.Text = IIf(IsNull(!简码), "", !简码)
'                End If
                .MoveNext
            Loop
        End With
        
        '提取显示当前售价
        If Me.cbo药价属性.ListIndex <> 0 Then
            '时价药品，取库存金额/库存数量做为其价格，无库存时取价表定价
            gstrSql = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                    " from 收费价目 P," & _
                    "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                    "      From 药品库存 Where 药品ID=[1]) K" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)" & _
                     GetPriceClassString("P")
        Else
            '非时价药品调价，取其价格记录中的价格
            gstrSql = "select P.现价,P.收入项目id" & _
                    " from 收费价目 P" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)" & _
                     GetPriceClassString("P")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        With rsTemp
            If .RecordCount > 0 Then
                If Val(Me.lbl批价单位(0).Tag) <> 0 Then
                    Me.txt当前售价.Text = FormatEx(!现价 * Val(txt药库包装.Text), mintPriceDigit, , True)
                Else
                    Me.txt当前售价.Text = FormatEx(!现价, mintPriceDigit, , True)
                End If
                For intCount = 0 To Me.cbo收入记入.ListCount - 1
                    If Me.cbo收入记入.ItemData(intCount) = !收入项目id Then
                        Me.cbo收入记入.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        End With
        
        '根据是否有发生，确定：药价属性、成本价格、零售价格可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品收发记录 Where 药品ID=[1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        mblnUsed = False
        With rsTemp
            If .Fields(0).Value > 0 Then
                mblnUsed = True
                Me.txt成本价格.Enabled = False
                Me.txt当前售价.Enabled = False
'                Me.cbo收入记入.Enabled = False
                Me.txt剂量系数.Enabled = False
'                Me.txt住院包装.Enabled = False
                Me.txt药房包装.Enabled = False
                Me.txt药库包装.Enabled = False
            Else
                Me.cbo药价属性.Enabled = True
                Me.txt当前售价.Enabled = True
                Me.txt成本价格.Enabled = True
'                Me.cbo收入记入.Enabled = True
                Me.txt剂量系数.Enabled = True
'                Me.txt住院包装.Enabled = True
                Me.txt药房包装.Enabled = True
                Me.txt药库包装.Enabled = True
            End If
        End With
        
        '根据是否存在医嘱记录，确定剂量系数是否能够修改
        gstrSql = "Select 1 From 病人医嘱记录 Where 收费细目ID=[1] And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        If rsTemp.RecordCount > 0 Then
            Me.txt剂量系数.Enabled = False
        Else
            Me.txt剂量系数.Enabled = True
        End If
        
        '根据是否有库存，确定：分批特性可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                 " Where A.药品ID=[1] And A.库房ID=B.部门ID And B.工作性质 Like '%药库'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        If rsTemp.Fields(0).Value > 0 Then
            Me.chk药库.Enabled = False
'            Me.chk效期.Enabled = False
        Else
            Me.chk药库.Enabled = True
        End If
        If Me.chk药库.Value = 1 Then
            gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                     " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
            
            If rsTemp.Fields(0).Value > 0 Then
                Me.chk药房.Enabled = False
                If Me.chk药库.Enabled Then Me.chk药库.Enabled = IIf(chk药房.Value = 1, False, True)
            Else
                Me.chk药房.Enabled = True
            End If
        End If
    End If
    
    
    '----------程序权限控制-------------------------------------
    If Me.stbSpec.Tag = "增加" Or Me.stbSpec.Tag = "修改" Then
        If InStr(1, strPrivs, "医保用药目录") = 0 Then
            Me.cbo费用类型.Enabled = False: Me.txt标识码.Enabled = False:
        End If
        If InStr(1, strPrivs, "管理扣率") = 0 Then Me.txt扣率.Enabled = False
        If InStr(1, strPrivs, "指导价格管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.txt指导批价.Text = "0"
                Me.txt指导售价.Text = "0"
            End If
            Me.txt加成率.Enabled = False
            Me.txt指导批价.Enabled = False: Me.txt指导售价.Enabled = False
        End If
        If InStr(1, strPrivs, "售价管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.txt当前售价.Text = "0"
                Me.cbo药价属性.ListIndex = 0
            End If
            Me.cbo药价属性.Enabled = False
'            Me.txt差价让利.Enabled = False:
            Me.txt当前售价.Enabled = False
        End If
        If InStr(1, strPrivs, "调整收入项目") = 0 Then
            cbo收入记入.Enabled = False
        End If
        If InStr(1, strPrivs, "药价级别") = 0 Then
             Me.cbo药价级别.Enabled = False
        End If
        If InStr(1, strPrivs, "成本价管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.txt成本价格.Text = "0"
            End If
            Me.txt成本价格.Enabled = False
        End If
        If InStr(1, strPrivs, "调整服务对象") = 0 Then
            Me.cbo服务对象.Enabled = False
        End If
        If InStr(1, strPrivs, "草药分包管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.opt中药形态(0).Value = True
            End If
            Me.opt中药形态(0).Enabled = False: Me.opt中药形态(1).Enabled = False: Me.opt中药形态(2).Enabled = False
        End If
    Else
        Me.txt编码.Enabled = False: Me.txt本位码.Enabled = False: Me.txt规格.Enabled = False: Me.txt产地.Enabled = False: cmd产地.Enabled = False
        Me.txt原产地.Enabled = False: cmd原产地.Enabled = False
'        Me.txt商品名.Enabled = False: Me.txt拼音.Enabled = False: Me.txt五笔.Enabled = False:
        Me.txt数字码.Enabled = False
        Me.txt标识码.Enabled = False: Me.cbo药品来源.Enabled = False: Me.txt批准文号.Enabled = False: Me.txt注册商标.Enabled = False
        Me.txt售价单位.Enabled = False: Me.txt剂量系数.Enabled = False: Me.txt药房单位.Enabled = False: Me.txt药房包装.Enabled = False
        'Me.txt住院单位.Enabled = False: Me.txt住院包装.Enabled = False
        Me.txt药库单位.Enabled = False: Me.txt药库包装.Enabled = False
        Me.cbo申领单位.Enabled = False: Me.txt申领阀值.Enabled = False: Me.cbo发药类型.Enabled = False
        
        Me.cbo药价属性.Enabled = False: Me.txt指导批价.Enabled = False: Me.txt扣率.Enabled = False: Me.txt结算价.Enabled = False
        Me.txt指导售价.Enabled = False: Me.txt加成率.Enabled = False
        Me.cbo药价级别.Enabled = False: Me.cbo费用类型.Enabled = False: Me.cbo服务对象.Enabled = False: Me.chk屏蔽费别.Enabled = False
        Me.txt成本价格.Enabled = False: Me.txt当前售价.Enabled = False: Me.cbo收入记入.Enabled = False
        Me.cbo住院可否分零.Enabled = False: Me.chk药库.Enabled = False: Me.chk药房.Enabled = False ': Me.chk效期.Enabled = False: Me.txt效期.Enabled = False
        cbo门诊可否分零.Enabled = False
        Me.chk非常备药.Enabled = False
        Me.txt合同单位.Enabled = False: Me.cmd合同单位.Enabled = False
        Me.txt说明.Enabled = False
'        Me.txt差价让利.Enabled = False: Me.chkGMP认证.Enabled = False: Me.chk基本药物.Enabled = False
        Me.txt备选码.Enabled = False
        Me.opt中药形态(0).Enabled = False: Me.opt中药形态(1).Enabled = False: Me.opt中药形态(2).Enabled = False
        Me.cmbStationNo.Enabled = False
        cmdOK.Visible = False: cmdCancel.Caption = "关闭(&C)"
        txt病案费目.Enabled = False
        cmd病案.Enabled = False
        chk摆药.Enabled = False
        chk零差价.Enabled = False
        chkGMP认证.Enabled = False
        vsfItem.Enabled = False
    End If
    
    '如果本次操作是修改，则检查是否存在“药品单位管理”的权限，没有则不允许修改药品单位与系数
    If Me.stbSpec.Tag = "修改" Then
        If InStr(1, strPrivs, "药品单位管理") = 0 Then
            txt售价单位.Enabled = False
'            txt住院单位.Enabled = False
            txt药房单位.Enabled = False
            txt药库单位.Enabled = False
            txt剂量系数.Enabled = False
'            txt住院包装.Enabled = False
            txt药房包装.Enabled = False
            txt药库包装.Enabled = False
        End If
    End If
    
'    mblnLoad = True
    Me.stbSpec.Tab = 0
    mstr所有记录 = ""
    mstr所有记录 = txt编码.Text & "|" & txt本位码.Text & "|" & txt规格.Text & "|" & txt产地.Text & "|" & txt原产地.Text & "|" & _
                    txt数字码.Text & "|" & txt标识码.Text & "|" & cbo药品来源.Text & "|" & txt合同单位.Text & "|" & txt说明.Text & "|" & cbo发药类型.Text & "|" & _
                    cmbStationNo.Text & "|" & txt批准文号.Text & "|" & txt注册商标.Text & "|" & txt售价单位.Text & "|" & txt剂量系数.Text & "|" & _
                    txt药房单位.Text & "|" & txt药房包装.Text & "|" & txt药库单位.Text & "|" & txt药库包装.Text & "|" & cbo申领单位.Text & "|" & txt申领阀值.Text & "|" & _
                    txt备选码.Text & "|" & cbo药价属性.Text & "|" & txt成本价格.Text & "|" & txt当前售价.Text & "|" & txt指导批价.Text & "|" & txt扣率.Text & "|" & txt结算价.Text & "|" & _
                    txt指导售价.Text & "|" & txt加成率.Text & "|" & cbo收入记入.Text & "|" & txt病案费目.Text & "|" & cbo药价级别.Text & "|" & _
                    chk屏蔽费别.Value & "|" & cbo费用类型.Text & "|" & cbo服务对象.Text & "|" & cbo住院可否分零.Text & "|" & chk住院动态分零.Value & "|" & _
                    chkGMP认证.Value & "|" & chk非常备药.Value & "|" & chk药库.Value & "|" & chk药房.Value & "|" & cbo门诊可否分零.Text
    
    If txt规格.Enabled = True Then
        txt规格.SetFocus
    End If
    
    '扩展属性
    gstrSql = "Select 1 From 药品规格扩展项目 Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "药品规格扩展项目")
    If rsTmp.RecordCount = 0 Then
        '如果没有扩展项目就不显示扩展页面
        stbSpec.TabVisible(2) = False
    Else
        gstrSql = "Select b.名称, a.内容 From 药品规格扩展信息 A, 药品规格扩展项目 B " & _
            " Where a.项目(+) = b.名称 And a.药品id(+) = [1] Order By b.编码 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "药品规格扩展信息", lng药品ID)
        
        mlngExpItemMaxLength = rsTmp.Fields("内容").DefinedSize
        
        With vsfItem
            .Rows = 1
            
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("项目")) = rsTmp!名称
                .TextMatrix(.Rows - 1, .ColIndex("内容")) = Nvl(rsTmp!内容)
                .TextMatrix(.Rows - 1, .ColIndex("原内容")) = Nvl(rsTmp!内容)
                
                rsTmp.MoveNext
            Loop
        End With
    End If
    
    '零差价模式控制
    If Val(zlDatabase.GetPara(275, glngSys, , 0)) = 0 Then
        chk零差价.Enabled = False
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt病案费目_GotFocus()
    txt病案费目.SelStart = 0
    txt病案费目.SelLength = Len(txt病案费目)
    If Me.stbSpec.Tag = "增加" Or Me.stbSpec.Tag = "修改" Then txt病案费目.SetFocus
End Sub

Private Sub txt病案费目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyDelete Then
        txt病案费目.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
    
        If msf合同单位.Visible Then
            msf合同单位.Visible = False: txt合同单位.SetFocus: Exit Sub
        End If
        
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mint招标药品 = 0
            
    Call GetMaxDigit    '获取金额最大精度
    '如果是药店管理系统，则不显示门诊及住院相关单位及系数，但内容与售价单位与系数一致
    If glngSys \ 100 = 8 Then
        Me.lbl药房单位.Visible = False: Me.txt药房单位.Visible = False: Me.lbl药房包装.Visible = False: Me.txt药房包装.Visible = False: Me.lbl药房单位Child.Visible = False
'        Me.lbl住院单位.Visible = False: Me.txt住院单位.Visible = False: Me.lbl住院包装.Visible = False: Me.txt住院包装.Visible = False: Me.lbl药库单位Child.Visible = False
        Me.lbl药库包装.Top = Me.lbl药房包装.Top: Me.txt药库单位.Top = Me.txt药房单位.Top: Me.lbl药库单位.Top = Me.lbl药房单位.Top: Me.txt药库包装.Top = Me.txt药房包装.Top
        Me.lbl药库单位.Caption = "采购单位(&W)"
    End If
    
    Call GetDefineSize
    Call IniStationNo
    On Error GoTo ErrHandle
    mint分段加成 = Val(zlDatabase.GetPara("售价按加成计算", glngSys, glngModul, 0))
    
    Set mrs分段加成 = Nothing
    If mint分段加成 = 1 Then
        gstrSql = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zlDatabase.OpenSQLRecord(gstrSql, "药品加成方案")
    End If
    
    '----------------装入可选的基础数据----------------------
    With Me.cbo药价属性
        .Clear
        aryTemp = Split("0-定价;1-时价", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(intCount): .ItemData(.NewIndex) = intCount
        Next
        .ListIndex = 0
    End With
    
    gstrSql = "Select 编码||'-'||名称 名称 From 药价管理级别 where 编码=1 Order By 编码"
    With rsTemp
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo药价级别.Clear
        Do While Not rsTemp.EOF
            Me.cbo药价级别.AddItem rsTemp!名称
            rsTemp.MoveNext
        Loop
    End With
    
    With Me.cbo住院可否分零
        .Clear
        .AddItem "0-可以分零": .ItemData(.NewIndex) = 0
        .AddItem "1-不可分零": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    With Me.cbo门诊可否分零
        .Clear
        .AddItem "0-可以分零": .ItemData(.NewIndex) = 0
        .AddItem "1-不可分零": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    With Me.cbo服务对象
        If glngSys \ 100 <> 8 Then
            aryTemp = Split("0-不应用于病人;1-门诊;2-住院;3-门诊和住院", ";")
            For intCount = LBound(aryTemp) To UBound(aryTemp)
                .AddItem aryTemp(intCount): .ItemData(.NewIndex) = intCount
            Next
            .ListIndex = 3
        Else
            .AddItem "0-不外卖": .ItemData(.NewIndex) = 0
            .AddItem "1-外售": .ItemData(.NewIndex) = 3
            .ListIndex = 0
        End If
    End With
    
    With cbo申领单位
        .Clear
        .AddItem "售价单位"
        .AddItem "药房单位"
        .AddItem "药库单位"
        .ListIndex = 0
    End With
    
    With rsTemp
        gstrSql = "Select 编码||'-'||名称 From 药品来源分类 Order By 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo药品来源.Clear
        Do While Not rsTemp.EOF
            Me.cbo药品来源.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo药品来源.ListCount > 0 Then Me.cbo药品来源.ListIndex = 0
        
        gstrSql = "Select 名称 From 发药类型 Order By 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo发药类型.Clear
        Me.cbo发药类型.AddItem ""
        Do While Not rsTemp.EOF
            Me.cbo发药类型.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
    
        gstrSql = "Select 编码||'-'||名称 From 费用类型 where 性质=1 Order By 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo费用类型.Clear
        Me.cbo费用类型.AddItem ""
        Do While Not rsTemp.EOF
            Me.cbo费用类型.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo费用类型.ListCount > 0 Then Me.cbo费用类型.ListIndex = 0
        
        gstrSql = "Select ID,'['||编码||']'||名称 as 名称" & _
                " From 收入项目" & _
                " where 末级=1 and (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo收入记入.Clear
        Do While Not rsTemp.EOF
            Me.cbo收入记入.AddItem rsTemp!名称: Me.cbo收入记入.ItemData(Me.cbo收入记入.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo收入记入.ListCount > 0 Then Me.cbo收入记入.ListIndex = 0
    
'        Me.lbl收入记入.Tag = zlDatabase.GetPara("中草药收入项目", glngSys, 1023, False)
'        For intCount = 0 To Me.cbo收入记入.ListCount - 1
'            If Me.cbo收入记入.ItemData(intCount) = Val(Me.lbl收入记入.Tag) Then
'                Me.cbo收入记入.ListIndex = intCount: Exit For
'            End If
'        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    
    If mblnOtherSave = False Then
        If mblnOK = False And mblnCancel = False Then
            strTemp = txt编码.Text & "|" & txt本位码.Text & "|" & txt规格.Text & "|" & txt产地.Text & "|" & txt原产地.Text & "|" & _
                            txt数字码.Text & "|" & txt标识码.Text & "|" & cbo药品来源.Text & "|" & txt合同单位.Text & "|" & txt说明.Text & "|" & cbo发药类型.Text & "|" & _
                            cmbStationNo.Text & "|" & txt批准文号.Text & "|" & txt注册商标.Text & "|" & txt售价单位.Text & "|" & txt剂量系数.Text & "|" & _
                            txt药房单位.Text & "|" & txt药房包装.Text & "|" & txt药库单位.Text & "|" & txt药库包装.Text & "|" & cbo申领单位.Text & "|" & txt申领阀值.Text & "|" & _
                            txt备选码.Text & "|" & cbo药价属性.Text & "|" & txt成本价格.Text & "|" & txt当前售价.Text & "|" & txt指导批价.Text & "|" & txt扣率.Text & "|" & txt结算价.Text & "|" & _
                            txt指导售价.Text & "|" & txt加成率.Text & "|" & cbo收入记入.Text & "|" & txt病案费目.Text & "|" & cbo药价级别.Text & "|" & _
                            chk屏蔽费别.Value & "|" & cbo费用类型.Text & "|" & cbo服务对象.Text & "|" & cbo住院可否分零.Text & "|" & chk住院动态分零.Value & "|" & _
                            chkGMP认证.Value & "|" & chk非常备药.Value & "|" & chk药库.Value & "|" & chk药房.Value & "|" & cbo门诊可否分零.Text
            If strTemp <> mstr所有记录 Then
                If MsgBox("有数据被修改了确定退出？", vbYesNo, gstrSysName) = vbYes Then
                    mblnLoad = False
                    mblnOK = False
                    mblnCancel = False
                    mbln病案费目 = False
                    Unload Me
                Else
                    Cancel = 1
                End If
            Else
                mblnLoad = False
                mblnOK = False
                mblnCancel = False
                mbln病案费目 = False
                Unload Me
            End If
        End If
    End If
    mblnLoad = False
    mblnOK = False
    mblnCancel = False
    mblnOtherSave = False
    mbln病案费目 = False
End Sub


Private Sub msf合同单位_DblClick()
    With Me.msf合同单位
        Me.txt合同单位.Text = .TextMatrix(.Row, 1)
        Me.txt合同单位.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With
    Me.txt合同单位.SetFocus
    Call OS.PressKey(vbKeyTab)
End Sub

Private Sub msf合同单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call msf合同单位_DblClick
End Sub

Private Sub msf合同单位_LostFocus()
    Me.msf合同单位.Visible = False
End Sub

Private Sub opt中药形态_Click(Index As Integer)
    If Index = 0 Then
'        cbo住院可否分零.ListIndex = 0
'        cbo住院可否分零.Enabled = False
'        cbo门诊可否分零.ListIndex = 0
'        cbo门诊可否分零.Enabled = False
    ElseIf Me.stbSpec.Tag = "查阅" Then
        opt中药形态(opt中药形态(0).Tag).Value = True
    Else
        cbo住院可否分零.Enabled = True
        cbo门诊可否分零.Enabled = True
'        If Val(Me.opt中药形态(0).Tag) = 0 Or Val(Me.opt中药形态(0).Tag) <> Index Or Me.stbSpec.Tag = "增加" Then
'            If blnActivate And Me.stbSpec.Tag <> "增加" Then
'                MsgBox "你修改了“中药形态”，系统将强制设定“药品信息”页中“不可分零”使用！", vbInformation, gstrSysName
'            End If
'            cbo住院可否分零.ListIndex = 1
'            cbo门诊可否分零.ListIndex = 1
'        End If
    End If
    Me.opt中药形态(0).Tag = Index
    blnActivate = True
End Sub

Private Sub opt中药形态_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt备选码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.stbSpec.Tab = 1
        If Me.cbo药价属性.Enabled Then
            Me.cbo药价属性.SetFocus
        Else
            Me.txt指导批价.SetFocus
        End If
    End If
    If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub


Private Sub txt备选码_Validate(Cancel As Boolean)
    Dim i As Integer
    
    If Len(Trim(txt备选码.Text)) > 0 Then
        For i = 1 To Len(Trim(txt备选码.Text))
            If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Trim(txt备选码.Text), i, 1)) < 1 Then
                MsgBox "备选码必须是由字母与数字组成。", vbExclamation, gstrSysName
                Me.stbSpec.Tab = 0
                If txt备选码.Enabled And txt备选码.Visible Then
                    txt备选码.SetFocus
                End If
            End If
        Next
    End If
End Sub
Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Asc("-")
        If InStr(1, txt编码.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt标识码_GotFocus()
    Me.txt标识码.SelStart = 0: Me.txt标识码.SelLength = 100
End Sub

Private Sub txt标识码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*_+|=-`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii > 255 Or KeyAscii < 0 Then KeyAscii = 0
End Sub

'Private Sub txt差价让利_GotFocus()
'        Me.txt差价让利.SelStart = 0: Me.txt差价让利.SelLength = 100
'End Sub
'
'Private Sub txt差价让利_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyBack, vbKeyEscape, 3, 22
'        Exit Sub
'    Case vbKeyReturn
'        Call OS.PressKey(vbKeyTab): Exit Sub
'    Case Else
'        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
'    End Select
'    KeyAscii = 0
'End Sub
'
'Private Sub txt差价让利_LostFocus()
'    txt差价让利.Text = Format(txt差价让利.Text, "0.00")
'End Sub

Private Sub txt产地_GotFocus()
    Me.txt产地.SelStart = 0: Me.txt产地.SelLength = 100
End Sub
Private Sub txt产地_KeyPress(KeyAscii As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt产地.hwnd)
    
    If InStr("~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(txt产地))
    If strTemp = "" Then Me.txt产地.Tag = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    On Error GoTo ErrHandle
    gstrSql = "Select 编码 as id,名称,简码" & _
            " From 药品生产商" & _
            " where  编码 Like [1] " & _
            "       Or 名称 Like [1] " & _
            "       Or 简码 Like [1] Order By 编码 "

    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, Me.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrMatch & strTemp & "%")

    If blnCancel = True Then txt产地.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    With rsTmp
        If rsTmp Is Nothing Then
            If Me.txt产地.Tag <> strTemp Then
                If Asc(strTemp) > 0 Then
                    MsgBox "没有找到匹配的生产商，请重新输入！", vbInformation, gstrSysName
                    Me.txt产地.SelStart = 0: Me.txt产地.SelLength = LenB(StrConv(txt产地, vbFromUnicode)): Me.txt产地.Tag = "":
                    Exit Sub
                End If
                If MsgBox("没有找到相关的生产商，增加该生产商吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Me.txt产地.SelStart = 0: Me.txt产地.SelLength = LenB(StrConv(txt产地, vbFromUnicode)): Me.txt产地.Text = "": Me.txt产地.Tag = "": Exit Sub
                Else
                    If zlSureManufacturer = False Then
                        MsgBox "生产商的编码超长，无法自动增加。" & vbCrLf & "请输入或选择现有的药品生产商！", vbInformation, gstrSysName
                        Me.txt产地.Text = "": Me.txt产地.Tag = "": Exit Sub
                    Else
                        Me.txt产地.Tag = Me.txt产地: Call OS.PressKey(vbKeyTab): Exit Sub
                    End If
                End If
            End If
            Exit Sub
        End If
        
        txt产地.SetFocus
        txt产地 = !名称
        txt产地.Tag = txt产地
            
    End With
    
    Call OS.PressKey(vbKeyTab)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt原产地_GotFocus()
    Me.txt原产地.SelStart = 0: Me.txt原产地.SelLength = 100
End Sub
Private Sub txt原产地_KeyPress(KeyAscii As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt原产地.hwnd)

    If InStr("~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(txt原产地))
    If strTemp = "" Then Me.txt原产地.Tag = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    On Error GoTo ErrHandle
    gstrSql = "Select 编码 as id,名称,简码" & _
            " From 药品生产商" & _
            " where  编码 Like [1] " & _
            "       Or 名称 Like [1] " & _
            "       Or 简码 Like [1] Order By 编码 "
'    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, Me.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrMatch & strTemp & "%")

    If blnCancel = True Then txt原产地.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    With rsTmp
        If rsTmp Is Nothing Then
            If Me.txt原产地.Tag <> strTemp Then
                If Asc(strTemp) > 0 Then
                    MsgBox "没有找到匹配的原产地，请重新输入！", vbInformation, gstrSysName
                    Me.txt原产地.SelStart = 0: Me.txt原产地.SelLength = LenB(StrConv(txt原产地, vbFromUnicode)): Me.txt原产地.Tag = "":
                    Exit Sub
                End If
                If MsgBox("没有找到相关的原产地，增加该原产地吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Me.txt原产地.SelStart = 0: Me.txt原产地.SelLength = LenB(StrConv(txt原产地, vbFromUnicode)): Me.txt原产地.Text = "": Me.txt原产地.Tag = "": Exit Sub
                Else
                    If zlSureManufacturer = False Then
                        MsgBox "原产地的编码超长，无法自动增加。" & vbCrLf & "请输入或选择现有的药品生产商！", vbInformation, gstrSysName
                        Me.txt原产地.Text = "": Me.txt原产地.Tag = "": Exit Sub
                    Else
                        Me.txt原产地.Tag = Me.txt原产地: Call OS.PressKey(vbKeyTab): Exit Sub
                    End If
                End If
            End If
            Exit Sub
        End If
        
        txt原产地.SetFocus
        txt原产地 = !名称
        txt原产地.Tag = txt原产地
            
    End With
    
    Call OS.PressKey(vbKeyTab)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt成本价格_GotFocus()
    Me.txt成本价格.SelStart = 0: Me.txt成本价格.SelLength = 100
End Sub

Private Sub txt成本价格_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt成本价格.SelLength = Len(txt成本价格.Text) Then Exit Sub
            If Len(Mid(txt成本价格, InStr(1, txt成本价格.Text, ".") + 1)) >= mintCostDigit And txt成本价格.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Function get分段加成售价(ByVal dbl采购价 As Double) As Double
    Dim blnData As Boolean
    
    mdbl加成率 = 0
    mdbl差价额 = 0
    
    Do Until mrs分段加成.EOF
        If dbl采购价 > mrs分段加成!最低价 And dbl采购价 <= mrs分段加成!最高价 Then
            mdbl加成率 = mrs分段加成!加成率 / 100
            mdbl差价额 = IIf(IsNull(mrs分段加成!差价额), 0, mrs分段加成!差价额)
            blnData = True
            Exit Do
        End If
        mrs分段加成.MoveNext
    Loop
    If blnData = False Then
        MsgBox "没有设置金额段为：" & dbl采购价 & "  的分段加成数据，请在药品目录管理（分段加成率）中设置"
        get分段加成售价 = 0
        Exit Function
    Else
        get分段加成售价 = dbl采购价 * (1 + mdbl加成率) + mdbl差价额
    End If
End Function

Private Sub txt成本价格_LostFocus()
    Dim dblSalePrice As Double
    Dim dbl价格 As Double
    
    Me.txt成本价格.Text = FormatEx(Val(Me.txt成本价格.Text), mintCostDigit, , True)
    txt指导批价.Text = txt成本价格.Text
    If Val(Me.txt当前售价.Text) = 0 And Val(Me.txt成本价格.Text) <> 0 Then
        If mint分段加成 = 0 Then    '按普通加成方式
            dblSalePrice = Val(Me.txt成本价格.Text) * (1 + Val(Me.txt加成率.Text) / 100)
        Else    '按分段加成方式
            dblSalePrice = get分段加成售价(Val(Me.txt成本价格.Text))
        End If
'        dblSalePrice = dblSalePrice + (Val(Me.txt指导售价.Text) - dblSalePrice) * (1 - Val(Me.txt差价让利) / 100)
        If Val(Me.txt指导售价.Text) > 0 Then
            If dblSalePrice > Val(Me.txt指导售价.Text) Then dblSalePrice = Val(Me.txt指导售价.Text)
        End If
        Me.txt当前售价.Text = FormatEx(dblSalePrice, mintPriceDigit, , True)
        
        If mint分段加成 = 1 Then
            Me.txt加成率.Text = Format(mdbl加成率 * 100, "0.00")
        End If
    End If
End Sub

Private Sub txt当前售价_GotFocus()
    Me.txt当前售价.SelStart = 0: Me.txt当前售价.SelLength = 100
End Sub

Private Sub txt当前售价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt当前售价.SelLength = Len(txt当前售价.Text) Then Exit Sub
            If Len(Mid(txt当前售价, InStr(1, txt当前售价.Text, ".") + 1)) >= mintPriceDigit And txt当前售价.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt当前售价_LostFocus()
    Dim dbl成本价 As Double
    Dim dbl指导售价 As Double
    Dim dbl加成率 As Double
    Dim dbl差价率 As Double
'    Dim dbl差价让利 As Double
    Dim dbl现售价 As Double
    
    Me.txt当前售价.Text = FormatEx(Val(txt当前售价), mintPriceDigit, , True)
    txt指导售价.Text = FormatEx(Val(txt当前售价), mintPriceDigit, , True)
    
    dbl现售价 = Val(Me.txt当前售价.Text)
    dbl成本价 = Val(Me.txt成本价格.Text)
    dbl指导售价 = Val(Me.txt指导售价.Text)
'    dbl差价让利 = Val(Me.txt差价让利.Text) / 100
    
    '满足这些条件才计算加成率
    If dbl成本价 > 0 And dbl指导售价 > 0 And dbl现售价 > 0 And dbl现售价 <= dbl指导售价 Then 'And dbl差价让利 <> 0 Then
'        If dbl差价让利 = 1 Then
            dbl加成率 = dbl现售价 / dbl成本价 - 1
'        Else
'            dbl加成率 = ((dbl现售价 - dbl指导售价 * (1 - dbl差价让利)) / dbl差价让利) / dbl成本价 - 1
'        End If
        If dbl加成率 < 0 Then Exit Sub
'        Exit Sub
        If mint分段加成 = 0 Then
            dbl加成率 = dbl加成率 * 100
        Else
            dbl加成率 = mdbl加成率 * 100
        End If
        
        Me.txt加成率.Text = Format(dbl加成率, "0.00")
    End If
    
'    If Trim(txt当前售价.Text) <> "" And Val(Trim(txt指导售价.Text)) = 0 Then
'        txt指导售价.Text = txt当前售价.Text
'    End If
    
'这时根据成本价、加成率、差价让利、指导售价来计算售价的公式
'    Me.txt成本价格.Text = FormatEx(Val(Me.txt成本价格.Text), mintCostDigit)
'    If Val(Me.txt当前售价.Text) = 0 And Val(Me.txt成本价格.Text) <> 0 Then
'        dblSalePrice = Val(Me.txt成本价格.Text) * (1 + Val(Me.txt加成率.Text) / 100)
'        dblSalePrice = dblSalePrice + (Val(Me.txt指导售价.Text) - dblSalePrice) * (1 - Val(Me.txt差价让利) / 100)
'        If dblSalePrice > Val(Me.txt指导售价.Text) Then dblSalePrice = Val(Me.txt指导售价.Text)
'        Me.txt当前售价.Text = FormatEx(dblSalePrice, mintPriceDigit)
'    End If

'根据上面的公式得到加成率基本公式
'    If 让利售价 <= 指导售价 And 差价让利 <> 0 Then
'        If 差价让利 = 1 Then
'           加成率 = 现售价 / 成本价 - 1
'        Else
'           加成率 = ((现售价 - 指导售价 * (1 - 差价让利)) / 差价让利) / 成本价 - 1
'        End If
'    End If
 
'例1
'    成本价 = 1
'    指导售价 = 3
'    加成率 = 0.15
'
'    差价让利 = 0.6
'
'
'    加成售价 = 成本价 * (1 + 加成率) = 1 * (1 + 0.15) = 1.15
'    现售价 = 加成售价 + (指导售价 - 加成售价) * (1 - 差价让利) = 1.15 + (3 - 1.15) * (1 - 0.6) = 1.89

'例2
'    成本价 = 1
'    指导售价 = 3
'    加成率 = 0.20
'
'    差价让利 = 0.6
'
'
'    加成售价 = 成本价 * (1 + 加成率) = 1 * (1 + 0.2) = 1.2
'    现售价 = 加成售价 + (指导售价 - 加成售价) * (1 - 差价让利) = 1.2 + (3 - 1.2) * (1 - 0.6) = 1.92

'例3（差价让利=0）
'    成本价 = 1
'    指导售价 = 3
'    加成率 = 0.20
'
'    差价让利 = 0
'
'
'    加成售价 = 成本价 * (1 + 加成率) = 1 * (1 + 0.2) = 1.2
'    现售价 = 加成售价 + (指导售价 - 加成售价) * (1 - 差价让利) = 1.2 + (3 - 1.2) * (1 - 0) = 3

'例4（差价让利=100）
'    成本价 = 1
'    指导售价 = 3
'    加成率 = 0.20
'
'    差价让利 = 1
'
'
'    加成售价 = 成本价 * (1 + 加成率) = 1 * (1 + 0.2) = 1.2
'    现售价 = 加成售价 + (指导售价 - 加成售价) * (1 - 差价让利) = 1.2 + (3 - 1.2) * (1 - 1) = 1.2
End Sub

Private Sub txt规格_Change()
    Me.txt数字码.Text = zlGetDigitSign(lng药名id, Trim(Me.txt规格.Text))
End Sub

Private Sub txt规格_GotFocus()
    Me.txt规格.SelStart = 0: Me.txt规格.SelLength = 100
End Sub

Private Sub txt规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt合同单位_GotFocus()
    Me.txt合同单位.SelStart = 0: Me.txt合同单位.SelLength = Len(Me.txt合同单位.Text)
End Sub

Private Sub txt合同单位_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt合同单位.hwnd)
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
        
    strTmp = UCase(Trim(Me.txt合同单位.Text))
    
    If strTmp = "" Then
        Me.txt合同单位.Tag = "|"
        Call OS.PressKey(vbKeyTab): Exit Sub
    ElseIf strTmp = Split(Me.txt合同单位.Tag, "|")(1) Then
        Call OS.PressKey(vbKeyTab): Exit Sub
    End If
    On Error GoTo ErrHandle
    gstrSql = "Select 编码,名称,简码,id" & _
            " From 供应商" & _
            " where (编码 Like [1] " & _
            "       Or 名称 Like [2] " & _
            "       Or 简码 Like [2])" & _
            " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    'Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, Me.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, strTmp & "%", gstrMatch & strTmp & "%")

    If blnCancel = True Then txt合同单位.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理

    With rsTmp
        If rsTmp Is Nothing Then
            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
            Me.txt合同单位.Text = Split(Me.txt合同单位.Tag, "|")(1)
            Me.txt合同单位.SelStart = 0: Me.txt合同单位.SelLength = Len(Me.txt合同单位.Text)
            Exit Sub
        End If
        
        txt合同单位.SetFocus
        Me.txt合同单位.Text = rsTmp!名称
        Me.txt合同单位.Tag = rsTmp!ID & "|" & rsTmp!名称

    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt合同单位_Validate(Cancel As Boolean)
    If Me.txt合同单位.Text = "" Then
        Me.txt合同单位.Tag = "|"
    ElseIf Me.txt合同单位.Text <> Split(Me.txt合同单位.Tag, "|")(1) Then
        txt合同单位_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub txt剂量系数_Change()
    If glngSys \ 100 = 8 Then
        Me.txt药房包装 = 1
        'Me.txt住院包装 = 1
    End If
End Sub

Private Sub txt剂量系数_GotFocus()
    Me.txt剂量系数.SelStart = 0: Me.txt剂量系数.SelLength = 100
End Sub

Private Sub txt剂量系数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If InStr(txt剂量系数.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    '只能存在一个小数点
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt剂量系数.SelLength = Len(txt剂量系数.Text) Then Exit Sub
            If Len(Mid(txt剂量系数, InStr(1, txt剂量系数.Text, ".") + 1)) >= 5 And txt剂量系数.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt加成率_Change()
    If Val(txt加成率.Text) > 9900 Then txt加成率.Text = 9900
    If Val(txt加成率.Text) < 0 Then txt加成率.Text = 0
End Sub

Private Sub txt加成率_GotFocus()
    Call zlControl.TxtSelAll(txt加成率)
End Sub



Private Sub txt加成率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If InStr(txt加成率.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            Exit Sub
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt加成率_LostFocus()
    Dim cur价格 As Double
    '重新计算指导差价率和加成率
    Me.txt加成率.Text = Format(txt加成率.Text, "0.00")
End Sub

Private Sub txt结算价_GotFocus()
    Me.txt结算价.SelStart = 0: Me.txt结算价.SelLength = 100
End Sub

Private Sub txt结算价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt结算价.SelLength = Len(txt结算价.Text) Then Exit Sub
            If Len(Mid(txt结算价, InStr(1, txt结算价.Text, ".") + 1)) >= mintCostDigit And txt结算价.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt结算价_LostFocus()
    Me.txt结算价.Text = FormatEx(Val(txt结算价), mintCostDigit, , True)
End Sub

Private Sub txt扣率_Change()
    Me.txt结算价.Text = FormatEx(Val(Me.txt指导批价.Text) * Val(Me.txt扣率.Text) / 100, mintCostDigit, , True)
End Sub

Private Sub txt扣率_GotFocus()
    Me.txt扣率.SelStart = 0: Me.txt扣率.SelLength = 100
End Sub

Private Sub txt扣率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt扣率_LostFocus()
    Me.txt扣率.Text = Format(txt扣率, "0.00000")
End Sub

Private Sub txt送货包装_GotFocus()
    Me.txt送货包装.SelStart = 0: Me.txt送货包装.SelLength = 100
End Sub


Private Sub txt送货包装_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt送货单位_Change()
    Me.lbl送货包装.Caption = "(1" & Me.txt送货单位.Text & "="
    If Trim(txt送货单位.Text) <> "" Then
        txt送货包装.Enabled = True
    Else
        txt送货包装.Enabled = False
        txt送货包装.Text = ""
    End If
End Sub

Private Sub txt送货单位_GotFocus()
    Me.txt送货单位.SelStart = 0: Me.txt送货单位.SelLength = 100
End Sub


Private Sub txt送货单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt送货单位_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt药房包装_GotFocus()
    Me.txt药房包装.SelStart = 0: Me.txt药房包装.SelLength = 100
End Sub

Private Sub txt药房包装_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If InStr(txt药房包装.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    '只能存在一个小数点
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt药房包装.SelLength = Len(txt药房包装.Text) Then Exit Sub
            If Len(Mid(txt药房包装, InStr(1, txt药房包装.Text, ".") + 1)) >= 5 And txt药房包装.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt药房单位_Change()
    Me.lbl药房包装.Caption = "(1" & Me.txt药房单位.Text & "="
    Call cbo申领单位_Click
End Sub

Private Sub txt药房单位_GotFocus()
    Me.txt药房单位.SelStart = 0: Me.txt药房单位.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt药房单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt药房单位_LostFocus()
    Call OS.OpenIme(False)
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

'Private Sub txt拼音_GotFocus()
'    Me.txt拼音.SelStart = 0: Me.txt拼音.SelLength = 100
'End Sub
'
'Private Sub txt拼音_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyBack, vbKeyEscape, 3, 22
'        Exit Sub
'    Case vbKeyReturn
'        Call OS.PressKey(vbKeyTab): Exit Sub
'    Case Else
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
'    End Select
'    KeyAscii = 0
'End Sub

'Private Sub txt商品名_Change()
'    Dim strTmp As String
'    '重新检查名称，并去 掉特殊字符
'    strTmp = MoveSpecialChar(txt商品名.Text)
'    If txt商品名.Text <> strTmp Then
'        txt商品名.Text = strTmp
'    End If
'    Me.txt拼音.Text = zlGetSymbol(strTmp, 0, mint简码长度)
'    Me.txt五笔.Text = zlGetSymbol(strTmp, 1, mint简码长度)
'End Sub
'
'Private Sub txt商品名_GotFocus()
'    Me.txt商品名.SelStart = 0: Me.txt商品名.SelLength = 100
'    Call zlcommfun.OpenIme(True)
'End Sub
'
'Private Sub txt商品名_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Asc("?")
'            KeyAscii = Asc("？")
'        Case Asc("%")
'            KeyAscii = Asc("％")
'        Case Asc("_")
'            KeyAscii = Asc("＿")
'        Case vbKeyReturn
'            Call OS.PressKey(vbKeyTab)
'    End Select
'    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
'    Me.txt拼音.Text = zlGetSymbol(Me.txt商品名.Text, 0, mint简码长度)
'    Me.txt五笔.Text = zlGetSymbol(Me.txt商品名.Text, 1, mint简码长度)
'
'End Sub
'
'Private Sub txt商品名_LostFocus()
'    Call zlcommfun.OpenIme(False)
'End Sub

Private Sub txt申领阀值_GotFocus()
    txt申领阀值.SelStart = 0: txt申领阀值.SelLength = Len(txt申领阀值)
End Sub

Private Sub txt申领阀值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub txt售价单位_Change()
    Me.lbl剂量系数.Caption = "(1" & Me.txt售价单位.Text & "="
    If glngSys \ 100 = 8 Then
        Me.txt药房单位 = Me.txt药房单位
'        Me.txt住院单位 = Me.txt售价单位
    End If
'    Me.lbl住院单位Child.Caption = Me.txt售价单位 & ")"
    Me.lbl药房单位Child.Caption = Me.txt售价单位 & ")"
    Me.lbl药库单位Child.Caption = Me.txt售价单位 & ")"
    Me.lbl申领单位.Caption = Me.txt售价单位 & ")"
    If Val(Me.lbl批价单位(0).Tag) <> 0 Then
        Me.lbl批价单位(0).Caption = "元/" & Me.txt药库单位.Text
        Me.lbl批价单位(1).Caption = "元/" & Me.txt药库单位.Text
    Else
        Me.lbl批价单位(0).Caption = "元/" & Me.txt售价单位.Text
        Me.lbl批价单位(1).Caption = "元/" & Me.txt售价单位.Text
    End If
    Call cbo申领单位_Click
End Sub

Private Sub txt售价单位_GotFocus()
    Me.txt售价单位.SelStart = 0: Me.txt售价单位.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt售价单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt售价单位_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt数字码_GotFocus()
    txt数字码.MaxLength = Val(zlDatabase.GetPara("数字码", glngSys, 1023, 7))
    Me.txt数字码.SelStart = 0: Me.txt数字码.SelLength = 100
End Sub

Private Sub txt数字码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt本位码_GotFocus()
    txt本位码.MaxLength = Val(zlDatabase.GetPara("本位码", glngSys, 1023, 20))
    Me.txt本位码.SelStart = 0: Me.txt本位码.SelLength = 100
End Sub

Private Sub txt本位码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

'Private Sub txt五笔_GotFocus()
'    Me.txt五笔.SelStart = 0: Me.txt五笔.SelLength = 100
'End Sub
'
'Private Sub txt五笔_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyBack, vbKeyEscape, 3, 22
'        Exit Sub
'    Case vbKeyReturn
'        Call OS.PressKey(vbKeyTab): Exit Sub
'    Case Else
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
'    End Select
'    KeyAscii = 0
'End Sub

'Private Sub txt效期_GotFocus()
'    Me.txt效期.SelStart = 0: Me.txt效期.SelLength = 100
'End Sub
'
'Private Sub txt效期_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyBack, vbKeyEscape, 3, 22
'        Exit Sub
'    Case vbKeyReturn
'        Call OS.PressKey(vbKeyTab): Exit Sub
'    Case Else
'        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
'    End Select
'    KeyAscii = 0
'End Sub

Private Sub txt药库包装_GotFocus()
    Me.txt药库包装.SelStart = 0: Me.txt药库包装.SelLength = 100
End Sub

Private Sub txt药库包装_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
'        Me.stbSpec.Tab = 1
'        If Me.cbo药价属性.Enabled Then
'            Me.cbo药价属性.SetFocus
'        Else
'            Me.txt指导批价.SetFocus
'        End If
    Case Else
        If InStr(txt药库包装.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    '只能存在一个小数点
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt药库包装.SelLength = Len(txt药库包装.Text) Then Exit Sub
            If Len(Mid(txt药库包装, InStr(1, txt药库包装.Text, ".") + 1)) >= 5 And txt药库包装.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt药库单位_Change()
    Me.lbl药库包装.Caption = "(1" & Me.txt药库单位.Text & "="
    Me.lbl送货单位child.Caption = txt药库单位.Text & ")"
    If Val(Me.lbl批价单位(0).Tag) <> 0 Then
        Me.lbl批价单位(0).Caption = "元/" & Me.txt药库单位.Text
        Me.lbl批价单位(1).Caption = "元/" & Me.txt药库单位.Text
    Else
        Me.lbl批价单位(0).Caption = "元/" & Me.txt售价单位.Text
        Me.lbl批价单位(1).Caption = "元/" & Me.txt售价单位.Text
    End If
    Call cbo申领单位_Click
End Sub

Private Sub txt药库单位_GotFocus()
    Me.txt药库单位.SelStart = 0: Me.txt药库单位.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt药库单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt药库单位_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt指导批价_Change()
    Me.txt结算价.Text = FormatEx(Val(Me.txt指导批价.Text) * Val(Me.txt扣率.Text) / 100, mintCostDigit, , True)
End Sub

Private Sub txt指导批价_GotFocus()
    Me.txt指导批价.SelStart = 0: Me.txt指导批价.SelLength = 100
End Sub

Private Sub txt指导批价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt指导批价.SelLength = Len(txt指导批价.Text) Then Exit Sub
            If Len(Mid(txt指导批价, InStr(1, txt指导批价.Text, ".") + 1)) >= mintCostDigit And txt指导批价.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt指导批价_LostFocus()
    Me.txt指导批价.Text = FormatEx(Val(txt指导批价), mintCostDigit, , True)
End Sub

Private Sub txt指导售价_GotFocus()
    Me.txt指导售价.SelStart = 0: Me.txt指导售价.SelLength = 100
End Sub

Private Sub txt指导售价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt指导售价.SelLength = Len(txt指导售价.Text) Then Exit Sub
            If Len(Mid(txt指导售价, InStr(1, txt指导售价.Text, ".") + 1)) >= mintPriceDigit And txt指导售价.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt指导售价_LostFocus()
    Me.txt指导售价.Text = FormatEx(Val(txt指导售价), mintPriceDigit, , True)
End Sub

'Private Sub txt住院包装_GotFocus()
'    Me.txt住院包装.SelStart = 0: Me.txt住院包装.SelLength = 100
'End Sub
'
'Private Sub txt住院包装_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyBack, vbKeyEscape, 3, 22
'        Exit Sub
'    Case vbKeyReturn
'        Call OS.PressKey(vbKeyTab): Exit Sub
'    Case Else
'        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
'    End Select
'    KeyAscii = 0
'End Sub
'
'Private Sub txt住院单位_Change()
'    Me.lbl住院包装.Caption = "(1" & Me.txt住院单位.Text & "="
'    Call cbo申领单位_Click
'End Sub
'
'Private Sub txt住院单位_GotFocus()
'    Me.txt住院单位.SelStart = 0: Me.txt住院单位.SelLength = 100
'    Call zlcommfun.OpenIme(True)
'End Sub
'
'Private Sub txt住院单位_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
'    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt住院单位_LostFocus()
'    Call zlcommfun.OpenIme(False)
'End Sub

Private Sub stbSpec_Click(PreviousTab As Integer)
    Select Case stbSpec.Tab
    Case 0
        If Me.txt编码.Enabled Then Me.txt编码.SetFocus
    Case 1
'        If Me.txt指导批价.Enabled Then Me.txt指导批价.SetFocus
        If Me.cbo药价属性.Enabled Then Me.cbo药价属性.SetFocus
    End Select
End Sub

Private Function zlSureManufacturer() As Boolean
    '-------------------------------------------------------------
    '功能：判断是否可继续增加生产商（生产商编码字段宽度为:10）
    '-------------------------------------------------------------
    On Error GoTo ErrHandle
    zlSureManufacturer = False
    gstrSql = "Select Max(编码) 编码 From 药品生产商"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlSureManufacturer")
'        Call SQLTest
    With rsTemp
        If .EOF Then zlSureManufacturer = True: Exit Function
        If IsNull(!编码) Then zlSureManufacturer = True: Exit Function
        
        '如果超长则退出
        strTemp = .Fields(0).Value
        intCount = Len(strTemp)
        strTemp = strTemp + 1
        If Len(strTemp) > 10 Then Exit Function
        If intCount >= Len(strTemp) Then
            strTemp = String(intCount - Len(strTemp), "0") & strTemp
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

Public Function zlGetDigitSign(ByVal lngMediId As Long, ByVal strSpec As String) As String
    '-------------------------------------------------------------
    '功能：根据药品通用名称、剂型的数字标记码和规格前三位数值，产生返回药品七位码
    '入参：strSpellcode-通用名称的拼音码；strDoseCode:剂型的数字标记码, strSpec：规格数值
    '返回：药品简码
    '-------------------------------------------------------------
    Dim rsThis As New ADODB.Recordset
    Dim strSpellcode As String, strDoseCode As String
    Dim strChange As String
    Dim intLocate As Integer
    
    On Error GoTo ErrHandle
    gstrSql = "Select 简码 From 诊疗项目别名 where 诊疗项目id=[1] and 性质=1 and 码类=1"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strSpellcode = IIf(IsNull(rsThis!简码), "", rsThis!简码)
    Else
        strSpellcode = ""
    End If
    
    gstrSql = "select P.标记码 from 药品特性 T,药品剂型 P where T.药品剂型=P.名称(+) and 药名id=[1]"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strDoseCode = IIf(IsNull(rsThis!标记码), "", rsThis!标记码)
    Else
        strDoseCode = ""
    End If

    strChange = "AOEYUVBP MF DT NL GKHJQXZCSRW "
    
    strTemp = ""
    strSpellcode = Mid(strSpellcode, 1, 3)
    For intCount = 1 To Len(strSpellcode)
        intLocate = InStr(1, strChange, Mid(strSpellcode, intCount, 1))
        If intLocate Mod 3 = 0 Then
            intLocate = (intLocate \ 3) - 1
        Else
            intLocate = intLocate \ 3
        End If
        If intLocate <> -1 Then strTemp = strTemp & CStr(intLocate)
    Next
    strTemp = strTemp & strDoseCode & Format(Val(Mid(strSpec, 1, 3)), "000")
    zlGetDigitSign = strTemp
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

Private Sub txt注册商标_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Function CheckUnit() As Boolean
    Dim intOut As Integer, intIN As Integer
    Dim arr单位, arr系数
    Dim str单位 As String, str系数 As String
    Dim str单位_Tmp As String, str系数_Tmp As String
    
    '检查是否存在单位名称一样，但系数不一致的情况
    '检查是否存在系数一样，但单位名称不一样的情况
    str单位 = txt售价单位.Text & "|" & txt药房单位.Text & "|" & txt药库单位.Text
    str系数 = txt剂量系数.Text & "|" & txt药房包装.Text & "|" & txt药库包装.Text
    
    '考虑到其他单位可能与售价单位一致，但系数肯定不一致，所以必须分开判断
    '除售价单位外的检查
    For intOut = 2 To 3
        'str单位_Tmp = IIf(intOut = 1, txt售价单位.Text, IIf(intOut = 2, txt药房单位.Text, txt药库单位.Text))
        str单位_Tmp = IIf(intOut = 2, txt药房单位.Text, txt药库单位.Text)
        'str系数_Tmp = Val(IIf(intOut = 1, txt剂量系数.Text, IIf(intOut = 2, txt药房包装.Text, txt药库包装.Text)))
        str系数_Tmp = Val(IIf(intOut = 2, txt药房包装.Text, txt药库包装.Text))
        arr单位 = Split(str单位, "|")
        arr系数 = Split(str系数, "|")
        For intIN = 2 To 3
            If intIN <> intOut Then
                '单位相同系数不同
                If str单位_Tmp = arr单位(intIN - 1) And (Val(str系数_Tmp) <> Val(arr系数(intIN - 1))) Then
                    MsgBox IIf(intOut = 2, "药房", "药库") & "单位与" & IIf(intIN = 2, "药房", "药库") & "单位一致，但其系数却不相同，请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
                If str单位_Tmp <> arr单位(intIN - 1) And (Val(str系数_Tmp) = Val(arr系数(intIN - 1))) Then
                    MsgBox IIf(intOut = 2, "药房", "药库") & "包装与" & IIf(intIN = 2, "药房", "药库") & "包装一致，但其单位却不相同，请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    Next
    
    '避免其它单位与售价单位相同，但系数不为1的情况
    '各单位与售价单位进行检查
    For intOut = 2 To 3
        str单位_Tmp = IIf(intOut = 2, txt药房单位.Text, txt药库单位.Text)
        str系数_Tmp = Val(IIf(intOut = 2, txt药房包装.Text, txt药库包装.Text))
        If str单位_Tmp = txt售价单位.Text And Val(str系数_Tmp) <> 1 Then
            MsgBox IIf(intOut = 2, "药房", "药库") & "单位与售价单位一致，" & IIf(intOut = 2, "药房", "药库") & "系数应该为1", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    CheckUnit = True
End Function

Private Function CheckRequest() As Boolean
    Dim dbl零售数量 As Double
    Dim str零售数量 As String
    '检查申领阀值转换为零售单位后是否为整数，超过5位小数则提示操作员，可强制保存
    dbl零售数量 = Val(txt申领阀值.Text)
    
    Select Case cbo申领单位.ListIndex
    Case 1 '药房单位
        dbl零售数量 = dbl零售数量 * Val(txt药房包装.Text)
    Case 2 '药库单位
        dbl零售数量 = dbl零售数量 * Val(txt药库包装.Text)
    End Select
    txt申领阀值.Tag = dbl零售数量
    
    CheckRequest = True
End Function

Private Sub txt注册商标_KeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub




Private Sub vsfItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 0 Then Exit Sub
    With vsfItem
        If Col <> .ColIndex("内容") Then Exit Sub
        If .TextMatrix(Row, .ColIndex("内容")) <> .TextMatrix(Row, .ColIndex("原内容")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("内容")) = vbRed
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("内容")) = vbBlack
        End If
    End With
End Sub

Private Sub vsfItem_EnterCell()
    With vsfItem
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .Col = .ColIndex("内容") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub


Private Sub vsfItem_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfItem
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("内容") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("内容")
            End If
        End If
    End With
End Sub


Private Sub vsfItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsfItem.ColIndex("内容") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


