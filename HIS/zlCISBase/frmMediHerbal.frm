VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediHerbal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "中草药品种编辑"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmMediHerbal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2670
      ScaleHeight     =   210
      ScaleWidth      =   5550
      TabIndex        =   92
      Top             =   120
      Width           =   5550
      Begin VB.Label lblFoot 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "注：该药品建立于2002年12月20日，于2003年8月10日停用。"
         Height          =   180
         Left            =   885
         TabIndex        =   93
         Top             =   0
         Width           =   4770
      End
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   120
      TabIndex        =   91
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin TabDlg.SSTab stbSpec 
      Height          =   6075
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "基本信息(&1)"
      TabPicture(0)   =   "frmMediHerbal.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl编码"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl名称"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl单位"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl处方限量"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl价值"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl货源"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl毒理"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl简码"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lbl处方职务"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lbl医保职务"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lbl药品类型"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl别名"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl药库单位Child"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl药库包装"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl药库单位"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl标识码"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl产地"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl码类"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl售价单位"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "申领单位"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl申领阀值"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl申领单位Child"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblComment"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl合同单位"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl说明"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl药房单位"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl售价单位Child"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl剂量系数"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl规格"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl分类"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Lbl梯次"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl药房单位Child"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl药房包装"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl发药类型"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl备选码"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblStationNo"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl适用性别"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "msf别名"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cbo处方职务"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt编码"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt名称"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cbo单位"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt处方限量"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cbo价值"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cbo货源"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cbo毒理"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt拼音"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cbo医保职务"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cbo药品类型"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt五笔"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "chk原料药"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt药库包装"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt药库单位"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt标识码"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt产地"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "chk单独应用"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt申领阀值"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cbo申领单位"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmd参考"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt参考"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cmd合同单位"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt合同单位"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt说明"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt售价单位"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txt药房单位"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt剂量系数"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txt规格"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txt分类"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cmd分类"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cbo梯次"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txt药房包装"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cbo发药类型"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txt备选码"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "cmbStationNo"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "chk免煎药"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "cbo适用性别"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).ControlCount=   78
      TabCaption(1)   =   "药价信息(&2)"
      TabPicture(1)   =   "frmMediHerbal.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl药价属性"
      Tab(1).Control(1)=   "lbl批价单位(1)"
      Tab(1).Control(2)=   "lbl批价单位(0)"
      Tab(1).Control(3)=   "lbl药价级别"
      Tab(1).Control(4)=   "lbl结算价"
      Tab(1).Control(5)=   "lbl扣率"
      Tab(1).Control(6)=   "lblPercent(0)"
      Tab(1).Control(7)=   "lbl服务对象"
      Tab(1).Control(8)=   "lbl费用类型"
      Tab(1).Control(9)=   "lbl当前售价"
      Tab(1).Control(10)=   "lbl收入记入"
      Tab(1).Control(11)=   "lbl指导差率"
      Tab(1).Control(12)=   "lbl指导批价"
      Tab(1).Control(13)=   "lbl指导售价"
      Tab(1).Control(14)=   "lbl加成率"
      Tab(1).Control(15)=   "lbl成本价格"
      Tab(1).Control(16)=   "lblPercent(1)"
      Tab(1).Control(17)=   "lbl管理费比例"
      Tab(1).Control(18)=   "lbl可否分零"
      Tab(1).Control(19)=   "lbl增值税率"
      Tab(1).Control(20)=   "lblPercent(2)"
      Tab(1).Control(21)=   "fra分批核算"
      Tab(1).Control(22)=   "cbo药价属性"
      Tab(1).Control(23)=   "chk屏蔽费别"
      Tab(1).Control(24)=   "txt结算价"
      Tab(1).Control(25)=   "txt扣率"
      Tab(1).Control(26)=   "cbo服务对象"
      Tab(1).Control(27)=   "cbo费用类型"
      Tab(1).Control(28)=   "txt当前售价"
      Tab(1).Control(29)=   "cbo收入记入"
      Tab(1).Control(30)=   "cbo药价级别"
      Tab(1).Control(31)=   "txt指导差率"
      Tab(1).Control(32)=   "txt指导批价"
      Tab(1).Control(33)=   "txt指导售价"
      Tab(1).Control(34)=   "txt加成率"
      Tab(1).Control(35)=   "txt成本价格"
      Tab(1).Control(36)=   "txt管理费比例"
      Tab(1).Control(37)=   "cbo可否分零"
      Tab(1).Control(38)=   "txt增值税率"
      Tab(1).ControlCount=   39
      Begin VB.ComboBox cbo适用性别 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   3660
         Width           =   1455
      End
      Begin VB.CheckBox chk免煎药 
         Caption         =   "免煎药(&K)"
         Height          =   210
         Left            =   6105
         TabIndex        =   124
         Top             =   4370
         Width           =   1305
      End
      Begin VB.TextBox txt增值税率 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   121
         Top             =   3360
         Width           =   1665
      End
      Begin VB.ComboBox cmbStationNo 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   5700
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txt备选码 
         Height          =   300
         Left            =   5550
         MaxLength       =   20
         TabIndex        =   117
         Top             =   5340
         Width           =   2400
      End
      Begin VB.ComboBox cbo发药类型 
         Height          =   300
         Left            =   1230
         TabIndex        =   115
         Text            =   "cbo发药类型"
         Top             =   5340
         Width           =   3120
      End
      Begin VB.ComboBox cbo可否分零 
         Height          =   300
         Left            =   -67560
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   495
         Width           =   1500
      End
      Begin VB.TextBox txt药房包装 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "30"
         Top             =   3720
         Width           =   510
      End
      Begin VB.ComboBox cbo梯次 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   1668
         Width           =   1455
      End
      Begin VB.CommandButton cmd分类 
         Caption         =   "&P"
         Height          =   285
         Left            =   5550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   435
         Width           =   285
      End
      Begin VB.TextBox txt分类 
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   108
         Top             =   450
         Width           =   4275
      End
      Begin VB.TextBox txt规格 
         Height          =   300
         Left            =   1230
         MaxLength       =   40
         TabIndex        =   15
         Top             =   2480
         Width           =   2250
      End
      Begin VB.TextBox txt剂量系数 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "1"
         Top             =   3300
         Width           =   525
      End
      Begin VB.TextBox txt药房单位 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "袋"
         Top             =   3698
         Width           =   540
      End
      Begin VB.TextBox txt售价单位 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   18
         Text            =   "克"
         Top             =   3292
         Width           =   540
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   28
         Top             =   4980
         Width           =   3120
      End
      Begin VB.TextBox txt合同单位 
         Height          =   300
         Left            =   1230
         MaxLength       =   30
         TabIndex        =   27
         Top             =   4605
         Width           =   2820
      End
      Begin VB.CommandButton cmd合同单位 
         Caption         =   "…"
         Height          =   285
         Left            =   4080
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   4605
         Width           =   285
      End
      Begin VB.TextBox txt参考 
         Height          =   300
         Left            =   1230
         TabIndex        =   14
         Top             =   2074
         Width           =   4275
      End
      Begin VB.CommandButton cmd参考 
         Caption         =   "…"
         Height          =   285
         Left            =   5550
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   2074
         Width           =   285
      End
      Begin VB.ComboBox cbo申领单位 
         Height          =   300
         ItemData        =   "frmMediHerbal.frx":05C2
         Left            =   4620
         List            =   "frmMediHerbal.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2480
         Width           =   1215
      End
      Begin VB.TextBox txt申领阀值 
         Height          =   300
         Left            =   4620
         MaxLength       =   8
         TabIndex        =   25
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox txt管理费比例 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   75
         Top             =   1710
         Width           =   1350
      End
      Begin VB.TextBox txt成本价格 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   69
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox txt加成率 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   67
         Text            =   "35.00"
         Top             =   2910
         Width           =   1665
      End
      Begin VB.CheckBox chk单独应用 
         Caption         =   "单味使用(&Q)"
         Height          =   210
         Left            =   6105
         TabIndex        =   49
         Top             =   4080
         Width           =   1305
      End
      Begin VB.TextBox txt产地 
         Height          =   300
         Left            =   4050
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1262
         Width           =   1470
      End
      Begin VB.TextBox txt标识码 
         Height          =   300
         Left            =   4050
         MaxLength       =   29
         TabIndex        =   5
         Top             =   856
         Width           =   1455
      End
      Begin VB.TextBox txt药库单位 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   22
         Text            =   "千克"
         Top             =   4110
         Width           =   540
      End
      Begin VB.TextBox txt药库包装 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "1000"
         Top             =   4110
         Width           =   510
      End
      Begin VB.TextBox txt指导售价 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   63
         Top             =   2100
         Width           =   1665
      End
      Begin VB.TextBox txt指导批价 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   54
         Top             =   900
         Width           =   1665
      End
      Begin VB.TextBox txt指导差率 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   65
         Text            =   "25.92593"
         Top             =   2505
         Width           =   1665
      End
      Begin VB.ComboBox cbo药价级别 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2100
         Width           =   1635
      End
      Begin VB.ComboBox cbo收入记入 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1305
         Width           =   1635
      End
      Begin VB.TextBox txt当前售价 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   71
         Top             =   900
         Width           =   1635
      End
      Begin VB.ComboBox cbo费用类型 
         Height          =   300
         ItemData        =   "frmMediHerbal.frx":05C6
         Left            =   -70350
         List            =   "frmMediHerbal.frx":05C8
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2505
         Width           =   1635
      End
      Begin VB.ComboBox cbo服务对象 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   2910
         Width           =   1635
      End
      Begin VB.TextBox txt扣率 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   57
         Text            =   "100"
         Top             =   1305
         Width           =   1665
      End
      Begin VB.TextBox txt结算价 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   60
         Top             =   1695
         Width           =   1665
      End
      Begin VB.CheckBox chk屏蔽费别 
         Alignment       =   1  'Right Justify
         Caption         =   "屏蔽费别(&M)"
         Height          =   285
         Left            =   -68625
         TabIndex        =   84
         Top             =   960
         Width           =   1305
      End
      Begin VB.ComboBox cbo药价属性 
         Height          =   300
         Left            =   -73905
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   495
         Width           =   1665
      End
      Begin VB.Frame fra分批核算 
         Caption         =   "分批管理"
         Height          =   1875
         Left            =   -68610
         TabIndex        =   85
         Top             =   1335
         Width           =   1845
         Begin VB.CheckBox chk药房 
            Caption         =   "药房分批管理"
            Enabled         =   0   'False
            Height          =   210
            Left            =   180
            TabIndex        =   87
            Top             =   720
            Width           =   1500
         End
         Begin VB.CheckBox chk药库 
            Caption         =   "药库分批管理"
            Height          =   210
            Left            =   180
            TabIndex        =   86
            Top             =   345
            Width           =   1500
         End
      End
      Begin VB.CheckBox chk原料药 
         Caption         =   "原料药(&M)"
         Height          =   210
         Left            =   7470
         TabIndex        =   50
         Top             =   4080
         Width           =   1155
      End
      Begin VB.TextBox txt五笔 
         Height          =   300
         Left            =   4035
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1668
         Width           =   1110
      End
      Begin VB.ComboBox cbo药品类型 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2074
         Width           =   1455
      End
      Begin VB.ComboBox cbo医保职务 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2886
         Width           =   1455
      End
      Begin VB.TextBox txt拼音 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1668
         Width           =   1890
      End
      Begin VB.ComboBox cbo毒理 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   450
         Width           =   1455
      End
      Begin VB.ComboBox cbo货源 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1262
         Width           =   1455
      End
      Begin VB.ComboBox cbo价值 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   856
         Width           =   1455
      End
      Begin VB.TextBox txt处方限量 
         Height          =   300
         Left            =   7170
         MaxLength       =   16
         TabIndex        =   48
         Text            =   "0"
         Top             =   3292
         Width           =   1455
      End
      Begin VB.ComboBox cbo单位 
         Height          =   300
         Left            =   1230
         TabIndex        =   17
         Text            =   "g"
         Top             =   2886
         Width           =   780
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   1230
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1262
         Width           =   1890
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   1230
         MaxLength       =   13
         TabIndex        =   3
         Top             =   856
         Width           =   1890
      End
      Begin VB.ComboBox cbo处方职务 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2480
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit msf别名 
         Height          =   1140
         Left            =   3600
         TabIndex        =   26
         Top             =   3420
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2011
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lbl适用性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "适用性别(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   126
         Top             =   3720
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
         Index           =   2
         Left            =   -72195
         TabIndex        =   123
         Top             =   3420
         Width           =   90
      End
      Begin VB.Label lbl增值税率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "增值税率(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   122
         Top             =   3420
         Width           =   990
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "站点编号(&Z)"
         Height          =   180
         Left            =   135
         TabIndex        =   120
         Top             =   5760
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl备选码 
         AutoSize        =   -1  'True
         Caption         =   "备选码(&F)"
         Height          =   180
         Left            =   4680
         TabIndex        =   118
         Top             =   5400
         Width           =   810
      End
      Begin VB.Label lbl发药类型 
         AutoSize        =   -1  'True
         Caption         =   "发药类型(&H)"
         Height          =   180
         Left            =   135
         TabIndex        =   116
         Top             =   5400
         Width           =   990
      End
      Begin VB.Label lbl可否分零 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分零使用(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -68610
         TabIndex        =   114
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lbl药房包装 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1袋="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2130
         TabIndex        =   113
         Top             =   3780
         Width           =   450
      End
      Begin VB.Label lbl药房单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   112
         Top             =   3780
         Width           =   180
      End
      Begin VB.Label Lbl梯次 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "用药梯次(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   111
         Top             =   1725
         Width           =   990
      End
      Begin VB.Label lbl分类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品分类(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   109
         Top             =   510
         Width           =   990
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
         Left            =   150
         TabIndex        =   107
         Top             =   2540
         Width           =   990
      End
      Begin VB.Label lbl剂量系数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1克="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2130
         TabIndex        =   106
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label lbl售价单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   105
         Top             =   3345
         Width           =   180
      End
      Begin VB.Label lbl药房单位 
         AutoSize        =   -1  'True
         Caption         =   "药房单位(&I)"
         Height          =   180
         Left            =   165
         TabIndex        =   104
         Top             =   3758
         Width           =   990
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         Caption         =   "标识说明(&B)"
         Height          =   180
         Left            =   135
         TabIndex        =   103
         Top             =   5025
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(请填写适当的说明，来表示限用、适用症药品。)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4605
         TabIndex        =   102
         Top             =   5055
         Width           =   3960
      End
      Begin VB.Label lbl合同单位 
         AutoSize        =   -1  'True
         Caption         =   "合同单位(&C)"
         Height          =   180
         Left            =   135
         TabIndex        =   100
         Top             =   4650
         Width           =   990
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "(指定了合同单位，药品就只能按合同单位入库。)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4590
         TabIndex        =   99
         Top             =   4680
         Width           =   3945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "参考项目(&F)"
         Height          =   180
         Left            =   165
         TabIndex        =   97
         Top             =   2134
         Width           =   990
      End
      Begin VB.Label lbl申领单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   5520
         TabIndex        =   95
         Top             =   2910
         Width           =   300
      End
      Begin VB.Label lbl申领阀值 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申领阀值(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3585
         TabIndex        =   94
         Top             =   2910
         Width           =   990
      End
      Begin VB.Label 申领单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申领单位(&W)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3585
         TabIndex        =   33
         Top             =   2535
         Width           =   990
      End
      Begin VB.Label lbl管理费比例 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "管理费比例(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71550
         TabIndex        =   74
         Top             =   1770
         Width           =   1170
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
         Left            =   -68910
         TabIndex        =   76
         Top             =   1770
         Width           =   90
      End
      Begin VB.Label lbl售价单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "售价单位(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   3352
         Width           =   990
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
         Left            =   -71385
         TabIndex        =   68
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lbl加成率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "加成率(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74745
         TabIndex        =   66
         Top             =   2970
         Width           =   810
      End
      Begin VB.Label lbl码类 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(拼音)                (五笔)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   12
         Top             =   1710
         Width           =   2520
      End
      Begin VB.Label lbl产地 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "产地(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3210
         TabIndex        =   8
         Top             =   1322
         Width           =   630
      End
      Begin VB.Label lbl标识码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标识码(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3210
         TabIndex        =   4
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lbl药库单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药库单位(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   29
         Top             =   4170
         Width           =   990
      End
      Begin VB.Label lbl药库包装 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1千克="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1950
         TabIndex        =   30
         Top             =   4170
         Width           =   630
      End
      Begin VB.Label lbl药库单位Child 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   31
         Top             =   4170
         Width           =   180
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
         Left            =   -74925
         TabIndex        =   62
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lbl指导批价 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "采购限价(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   53
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbl指导差率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "指导差率(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   64
         Top             =   2565
         Width           =   990
      End
      Begin VB.Label lbl收入记入 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收入项目(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   72
         Top             =   1365
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
         Left            =   -71385
         TabIndex        =   70
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbl费用类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保分型(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   79
         Top             =   2565
         Width           =   990
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
         Left            =   -71385
         TabIndex        =   81
         Top             =   2970
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
         Left            =   -72195
         TabIndex        =   58
         Top             =   1365
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
         Left            =   -74925
         TabIndex        =   56
         Top             =   1365
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
         Left            =   -74745
         TabIndex        =   59
         Top             =   1755
         Width           =   810
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
         Left            =   -71385
         TabIndex        =   77
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lbl批价单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元/g"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   -72195
         TabIndex        =   55
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lbl批价单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元/g"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   -72195
         TabIndex        =   61
         Top             =   1755
         Width           =   360
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
         Left            =   -74925
         TabIndex        =   51
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lbl别名 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其他别名(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3615
         TabIndex        =   34
         Top             =   3210
         Width           =   990
      End
      Begin VB.Label Lbl药品类型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品类型(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   41
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label Lbl医保职务 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保职务(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   45
         Top             =   2940
         Width           =   990
      End
      Begin VB.Label Lbl处方职务 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方职务(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   43
         Top             =   2535
         Width           =   990
      End
      Begin VB.Label lbl简码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称简码(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   1728
         Width           =   990
      End
      Begin VB.Label Lbl毒理 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "毒理分类(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   35
         Top             =   510
         Width           =   990
      End
      Begin VB.Label Lbl货源 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "货源情况(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   39
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Lbl价值 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "价值分类(&V)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   37
         Top             =   915
         Width           =   990
      End
      Begin VB.Label Lbl处方限量 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方限量(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   47
         Top             =   3345
         Width           =   990
      End
      Begin VB.Label Lbl单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "剂量单位(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   2946
         Width           =   990
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品名称(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   6
         Top             =   1322
         Width           =   990
      End
      Begin VB.Label lbl编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品编码(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   2
         Top             =   916
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3135
      Top             =   4125
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
            Picture         =   "frmMediHerbal.frx":05CA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbal.frx":0B64
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   88
      Top             =   6300
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frmMediHerbal.frx":10FE
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7950
      TabIndex        =   89
      Top             =   6300
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf合同单位 
      Height          =   1845
      Left            =   3675
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   6720
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
End
Attribute VB_Name = "frmMediHerbal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、编辑状态：由Me.stbSpec.Tag存放，分别为"增加"、"修改"、"查阅"，由上级程序传入
'---------------------------------------------------
Public lng分类id As Long        '被编辑的药品分类ID，上级程序传递进入
Public lng药名ID As Long        '修改和、查询时由外部程序传递进入
Public strPrivs As String       '当前用户对本程序的权限，由上级别程序传递进入

Private lng药品ID As Long       '修改或查询时根据传递进入的参数lng药名ID查找到
Private mint编码规则 As Integer     '药品品种编码的产生规则

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim ObjItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mstrMatch As String, strRefer As String '参考名称
Dim mblnUsed As Boolean         '是否已使用

Private mlng编码长度 As Long
Private mlng规格长度 As Long
Private mlng产地长度 As Long
Private mlng说明长度 As Long
Private mlng名称长度 As Long
Private mint简码长度 As Integer
Private mint备选码长度 As Integer

'从参数表中取药品价格小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数

Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer
Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    
    gstrSql = "Select A.编码, A.名称, A.规格, A.说明, A.产地, B.简码, A.备选码 " & _
        " From 收费项目目录 A, 收费项目别名 B " & _
        " Where A.ID = B.收费细目id And A.ID = 0 And B.码类 = 1 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng编码长度 = rsTmp.Fields("编码").DefinedSize
    mlng规格长度 = rsTmp.Fields("规格").DefinedSize
    mlng产地长度 = rsTmp.Fields("产地").DefinedSize
    mlng说明长度 = rsTmp.Fields("说明").DefinedSize
    mlng名称长度 = rsTmp.Fields("名称").DefinedSize
    mint简码长度 = rsTmp.Fields("简码").DefinedSize
    mint备选码长度 = rsTmp.Fields("备选码").DefinedSize
    
    txt规格.MaxLength = mlng规格长度
    txt产地.MaxLength = mlng产地长度
    txt说明.MaxLength = mlng说明长度
    txt备选码.MaxLength = mint备选码长度
       
    gstrSql = "Select A.名称, A.编码, B.简码 From 诊疗项目目录 A, 诊疗项目别名 B " & _
            " Where A.ID = B.诊疗项目id And A.ID = 0 And B.码类 = 1"
    Call zldatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
        
    If mlng编码长度 > rsTmp.Fields("编码").DefinedSize Then
        mlng编码长度 = rsTmp.Fields("编码").DefinedSize
    End If
    
    If mlng名称长度 > rsTmp.Fields("名称").DefinedSize Then
        mlng名称长度 = rsTmp.Fields("名称").DefinedSize
    End If
        
    If mint简码长度 > rsTmp.Fields("简码").DefinedSize Then
        mint简码长度 = rsTmp.Fields("简码").DefinedSize
    End If
        
    txt编码.MaxLength = mlng编码长度
    txt名称.MaxLength = mlng名称长度
    txt拼音.MaxLength = mint简码长度
    txt五笔.MaxLength = mint简码长度
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSQL As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    strSQL = "Select 类型 From 诊疗分类目录 Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng分类id)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSQL = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 编码"
    Else
        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
            " Where A.ID=B.参考目录ID And A.类型=" & iAttr & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.简码) Like '" & mstrMatch & UCase(strName) & "%')"

        strSQL = " Select Distinct A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
    End If
    Set SelectRefer = zldatabase.ShowSelect(Me, strSQL, 0, "参考", , , , , True)
End Function

Private Sub cbo处方职务_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo单位_Change()
    Me.lbl售价单位Child.Caption = Me.cbo单位.Text & ")"
End Sub

Private Sub cbo单位_Click()
    Me.lbl售价单位Child.Caption = Me.cbo单位.Text & ")"
End Sub

Private Sub cbo单位_GotFocus()
    Me.cbo单位.SelStart = 0: Me.cbo单位.SelLength = 100
End Sub

Private Sub cbo单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbo毒理_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo发药类型_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cbo费用类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo服务对象_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo货源_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo价值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo申领单位_Click()
    Select Case cbo申领单位.ListIndex
    Case 0
        lbl申领单位Child.Caption = txt售价单位.Text & ")"
    Case 1
        lbl申领单位Child.Caption = txt药房单位.Text & ")"
    Case 2
        lbl申领单位Child.Caption = txt药库单位.Text & ")"
    End Select
End Sub

Private Sub cbo申领单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub cbo收入记入_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo梯次_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo药价级别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo药价属性_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo药品类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo医保职务_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk单独应用_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk屏蔽费别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk药库_Click()
    Dim blnEnable As Boolean
    
    '在药库分批的前提下，如果药房没有库存，则可设置其是否分批
    gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
             " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
    
    With rsTemp
        blnEnable = True
        If .Fields(0).Value <> 0 Then
            blnEnable = False
        End If
    End With
    If Me.chk药库.Value = 0 Then
        Me.chk药房.Value = 0: Me.chk药房.Enabled = False
    Else
        Me.chk药房.Enabled = True
    End If
End Sub

Private Sub chk药库_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk原料药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.stbSpec.Tab = 1
        If Me.cbo药价属性.Enabled Then
            Me.cbo药价属性.SetFocus
        Else
            Me.txt指导批价.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub IniStationNo()
    lblStationNo.Visible = False
    cmbStationNo.Visible = False
    
    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        With cmbStationNo
            .Clear
            .AddItem ""
            .AddItem "0"
            .AddItem "1"
            .AddItem "2"
            .AddItem "3"
            .AddItem "4"
            .AddItem "5"
            .AddItem "6"
            .AddItem "7"
            .AddItem "8"
            .AddItem "9"
            
            .ListIndex = 0
        End With
    End If
End Sub

Private Sub SetStationNo(ByVal strNO As String)
    Dim n As Integer
    
    If gstrNodeNo = "-" Then Exit Sub
    
    If strNO = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If cmbStationNo.List(n) = strNO Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub
Private Sub cmdOK_Click()
    Dim dbl指导售价 As Double, dbl当前售价 As Double, dbl成本价格 As Double
    Dim rsData As ADODB.Recordset
    Dim blnPackerReturn As Boolean
    
    '编辑数据检查
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt编码.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt编码.Text, vbFromUnicode)) > mlng编码长度 Then MsgBox "编码超长(最多" & mlng编码长度 & "个字符)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > mlng名称长度 Then MsgBox "名称超长（最多" & mlng名称长度 & "个字符或" & Int(mlng名称长度 / 2) & "个汉字）！", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.txt名称.SetFocus: Exit Sub
    If Trim(Me.cbo单位.Text) = "" Then MsgBox "请输入剂量单位！", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.cbo单位.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.cbo单位.Text), vbFromUnicode)) > 6 Then MsgBox "剂量单位的超长（最多6个字符或3个汉字）！", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.cbo单位.SetFocus: Exit Sub
    
    If LenB(StrConv(Me.txt备选码.Text, vbFromUnicode)) > mint备选码长度 Then MsgBox "备选码超长(最多" & mint备选码长度 & "个字符)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt备选码.SetFocus: Exit Sub
    
    If Trim(Me.txt售价单位.Text) = "" Then MsgBox "请输入售价单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt售价单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt售价单位.Text, vbFromUnicode)) > 8 Then MsgBox "售价单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt售价单位.SetFocus: Exit Sub
    If Val(Me.txt剂量系数.Text) = 0 Then MsgBox "剂量系数错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt剂量系数.SetFocus: Exit Sub
    If Val(Me.txt剂量系数.Text) >= 100000 Then MsgBox "剂量系数超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt剂量系数.SetFocus: Exit Sub
    
    If Trim(Me.txt药房单位.Text) = "" Then MsgBox "请输入药房单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt药房单位.Text, vbFromUnicode)) > 8 Then MsgBox "药房单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房单位.SetFocus: Exit Sub
    If Val(Me.txt药房包装.Text) = 0 Then MsgBox "药房包装错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房包装.SetFocus: Exit Sub
    If Val(Me.txt药房包装.Text) >= 100000 Then MsgBox "药房包装超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药房包装.SetFocus: Exit Sub
    
    strTemp = IIf(glngSys \ 100 <> 8, "药库", "采购")
    If Trim(Me.txt药库单位.Text) = "" Then MsgBox "请输入" & strTemp & "单位！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库单位.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt药库单位.Text, vbFromUnicode)) > 8 Then MsgBox strTemp & "单位超长(最多8个字符或4个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库单位.SetFocus: Exit Sub
    If Val(Me.txt药库包装.Text) = 0 Then MsgBox strTemp & "包装错误(不能为0)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库包装.SetFocus: Exit Sub
    If Val(Me.txt药库包装.Text) >= 100000 Then MsgBox strTemp & "包装超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt药库包装.SetFocus: Exit Sub
    
    If Val(Me.txt申领阀值.Text) < 0 Then MsgBox strTemp & "申领阀值不能小于零！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt申领阀值.SetFocus: Exit Sub
    If Val(Me.txt申领阀值.Text) >= 100000 Then MsgBox strTemp & "申领阀值超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt申领阀值.SetFocus: Exit Sub
    
    If Val(Me.txt指导批价.Text) = 0 And mblnUsed = True Then
        MsgBox "请输入指导批价！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导批价.Enabled Then Me.txt指导批价.SetFocus: Exit Sub
    End If
    If Val(Me.txt指导批价.Text) > 1000000 Then
        MsgBox "指导批价超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
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
'    If Val(Me.txt指导差率.Text) = 0 Then
'        MsgBox "请输入指导差率！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
'        If Me.txt指导差率.Enabled Then Me.txt指导差率.SetFocus: Exit Sub
'    End If
    If Val(Me.txt指导差率.Text) > 100 Then
        MsgBox "指导差率超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txt指导差率.Enabled Then Me.txt指导差率.SetFocus: Exit Sub
    End If
    If Val(Me.txt扣率.Text) = 0 Then MsgBox "请输入扣率！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt扣率.SetFocus: Exit Sub
    If Val(Me.txt扣率.Text) > 100 Then MsgBox "扣率超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt扣率.SetFocus: Exit Sub
    If Val(Me.txt管理费比例.Text) < 0 Then MsgBox "管理费比例不能小于零！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt管理费比例.SetFocus: Exit Sub
    If Val(Me.txt管理费比例.Text) > 100 Then MsgBox "管理费比例超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt管理费比例.SetFocus: Exit Sub
    
    If Val(Me.txt增值税率.Text) < 0 Then MsgBox "增值税率比例不能小于零！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt增值税率.SetFocus: Exit Sub
    If Val(Me.txt增值税率.Text) > 100 Then MsgBox "增值税率比例超过最大值！", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt增值税率.SetFocus: Exit Sub
    
    If Me.cbo药价属性.ItemData(cbo药价属性.ListIndex) = 0 Then
'        If Val(Me.txt当前售价.Text) = 0 And Me.txt当前售价.Enabled = True Then
'            MsgBox "请输入当前售价！", vbInformation, gstrSysName
'            Me.stbSpec.Tab = 1
'            If Me.txt当前售价.Enabled Then Me.txt当前售价.SetFocus
'            Exit Sub
'        End If
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
    
    If LenB(StrConv(Me.txt产地.Text, vbFromUnicode)) > 60 Then MsgBox "产地超长(最多60个字符或30个汉字)！", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt产地.SetFocus: Exit Sub
    
    '别名检查
    strTemp = ";" & Trim(Me.txt名称.Text)
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "别名存在重复（包括名称）！", vbInformation, gstrSysName
                    stbSpec.Tab = 0: .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    '数据保存
    If Me.stbSpec.Tag = "增加" Then
        lng药名ID = zldatabase.GetNextId("诊疗项目目录")
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text)) = True Then Exit Sub
        If zlExseCodeRepeat(Trim(Me.txt编码.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text), lng药名ID) = True Then Exit Sub
    End If
    If Not CheckRequest Then Exit Sub
    
    gstrSql = Me.txt分类.Tag & "," & lng药名ID & ",'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt标识码.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt拼音.Text) & "','" & Trim(Me.txt五笔.Text) & "','" & Trim(Me.txt产地.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.cbo单位.Text) & "','" & Trim(Me.txt规格.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt售价单位.Text) & "'," & Val(Me.txt剂量系数.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txt药房单位.Text) & "'," & Val(Me.txt药房包装.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txt药房单位.Text) & "'," & Val(Me.txt药房包装.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txt药库单位.Text) & "'," & Val(Me.txt药库包装.Text)
    gstrSql = gstrSql & "," & IIf(cbo申领单位.ListIndex = 0, 1, IIf(cbo申领单位.ListIndex = 2, 4, 3)) '申领单位（1-零售单位;2-住院单位;3-药房单位;4-药库单位），中草药只有1,4
    gstrSql = gstrSql & "," & Val(txt申领阀值.Tag)           '始终以零售单位保存
    gstrSql = gstrSql & ",'" & Mid(Me.cbo毒理.Text, InStr(1, Me.cbo毒理.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo价值.Text, InStr(1, Me.cbo价值.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo货源.Text, InStr(1, Me.cbo货源.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo梯次.Text, InStr(1, Me.cbo梯次.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Left(Me.cbo药品类型.Text, 1) & ",'" & Left(Me.cbo处方职务.Text, 1) & Left(Me.cbo医保职务.Text, 1) & "'"
    gstrSql = gstrSql & "," & Val(Trim(Me.txt处方限量.Text)) & "," & Me.chk单独应用.Value & "," & Me.chk原料药.Value
    
    gstrSql = gstrSql & "," & Me.cbo药价属性.ItemData(Me.cbo药价属性.ListIndex)
    If Val(Me.lbl批价单位(0).Tag) <> 0 Then
        dbl指导售价 = Round(Val(txt指导售价.Text) / Val(txt药库包装.Text), mintSalePriceDigit)
        dbl当前售价 = Round(Val(txt当前售价.Text) / Val(txt药库包装.Text), mintSalePriceDigit)
        dbl成本价格 = Round(Val(txt成本价格.Text) / Val(txt药库包装.Text), mintSaleCostDigit)
        gstrSql = gstrSql & "," & Round(Val(Me.txt指导批价.Text) / Val(Me.txt药库包装), mintSaleCostDigit)
    Else
        dbl当前售价 = Round(Val(txt当前售价.Text), mintPriceDigit)
        dbl指导售价 = Round(Val(txt指导售价.Text), mintPriceDigit)
        dbl成本价格 = Round(Val(txt成本价格.Text), mintCostDigit)
        gstrSql = gstrSql & "," & Val(Me.txt指导批价.Text)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt扣率.Text) & "," & dbl指导售价 & "," & Val(Me.txt指导差率.Text) & "," & Val(Me.txt管理费比例.Text)
    gstrSql = gstrSql & ",'" & Mid(Me.cbo药价级别.Text, InStr(1, Me.cbo药价级别.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo费用类型.Text, InStr(1, Me.cbo费用类型.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Me.cbo服务对象.ItemData(Me.cbo服务对象.ListIndex) & "," & Me.chk屏蔽费别.Value
    gstrSql = gstrSql & "," & Me.chk药库 & "," & Me.chk药房
    gstrSql = gstrSql & "," & IIf(Val(Me.txt参考.Tag) = 0, "NULL", Val(Me.txt参考.Tag))
    strTemp = ""
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    '检查别名串长度
    If LenB(strTemp) > 4000 Then
        msf别名.SetFocus
        MsgBox "别名字符串太长，请减少别名个数或者别名长度。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    gstrSql = gstrSql & ",'" & strTemp & "'"
    gstrSql = gstrSql & "," & dbl成本价格
    gstrSql = gstrSql & "," & dbl当前售价
    gstrSql = gstrSql & "," & Me.cbo收入记入.ItemData(Me.cbo收入记入.ListIndex)
    gstrSql = gstrSql & "," & IIf(Split(Me.txt合同单位.Tag, "|")(0) = "", "NULL", Split(Me.txt合同单位.Tag, "|")(0))
    gstrSql = gstrSql & ",'" & Me.txt说明.Text & "'"
    gstrSql = gstrSql & "," & Me.cbo可否分零.ItemData(Me.cbo可否分零.ListIndex)
    gstrSql = gstrSql & ",'" & cbo发药类型.Text
    gstrSql = gstrSql & "','" & txt备选码.Text & "'"
    gstrSql = gstrSql & "," & Val(Me.txt增值税率.Text)
    gstrSql = gstrSql & "," & Me.chk免煎药.Value
    gstrSql = gstrSql & "," & Left(Me.cbo适用性别.Text, 1)
    gstrSql = gstrSql & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", cmbStationNo.Text)
    
    If Me.stbSpec.Tag = "增加" Then
        gstrSql = "zl_草药药品_INSERT(" & gstrSql & ")"
    Else
        gstrSql = "zl_草药药品_Update(" & gstrSql & ")"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    '传送数据到药品分包接口数据库
    If gblnStartPacker = True And gblnPackerConnect = True Then
        gstrSql = "Select 药品id From 药品规格 Where 药名id = [1] "
        Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取药品ID", lng药名ID)
        If Not rsData.EOF Then
            blnPackerReturn = gobjPacker.TranDrugSingle(gcnOracle, Val(rsData!药品ID))
        End If
    End If
    
    If Me.stbSpec.Tag = "增加" And Val(zldatabase.GetPara("品种增加模式", glngSys, 1023, 0)) = 1 Then
        Call frmMediLists.zlRefRecords(lng药名ID)
        lng药名ID = 0
        Call Form_Activate
        Me.txt分类.SetFocus
    Else
        Unload Me
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hWnd, "frmMediItem", Int((glngSys) / 100)
End Sub

Private Sub cmd参考_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
    End If
End Sub

Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left + Me.stbSpec.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height + Me.stbSpec.Top
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub cmd合同单位_Click()
    With rsTemp
        gstrSql = "Select 编码,名称,简码,id" & _
        " From 供应商" & _
        " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By 编码 "
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .EOF Then
            MsgBox "请初始化供应商（字典管理）！", vbInformation, gstrSysName
            Me.txt合同单位.Tag = "|": Me.txt合同单位.SetFocus: Exit Sub
        End If
        With Me.msf合同单位
            .Left = Me.stbSpec.Left + Me.txt合同单位.Left
            .Top = Me.stbSpec.Top + Me.txt合同单位.Top - Me.msf合同单位.Height
            .Clear
            Set .DataSource = rsTemp
            .ColWidth(0) = 800: .ColWidth(1) = 1500: .ColWidth(2) = 800: .ColWidth(3) = 0
            .Row = 1: .ColSel = .Cols - 1
            .ZOrder 0: .Visible = True: .SetFocus
        End With
    End With
End Sub

Private Sub Form_Activate()
    Dim blnExit As Boolean, strMsg As String, strCode As String
    '基础数据检测
    If Me.cbo毒理.ListCount = 0 Then
        strMsg = "无毒理分类数据，请联系系统管理员"
        blnExit = True
    End If
    If Me.cbo价值.ListCount = 0 And Not blnExit Then
        strMsg = "无价值分类数据，请联系系统管理员"
        blnExit = True
    End If
    If Me.cbo货源.ListCount = 0 And Not blnExit Then
        strMsg = "无货源分类数据，请联系系统管理员"
        blnExit = True
    End If
    If Me.cbo梯次.ListCount = 0 And Not blnExit Then
        strMsg = "无用药梯次数据，请联系系统管理员"
        blnExit = True
    End If
    If Me.cbo费用类型.ListCount = 0 And Not blnExit Then
        strMsg = "未设置用于药品的医保分型（字典管理）"
        blnExit = True
    End If
    If Me.cbo收入记入.ListCount = 0 And Not blnExit Then
        strMsg = "未设置明细的收入项目！"
        blnExit = True
    End If
    If Me.stbSpec.Tag = "增加" And Val(Me.lbl收入记入.Tag) = 0 Then
        strMsg = "没有设置“中草药”对应的收入项目（参数设置）！"
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
    
    '----------数据装载-------------------------------------
    Me.tvwClass.Nodes("_" & lng分类id).Selected = True
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    
    lng药品ID = 0
    
    gstrSql = "select I.分类ID,S.药品ID,I.编码,I.名称,I.计算单位 剂量单位,S.药库单位,S.剂量系数,S.药库包装,C.产地,S.标识码," & _
            "        T.毒理分类,T.货源情况,T.价值分类,T.用药梯次,S.申领单位,S.申领阀值," & _
            "        nvl(T.药品类型,0) as 药品类型,nvl(T.处方职务,'00') as 处方职务,nvl(T.处方限量,0) as 处方限量," & _
            "        nvl(T.是否原料,0) as 是否原料,nvl(I.单独应用,0) as 单独应用," & _
            "        C.是否变价,S.指导批发价,S.扣率,S.指导零售价,S.指导差价率,S.管理费比例,S.成本价," & _
            "        S.药价级别,C.费用类型,C.服务对象,C.屏蔽费别,S.可否分零,S.发药类型," & _
            "        S.药库分批,S.药房分批,S.门诊单位,S.门诊包装,C.规格,C.计算单位,C.备选码," & _
            "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,B.名称 as 参考名称,I.参考目录id,S.合同单位id,G.名称 合同单位,C.说明,C.站点,S.增值税率,S.免煎,Nvl(I.适用性别,0) As 适用性别 " & _
            " from 诊疗项目目录 I,药品特性 T,药品规格 S,收费项目目录 C,诊疗参考目录 B,(Select Id,名称 From 供应商 Where 末级 = 1 And substr(类型,1,1) = '1' And " & _
            " 撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) G " & _
            " where I.ID=T.药名ID and T.药名ID=S.药名ID and S.药品ID=C.ID and I.ID=[1] and I.参考目录id=B.id(+) and G.id(+)=S.合同单位id "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名ID)
    
    With rsTemp
        If .RecordCount > 0 Then
            lng药品ID = !药品ID
            Me.lblFoot.Caption = "注：该药品建立于" & Format(!建档时间, "YYYY-MM-DD")
            If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                Me.lblFoot.Caption = Me.lblFoot.Caption & "，于" & Format(!撤档时间, "YYYY-MM-DD") & "停用。"
            End If
            Me.tvwClass.Nodes("_" & !分类id).Selected = True
            Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
            Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            Me.txt编码.Text = !编码
            Me.txt名称.Text = !名称
            Me.txt产地.Text = IIf(IsNull(!产地), "", !产地)
            Me.txt标识码.Text = IIf(IsNull(!标识码), "", !标识码)
            Me.txt备选码.Text = IIf(IsNull(!备选码), "", !备选码)
            
            Me.txt合同单位.Text = IIf(IsNull(!合同单位), "", !合同单位)
            Me.txt合同单位.Tag = IIf(IsNull(!合同单位id), "|", !合同单位id & "|" & !合同单位)
            
            Me.txt说明.Text = IIf(IsNull(!说明), "", !说明)
            
            Me.txt规格.Text = IIf(IsNull(!规格), "", !规格)
            Me.cbo申领单位.ListIndex = IIf(Nvl(!申领单位, 1) = 1, 0, IIf(Nvl(!申领单位, 4) = 4, 2, 1))
            Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0), "#0.00;-#0.00; ;")          '缺省按零售单位显示
            Me.cbo单位.Text = IIf(IsNull(!剂量单位), "", !剂量单位)
            Me.txt售价单位.Text = IIf(IsNull(!计算单位), "", !计算单位)
            Me.lbl药库单位Child.Caption = Me.cbo单位.Text & ")"
            Me.txt药库单位.Text = IIf(IsNull(!药库单位), "", !药库单位)
            Me.lbl药库包装.Caption = "(1" & Me.txt药库单位.Text & "="
            Me.txt参考.Text = Nvl(!参考名称)
            Me.txt参考.Tag = Nvl(!参考目录ID)
            strRefer = Me.txt参考.Text
            Me.txt剂量系数.Text = IIf(IsNull(!剂量系数), 1, !剂量系数)
            Me.txt药库包装.Text = IIf(IsNull(!药库包装), 1, !药库包装)
            
            Me.txt药房单位.Text = IIf(IsNull(!门诊单位), "", !门诊单位)
            Me.txt药房包装.Text = IIf(IsNull(!门诊包装), 1, !门诊包装)
            
            Me.lbl药房单位Child.Caption = Me.txt售价单位 & ")"
            Me.lbl药库单位Child.Caption = Me.txt售价单位 & ")"
            
            Me.cbo发药类型.Text = Nvl(!发药类型)
            Me.cbo适用性别.ListIndex = !适用性别
            
            SetStationNo IIf(IsNull(!站点), "", !站点)
            
            Select Case IIf(IsNull(!可否分零), 0, !可否分零)
            Case 0, 1
                Me.cbo可否分零.ListIndex = IIf(IsNull(!可否分零), 0, !可否分零)
            Case Else
                Me.cbo可否分零.ListIndex = 0
            End Select
            
            '如果是按药库单位定的申领阀值，按药库单位显示
            If Me.cbo申领单位.ListIndex = 1 Then
                Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0) / Val(txt药房包装.Text), "#0.00;-#0.00; ;")
            ElseIf Me.cbo申领单位.ListIndex = 2 Then
                Me.txt申领阀值.Text = Format(Nvl(!申领阀值, 0) / Val(txt药库包装.Text), "#0.00;-#0.00; ;")
            End If
            
            For intCount = 0 To Me.cbo毒理.ListCount - 1
                If Mid(Me.cbo毒理.List(intCount), InStr(1, Me.cbo毒理.List(intCount), "-") + 1) = IIf(IsNull(!毒理分类), "", !毒理分类) Then
                    Me.cbo毒理.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo价值.ListCount - 1
                If Mid(Me.cbo价值.List(intCount), InStr(1, Me.cbo价值.List(intCount), "-") + 1) = IIf(IsNull(!价值分类), "", !价值分类) Then
                    Me.cbo价值.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo货源.ListCount - 1
                If Mid(Me.cbo货源.List(intCount), InStr(1, Me.cbo货源.List(intCount), "-") + 1) = IIf(IsNull(!货源情况), "", !货源情况) Then
                    Me.cbo货源.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo梯次.ListCount - 1
                If Mid(Me.cbo梯次.List(intCount), InStr(1, Me.cbo梯次.List(intCount), "-") + 1) = IIf(IsNull(!用药梯次), "", !用药梯次) Then
                    Me.cbo梯次.ListIndex = intCount: Exit For
                End If
            Next
            Me.cbo药品类型.ListIndex = !药品类型
            Me.cbo处方职务.ListIndex = IIf(CInt(Left(Format(!处方职务, "00"), 1)) <> 9, CInt(Left(Format(!处方职务, "00"), 1)), Me.cbo处方职务.ListCount - 1)
            Me.cbo医保职务.ListIndex = IIf(CInt(Right(Format(!处方职务, "00"), 1)) <> 9, CInt(Right(Format(!处方职务, "00"), 1)), Me.cbo医保职务.ListCount - 1)
            Me.chk单独应用.Value = IIf(!单独应用 = 0, 0, 1)
            Me.chk原料药.Value = IIf(!是否原料 = 0, 0, 1)
            
            Me.cbo药价属性.ListIndex = IIf(IsNull(!是否变价), 0, !是否变价)
            Me.txt扣率.Text = IIf(IsNull(!扣率), 100, !扣率)
            If Val(Me.lbl批价单位(0).Tag) <> 0 = True Then
                Me.txt指导批价.Text = FormatEx(IIf(IsNull(!指导批发价), 0, !指导批发价) * Me.txt药库包装.Text, mintCostDigit)
                Me.txt指导售价.Text = FormatEx(IIf(IsNull(!指导零售价), 0, !指导零售价) * Me.txt药库包装.Text, mintPriceDigit)
                Me.txt成本价格.Text = FormatEx(IIf(IsNull(!成本价), 0, !成本价) * Me.txt药库包装.Text, mintCostDigit)
            Else
                Me.txt指导批价.Text = FormatEx(IIf(IsNull(!指导批发价), 0, !指导批发价), mintCostDigit)
                Me.txt指导售价.Text = FormatEx(IIf(IsNull(!指导零售价), 0, !指导零售价), mintPriceDigit)
                Me.txt成本价格.Text = FormatEx(IIf(IsNull(!成本价), 0, !成本价), mintCostDigit)
            End If
            Me.txt结算价 = FormatEx((Me.txt指导批价.Text) * Me.txt扣率.Text / 100, mintPriceDigit)
            Me.txt指导差率.Text = Format(IIf(IsNull(!指导差价率), 0, !指导差价率), "0.00000")
            Me.txt管理费比例.Text = Format(Nvl(!管理费比例, 0), "#0.00")
            Me.txt增值税率.Text = Format(Nvl(!增值税率, 0), "0.00")
            Me.chk免煎药.Value = IIf(!免煎 = 0, 0, 1)
            '计算指导加成率
            Dim cur价格 As Double
            cur价格 = Val(txt指导差率.Text)
            If cur价格 < 100 Then
                Call Calc(cur价格, True)
                Me.txt加成率.Text = Format(cur价格, "0.00")
            End If
            
            For intCount = 0 To Me.cbo药价级别.ListCount - 1
                If Mid(Me.cbo药价级别.List(intCount), InStr(1, Me.cbo药价级别.List(intCount), "-") + 1) = IIf(IsNull(!药价级别), "", !药价级别) Then
                    Me.cbo药价级别.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo费用类型.ListCount - 1
                If Mid(Me.cbo费用类型.List(intCount), 4) = IIf(IsNull(!费用类型), "", !费用类型) Then
                    Me.cbo费用类型.ListIndex = intCount: Exit For
                End If
            Next
            Me.cbo服务对象.ListIndex = IIf(IsNull(!服务对象), 0, !服务对象)
            Me.chk屏蔽费别.Value = IIf(IsNull(!屏蔽费别), 0, !屏蔽费别)
            
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFoot.Caption = "注：该药品于" & Format(!建档时间, "YYYY年MM月DD日") & "建立，" & Format(!撤档时间, "YYYY年MM月DD日") & "停用"
            Else
                Me.lblFoot.Caption = ""
            End If
            
            Me.chk药房.Tag = IIf(IsNull(!药房分批), 0, !药房分批)
            Me.chk药库.Value = IIf(IsNull(!药库分批), 0, Abs(!药库分批))
            If Me.chk药库.Value = 0 Then
                Me.chk药房.Enabled = False: Me.chk药房.Value = 0
            Else
                Me.chk药房.Enabled = True: Me.chk药房.Value = Me.chk药房.Tag
            End If
        End If
        If Trim(Me.txt合同单位.Tag) = "" Then
            Me.txt合同单位.Tag = "|"
        End If
        If Val(Me.lbl批价单位(0).Tag) <> 0 Then
            Me.lbl批价单位(0).Caption = "元/" & Me.txt药库单位.Text
            Me.lbl批价单位(1).Caption = "元/" & Me.txt药库单位.Text
        Else
            Me.lbl批价单位(0).Caption = "元/" & Me.txt售价单位.Text
            Me.lbl批价单位(1).Caption = "元/" & Me.txt售价单位.Text
        End If
    End With
    
    If Me.stbSpec.Tag = "增加" Then
        '增加时，重新提取编码
        Me.txt编码.Text = "": Me.txt名称.Text = "": Me.txt产地.Text = "": Me.lblFoot.Caption = ""
        lng药名ID = 0
        Me.txt参考 = "": Me.txt参考.Tag = "": strRefer = ""
        If mint编码规则 = 0 Then
            gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 类别 = '7'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): rsTemp.Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
            Me.txt编码.Text = zlcommfun.IncStr(rsTemp!编码)
        Else
            strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
            gstrSql = "select nvl(max(编码),'') as 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 类别 = '7' and 编码 like [1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, "7" & strTemp & "%")
            
            Err = 0: On Error Resume Next
            strTemp = "7" & strTemp
            If Nvl(rsTemp!编码) = "" Then
                Me.txt编码.Text = strTemp & "01"
            Else
                Me.txt编码.Text = zlcommfun.IncStr(rsTemp!编码)
            End If
        End If
    Else
        '正名简码
        gstrSql = "select 名称,性质,简码,码类 from 诊疗项目别名 where 性质 in (1,2) and 诊疗项目ID=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名ID)
        
        Do While Not rsTemp.EOF
            If rsTemp!性质 = 1 And rsTemp!码类 = 1 Then Me.txt拼音.Text = rsTemp!简码
            If rsTemp!性质 = 1 And rsTemp!码类 = 2 Then Me.txt五笔.Text = rsTemp!简码
            rsTemp.MoveNext
        Loop
        '其他别名
        gstrSql = "select N.名称,P.简码 as 拼音,W.简码 as 五笔" & _
                " from (select distinct 名称 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9) N," & _
                "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=" & lng药名ID & " and 性质=9 and 码类=1) P," & _
                "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=" & lng药名ID & " and 性质=9 and 码类=2) W" & _
                " where N.名称=P.名称(+) and N.名称=W.名称(+)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名ID)
        
        With rsTemp
            Do While Not .EOF
                If Me.msf别名.Rows - 1 < .AbsolutePosition Then Me.msf别名.Rows = Me.msf别名.Rows + 1
                Me.msf别名.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                Me.msf别名.TextMatrix(.AbsolutePosition, 1) = !名称
                Me.msf别名.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!拼音), "", !拼音)
                Me.msf别名.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!五笔), "", !五笔)
                .MoveNext
            Loop
        End With
        
        '提取显示当前售价
        If Me.cbo药价属性.ListIndex <> 0 Then
            Me.cbo药价属性.Enabled = False
            gstrSql = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                    " from 收费价目 P," & _
                    "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                    "      From 药品库存 Where 药品ID=[1]) K" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)"
        Else
            '非时价药品调价，取其价格记录中的价格
            gstrSql = "select P.现价,P.收入项目id" & _
                    " from 收费价目 P" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)"
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        With rsTemp
            If .RecordCount > 0 Then
                If Val(Me.lbl批价单位(0).Tag) <> 0 Then
                    Me.txt当前售价.Text = FormatEx(!现价 * Val(Me.txt药库包装.Text), mintPriceDigit)
                Else
                    Me.txt当前售价.Text = FormatEx(!现价, mintPriceDigit)
                End If
        
                For intCount = 0 To Me.cbo收入记入.ListCount - 1
                    If Me.cbo收入记入.ItemData(intCount) = !收入项目id Then
                        Me.cbo收入记入.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        End With

        
        '根据是否有发生，确定：售价单位、药价属性、成本价、零售价格可修改否
        gstrSql = " Select nvl(Count(*),0) " & _
            " From (Select 1 From 药品收发记录 Where 药品ID=[1] And rownum<2" & _
            "       Union Select 1 From 药品库存 Where 药品ID=[1] And rownum<2)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        mblnUsed = False
        If rsTemp.Fields(0).Value > 0 Then
            mblnUsed = True
            If Me.cbo药价属性.ListIndex <> 0 Then Me.cbo药价属性.Enabled = False
            Me.txt成本价格.Enabled = False
            Me.txt当前售价.Enabled = False
            Me.cbo收入记入.Enabled = False
'            Me.txt剂量系数.Enabled = False
            Me.txt药房包装.Enabled = False
            Me.txt药库包装.Enabled = False
        Else
            Me.cbo药价属性.Enabled = True
            Me.txt成本价格.Enabled = True
            Me.txt当前售价.Enabled = True
            Me.cbo收入记入.Enabled = True
'            Me.txt剂量系数.Enabled = True
            Me.txt药房包装.Enabled = True
            Me.txt药库包装.Enabled = True
        End If
        
        '根据是否存在医嘱记录，确定剂量系数是否能够修改
        gstrSql = "Select 1 From 病人医嘱记录 Where 收费细目ID=[1] And Rownum=1"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        If rsTemp.RecordCount > 0 Then
            Me.txt剂量系数.Enabled = False
        Else
            Me.txt剂量系数.Enabled = True
        End If
        
        '根据是否有库存，确定分批特性可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                 " Where A.药品ID=[1] And A.库房ID=B.部门ID And B.工作性质 Like '%药库'"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        If rsTemp.Fields(0).Value > 0 Then
            Me.chk药库.Enabled = False
        Else
            Me.chk药库.Enabled = True
        End If
        If Me.chk药库.Value = 1 Then
            gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                     " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
            
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
        If InStr(1, strPrivs, "目录增删改") = 0 Then
            Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
            Me.txt编码.Enabled = False: Me.txt产地.Enabled = False
            Me.txt名称.Enabled = False: Me.txt拼音.Enabled = False: Me.txt五笔.Enabled = False: Me.msf别名.Active = False
            Me.cbo单位.Enabled = False: Me.txt药库单位.Enabled = False: Me.txt药库包装.Enabled = False
            Me.cbo毒理.Enabled = False: Me.cbo价值.Enabled = False: Me.cbo货源.Enabled = False: Me.cbo梯次.Enabled = False
            Me.cbo药品类型.Enabled = False: Me.cbo处方职务.Enabled = False: Me.txt处方限量.Enabled = False ': Me.txt申领阀值.Enabled = False
            Me.chk原料药.Enabled = False: Me.chk单独应用.Enabled = False
            Me.cbo服务对象.Enabled = False: Me.chk屏蔽费别.Enabled = False
            Me.chk药库.Enabled = False: Me.chk药房.Enabled = False
            Me.txt参考.Enabled = False
            Me.cmd参考.Enabled = False
            Me.txt合同单位.Enabled = False: Me.cmd合同单位.Enabled = False
            Me.txt说明.Enabled = False
            Me.txt规格.Enabled = False
            Me.txt售价单位.Enabled = False
            Me.txt剂量系数.Enabled = False
            Me.txt药房单位.Enabled = False
            Me.txt药房包装.Enabled = False
            Me.cbo可否分零.Enabled = False
            Me.cbo发药类型.Enabled = False
            Me.txt备选码.Enabled = False
            Me.cmbStationNo.Enabled = False
            Me.txt增值税率.Enabled = False
            Me.chk免煎药.Enabled = False
            Me.cbo适用性别.Enabled = False
        End If
        If InStr(1, strPrivs, "医保用药目录") = 0 Then
            Me.cbo医保职务.Enabled = False: Me.cbo费用类型.Enabled = False: Me.txt标识码.Enabled = False
        End If
        If InStr(1, strPrivs, "管理扣率") = 0 Then Me.txt扣率.Enabled = False
        If InStr(1, strPrivs, "指导价格管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.txt指导批价.Text = "0"
                Me.txt指导售价.Text = "0"
            End If
            Me.txt指导差率.Enabled = False: Me.txt加成率.Enabled = False
            Me.txt指导批价.Enabled = False: Me.txt指导售价.Enabled = False
        End If
        If InStr(1, strPrivs, "售价管理") = 0 Then
            If Me.stbSpec.Tag = "增加" Then
                Me.txt当前售价.Text = "0"
                Me.cbo药价属性.ListIndex = 0
            End If
            Me.cbo药价属性.Enabled = False: Me.cbo收入记入.Enabled = False
            Me.txt当前售价.Enabled = False
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
    Else
        cmdOK.Visible = False: cmdCancel.Caption = "关闭(&C)"
        
        Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
        Me.txt编码.Enabled = False: Me.txt标识码.Enabled = False: Me.txt产地.Enabled = False
        Me.txt名称.Enabled = False: Me.txt拼音.Enabled = False: Me.txt五笔.Enabled = False: Me.msf别名.Active = False
        Me.cbo单位.Enabled = False: Me.txt药库单位.Enabled = False: Me.txt药库包装.Enabled = False
        Me.txt申领阀值.Enabled = False: Me.cbo申领单位.Enabled = False
        Me.cbo毒理.Enabled = False: Me.cbo价值.Enabled = False: Me.cbo货源.Enabled = False: Me.cbo梯次.Enabled = False
        Me.cbo药品类型.Enabled = False: Me.cbo处方职务.Enabled = False: Me.cbo医保职务.Enabled = False: Me.txt处方限量.Enabled = False
        Me.chk原料药.Enabled = False: Me.chk单独应用.Enabled = False
        
        Me.cbo药价属性.Enabled = False: Me.txt指导批价.Enabled = False: Me.txt扣率.Enabled = False: Me.txt结算价.Enabled = False
        Me.txt指导售价.Enabled = False: Me.txt指导差率.Enabled = False: Me.txt加成率.Enabled = False
        Me.cbo药价级别.Enabled = False: Me.cbo费用类型.Enabled = False: Me.cbo服务对象.Enabled = False: Me.chk屏蔽费别.Enabled = False
        Me.txt成本价格.Enabled = False: Me.txt当前售价.Enabled = False: Me.cbo收入记入.Enabled = False
        Me.chk药库.Enabled = False: Me.chk药房.Enabled = False: Me.txt管理费比例.Enabled = False
        Me.txt参考.Enabled = False
        Me.cmd参考.Enabled = False
        Me.txt合同单位.Enabled = False: Me.cmd合同单位.Enabled = False
        Me.txt说明.Enabled = False
        Me.txt规格.Enabled = False
        Me.txt售价单位.Enabled = False
        Me.txt剂量系数.Enabled = False
        Me.txt药房单位.Enabled = False
        Me.txt药房包装.Enabled = False
        Me.cbo可否分零.Enabled = False
        Me.cbo发药类型.Enabled = False
        Me.txt备选码.Enabled = False
        Me.cmbStationNo.Enabled = False
        Me.txt增值税率.Enabled = False
        Me.chk免煎药.Enabled = False
        Me.cbo适用性别.Enabled = False
    End If
    
    '如果本次操作是修改，则检查是否存在“药品单位管理”的权限，没有则不允许修改药品单位与系数
    If Me.stbSpec.Tag = "修改" Then
        If InStr(1, strPrivs, "药品单位管理") = 0 Then
            cbo单位.Enabled = False
            txt药库单位.Enabled = False
            txt药库包装.Enabled = False
        End If
    End If
    
    Me.stbSpec.Tab = 0
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Me.tvwClass.Visible Then
            tvwClass.Visible = False: txt分类.SetFocus: Exit Sub
        End If
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    If glngSys \ 100 = 8 Then Me.lbl药库单位.Caption = "采购单位(&W)"
    mint编码规则 = Val(GetSysPara(87))
    
    Call GetDefineSize
    Call IniStationNo
    
    '-------------下拉选择数据装载-----------------------
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        '分类选择树装入
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = 3" & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        
        gstrSql = "select distinct 计算单位 from 诊疗项目目录 where 类别='7' and 计算单位 is not null"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Me.cbo单位.AddItem .Fields(0).Value
            .MoveNext
        Loop
        
        gstrSql = "select 编码||'-'||名称 from 药品毒理分类 order by 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo毒理.Clear
        Do While Not .EOF
            Me.cbo毒理.AddItem .Fields(0).Value
            If InStr(1, .Fields(0).Value, "普通") > 0 Then
                Me.cbo毒理.ListIndex = Me.cbo毒理.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo毒理.ListIndex = -1 And Me.cbo毒理.ListCount > 0 Then Me.cbo毒理.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品价值分类 order by 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo价值.Clear
        Do While Not .EOF
            Me.cbo价值.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo价值.ListCount > 0 Then Me.cbo价值.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品货源情况 order by 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo货源.Clear
        Do While Not .EOF
            Me.cbo货源.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo货源.ListCount > 0 Then Me.cbo货源.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品用药梯次 order by 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo梯次.Clear
        Do While Not .EOF
            Me.cbo梯次.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo梯次.ListCount > 0 Then Me.cbo梯次.ListIndex = 0
    
        gstrSql = "Select 编码||'-'||名称 From 费用类型 where 性质=1 Order By 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo费用类型.Clear
        Me.cbo费用类型.AddItem ""
        Do While Not .EOF
            Me.cbo费用类型.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo费用类型.ListCount > 0 Then Me.cbo费用类型.ListIndex = 0
        
        gstrSql = "Select ID,'['||编码||']'||名称 as 名称" & _
                " From 收入项目" & _
                " where 末级=1 and (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo收入记入.Clear
        Do While Not .EOF
            Me.cbo收入记入.AddItem !名称: Me.cbo收入记入.ItemData(Me.cbo收入记入.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo收入记入.ListCount > 0 Then Me.cbo收入记入.ListIndex = 0
    
        Me.lbl收入记入.Tag = zldatabase.GetPara("中草药收入项目", glngSys, 1023, False)
        For intCount = 0 To Me.cbo收入记入.ListCount - 1
            If Me.cbo收入记入.ItemData(intCount) = Val(Me.lbl收入记入.Tag) Then
                Me.cbo收入记入.ListIndex = intCount: Exit For
            End If
        Next
        
        gstrSql = "Select 名称 From 发药类型 Order By 编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo发药类型.Clear
        Do While Not .EOF
            Me.cbo发药类型.AddItem .Fields(0).Value
            .MoveNext
        Loop
        
        Me.lbl批价单位(0).Tag = Val(GetSysPara(29))
        
        mintCostDigit = GetDigit(1, 1, IIf(Me.lbl批价单位(0).Tag = 0, 1, 4))
        mintPriceDigit = GetDigit(1, 2, IIf(Me.lbl批价单位(0).Tag = 0, 1, 4))
        
        mintSaleCostDigit = GetDigit(1, 1, 1)
        mintSalePriceDigit = GetDigit(1, 2, 1)
    End With
    
    With Me.cbo申领单位
        .Clear
        .AddItem "售价单位"
        .AddItem "药房单位"
        .AddItem "药库单位"
        .ListIndex = 0
    End With
    
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
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo药价级别.Clear
        Do While Not .EOF
            Me.cbo药价级别.AddItem !名称
            .MoveNext
        Loop
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
    
    With Me.cbo药品类型
        .Clear
        aryTemp = Split("0-未设定;1-处方药;2-甲类非处方药;3-乙类非处方药;4-非处方药;5-其它用药", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(intCount)
        Next
        .ListIndex = 0
    End With
    
    Me.cbo处方职务.Clear: Me.cbo医保职务.Clear
    aryTemp = Split("0-不限;1-正高;2-副高;3-中级;4-助理/师级;5-员/士;9-待聘", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo处方职务.AddItem aryTemp(intCount): Me.cbo医保职务.AddItem aryTemp(intCount)
    Next
    Me.cbo处方职务.ListIndex = 0: Me.cbo医保职务.ListIndex = 0
    
    aryTemp = Split("0-无性别区分;1-男性;2-女性", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo适用性别.AddItem aryTemp(intCount)
    Next
    Me.cbo适用性别.ListIndex = 0
    
    With Me.cbo可否分零
        .Clear
        .AddItem "0-可以分零": .ItemData(.NewIndex) = 0
        .AddItem "1-不可分零": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    '----------------编辑界面设置----------------------
    With Me.msf别名
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "药品名称": .TextMatrix(0, 2) = "拼音码": .TextMatrix(0, 3) = "五笔码"
        .ColData(0) = 5: .ColData(1) = 4: .ColData(2) = 4: .ColData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 1000: .ColWidth(2) = 650: .ColWidth(3) = 650
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    strRefer = ""
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msf别名_AfterAddRow(Row As Long)
    With Me.msf别名
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_AfterDeleteRow()
    With Me.msf别名
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_EditKeyPress(KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf别名_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf别名
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Exit Sub
            strTemp = Trim(.Text)
            If strTemp = "" Then Exit Sub
            .TextMatrix(.Row, 1) = strTemp
            .TextMatrix(.Row, 2) = zlGetSymbol(strTemp, 0, mint简码长度)
            .TextMatrix(.Row, 3) = zlGetSymbol(strTemp, 1, mint简码长度)
        End If
    End With
End Sub

Private Sub msf别名_KeyPress(KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf合同单位_DblClick()
    With Me.msf合同单位
        Me.txt合同单位.Text = .TextMatrix(.Row, 1)
        Me.txt合同单位.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With
    Me.txt合同单位.SetFocus
    Call zlcommfun.PressKey(vbKeyTab)
End Sub


Private Sub msf合同单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call msf合同单位_DblClick
End Sub

Private Sub msf合同单位_LostFocus()
    Me.msf合同单位.Visible = False
End Sub





Private Sub stbSpec_Click(PreviousTab As Integer)
    Select Case stbSpec.Tab
    Case 0
        If Me.txt分类.Enabled Then Me.txt分类.SetFocus
    Case 1
        If Me.txt指导批价.Enabled Then Me.txt指导批价.SetFocus
        If Me.cbo药价属性.Enabled Then Me.cbo药价属性.SetFocus
    End Select
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt备选码_KeyPress(KeyAscii As Integer)
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
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
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
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*_+|=-`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii > 255 Or KeyAscii < 0 Then KeyAscii = 0
End Sub

Private Sub txt参考_GotFocus()
    Me.txt参考.SelStart = 0: Me.txt参考.SelLength = 100
End Sub


Private Sub txt参考_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt参考 <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt参考))
            If rsTmp Is Nothing Then
                Me.txt参考 = strRefer
                Me.SetFocus
                Exit Sub
            Else
                Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
            End If
        End If
        Call zlcommfun.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt参考_LostFocus()
    If Me.txt参考 <> strRefer Then
        Me.txt参考 = strRefer
    End If
End Sub


Private Sub txt产地_GotFocus()
    Me.txt产地.SelStart = 0: Me.txt产地.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt产地_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt产地_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt成本价格_GotFocus()
    Me.txt成本价格.SelStart = 0: Me.txt成本价格.SelLength = 100
End Sub

Private Sub txt成本价格_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt成本价格_LostFocus()
    Dim dblSalePrice As Double
    Me.txt成本价格.Text = FormatEx(Val(Me.txt成本价格.Text), mintCostDigit)
    If Val(Me.txt当前售价.Text) = 0 And Val(Me.txt成本价格.Text) <> 0 Then
        dblSalePrice = Val(Me.txt成本价格.Text) * (1 + Val(Me.txt加成率.Text) / 100)
        If dblSalePrice > Val(Me.txt指导售价.Text) Then dblSalePrice = Val(Me.txt指导售价.Text)
        Me.txt当前售价.Text = FormatEx(dblSalePrice, mintPriceDigit)
    End If
End Sub

Private Sub txt处方限量_GotFocus()
    Me.txt处方限量.SelStart = 0: Me.txt处方限量.SelLength = 100
End Sub

Private Sub txt处方限量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt当前售价_GotFocus()
    Me.txt当前售价.SelStart = 0: Me.txt当前售价.SelLength = 100
End Sub

Private Sub txt当前售价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt当前售价_LostFocus()
Dim dbl成本价 As Double
    Dim dbl指导售价 As Double
    Dim dbl加成率 As Double
    Dim dbl差价率 As Double
    Dim dbl现售价 As Double
    
    Me.txt当前售价.Text = FormatEx(Val(txt当前售价), mintPriceDigit)
    
    dbl现售价 = Val(Me.txt当前售价.Text)
    dbl成本价 = Val(Me.txt成本价格.Text)
    dbl指导售价 = Val(Me.txt指导售价.Text)
    
    '满足这些条件才计算加成率
    If dbl成本价 > 0 And dbl指导售价 > 0 And dbl现售价 > 0 And dbl现售价 <= dbl指导售价 Then
        dbl加成率 = dbl现售价 / dbl成本价 - 1
        
        If dbl加成率 < 0 Then Exit Sub
        
        dbl加成率 = dbl加成率 * 100
        
        Me.txt加成率.Text = Format(dbl加成率, "0.00")
        
        '通过加成率计算指导差价率
        dbl差价率 = dbl加成率
        Call Calc(dbl差价率, False)
        Me.txt指导差率.Text = Format(dbl差价率, "0.00000")
    End If
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt管理费比例_GotFocus()
    txt管理费比例.SelStart = 0: txt管理费比例.SelLength = 100
End Sub

Private Sub txt管理费比例_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txt管理费比例_Validate(Cancel As Boolean)
    txt管理费比例.Text = Format(Val(txt管理费比例.Text), "#0.00")
End Sub

Private Sub txt规格_GotFocus()
    Me.txt规格.SelStart = 0: Me.txt规格.SelLength = 100
End Sub

Private Sub txt规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt合同单位_GotFocus()
    Me.txt合同单位.SelStart = 0: Me.txt合同单位.SelLength = Len(Me.txt合同单位.Text)
End Sub

Private Sub txt合同单位_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
        
    strTmp = UCase(Trim(Me.txt合同单位.Text))
    
    If strTmp = "" Then
        Me.txt合同单位.Tag = "|"
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    ElseIf strTmp = Split(Me.txt合同单位.Tag, "|")(1) Then
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    End If
       
    gstrSql = "Select 编码,名称,简码,id" & _
            " From 供应商" & _
            " where (编码 Like [1] " & _
            "       Or 名称 Like [2] " & _
            "       Or 简码 Like [2])" & _
            " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    
    With rsTemp
        If .EOF Then
            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
            Me.txt合同单位.Text = Split(Me.txt合同单位.Tag, "|")(1)
            Me.txt合同单位.SelStart = 0: Me.txt合同单位.SelLength = Len(Me.txt合同单位.Text)
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            Me.txt合同单位.Text = Trim(rsTemp!名称): Me.txt合同单位.Tag = rsTemp!ID & "|" & rsTemp!名称
            Call zlcommfun.PressKey(vbKeyTab): Exit Sub
        Else
            With Me.msf合同单位
                .Left = Me.stbSpec.Left + Me.txt合同单位.Left
                .Top = Me.stbSpec.Top + Me.txt合同单位.Top - Me.msf合同单位.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 1500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
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
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt加成率_GotFocus()
    Call zlControl.TxtSelAll(txt加成率)
End Sub

Private Sub txt加成率_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub txt加成率_LostFocus()
    Dim cur价格 As Double
    '重新计算指导差价率和加成率
    cur价格 = Val(txt加成率.Text)
    Call Calc(cur价格, False)
    Me.txt加成率.Text = Format(txt加成率.Text, "0.00")
    Me.txt指导差率.Text = Format(cur价格, "0.00000")
End Sub

Private Sub txt结算价_GotFocus()
    Me.txt结算价.SelStart = 0: Me.txt结算价.SelLength = 100
End Sub

Private Sub txt结算价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt结算价_LostFocus()
    Me.txt结算价.Text = FormatEx(Val(txt结算价), mintPriceDigit)
End Sub

Private Sub txt扣率_Change()
    Me.txt结算价.Text = FormatEx(Val(Me.txt指导批价.Text) * Val(Me.txt扣率.Text) / 100, mintCostDigit)
End Sub

Private Sub txt扣率_GotFocus()
    Me.txt扣率.SelStart = 0: Me.txt扣率.SelLength = 100
End Sub

Private Sub txt扣率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt扣率_LostFocus()
    Me.txt扣率.Text = FormatEx(Val(txt扣率), mintPriceDigit)
End Sub

Private Sub txt名称_Change()
    Dim strTmp As String
    '重新检查名称，并去 掉特殊字符
    strTmp = MoveSpecialChar(txt名称.Text)
    If txt名称.Text <> strTmp Then
        txt名称.Text = strTmp
    End If
    Me.txt拼音.Text = zlGetSymbol(strTmp, 0, mint简码长度)
    Me.txt五笔.Text = zlGetSymbol(strTmp, 1, mint简码长度)
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("？")
        Case Asc("%")
            KeyAscii = Asc("％")
        Case Asc("_")
            KeyAscii = Asc("＿")
    End Select
    If KeyAscii = vbKeyReturn Then
        Call zlcommfun.PressKey(vbKeyTab)
    Else
        If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Me.txt拼音.Text = zlGetSymbol(Me.txt名称.Text, 0, mint简码长度)
        Me.txt五笔.Text = zlGetSymbol(Me.txt名称.Text, 1, mint简码长度)
    End If
End Sub

Private Sub txt名称_LostFocus()
    Me.txt拼音.Text = zlGetSymbol(Me.txt名称.Text, 0, mint简码长度)
    Me.txt五笔.Text = zlGetSymbol(Me.txt名称.Text, 1, mint简码长度)
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt拼音_GotFocus()
    Me.txt拼音.SelStart = 0: Me.txt拼音.SelLength = 100
End Sub

Private Sub txt拼音_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt申领阀值_GotFocus()
    txt申领阀值.SelStart = 0: txt申领阀值.SelLength = Len(txt申领阀值)
End Sub

Private Sub txt申领阀值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt售价单位_Change()
    Me.lbl剂量系数.Caption = "(1" & Me.txt售价单位.Text & "="
    If glngSys \ 100 = 8 Then
        Me.txt药房单位 = Me.txt售价单位
    End If
    Me.lbl药房单位Child.Caption = Me.txt售价单位 & ")"
    Me.lbl药库单位Child.Caption = Me.txt售价单位 & ")"
    Me.lbl申领单位Child.Caption = Me.txt售价单位 & ")"
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
    Call zlcommfun.OpenIme(True)
End Sub


Private Sub txt售价单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt售价单位_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub


Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt五笔_GotFocus()
    Me.txt五笔.SelStart = 0: Me.txt五笔.SelLength = 100
End Sub

Private Sub txt五笔_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt药房包装_GotFocus()
    Me.txt药房包装.SelStart = 0: Me.txt药房包装.SelLength = 100
End Sub


Private Sub txt药房包装_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt药房单位_Change()
    Me.lbl药房包装.Caption = "(1" & Me.txt药房单位.Text & "="
    Call cbo申领单位_Click
End Sub

Private Sub txt药房单位_GotFocus()
    Me.txt药房单位.SelStart = 0: Me.txt药房单位.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub


Private Sub txt药房单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt药房单位_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub


Private Sub txt药库包装_GotFocus()
    Me.txt药库包装.SelStart = 0: Me.txt药库包装.SelLength = 100
End Sub

Private Sub txt药库包装_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt药库单位_Change()
    Me.lbl药库包装.Caption = "(1" & Me.txt药库单位.Text & "="
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
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt药库单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt药库单位_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt增值税率_GotFocus()
    Call zlControl.TxtSelAll(txt增值税率)
End Sub

Private Sub txt增值税率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt增值税率_LostFocus()
    txt增值税率.Text = Format(txt增值税率.Text, "0.00")
End Sub


Private Sub txt指导差率_GotFocus()
    Me.txt指导差率.SelStart = 0: Me.txt指导差率.SelLength = 100
End Sub

Private Sub txt指导差率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt指导差率_LostFocus()
    Dim cur价格 As Double
    '重新计算指导差价率和加成率
    cur价格 = Val(txt指导差率.Text)
    If cur价格 < 100 Then
        Call Calc(cur价格, True)
        Me.txt指导差率.Text = Format(txt指导差率.Text, "0.00000")
        Me.txt加成率.Text = Format(cur价格, "0.00")
    Else
        '不允许出现指导差价率大于等于100的情况，因此需要从加成率反算回来
        cur价格 = Val(txt加成率.Text)
        Call Calc(cur价格, False)
        Me.txt指导差率.Text = Format(cur价格, "0.00000")
    End If
End Sub

Private Sub txt指导批价_Change()
    Me.txt结算价.Text = FormatEx(Val(Me.txt指导批价.Text) * Val(Me.txt扣率.Text) / 100, mintCostDigit)
End Sub

Private Sub txt指导批价_GotFocus()
    Me.txt指导批价.SelStart = 0: Me.txt指导批价.SelLength = 100
End Sub

Private Sub txt指导批价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt指导批价_LostFocus()
    Me.txt指导批价.Text = FormatEx(Val(txt指导批价), mintCostDigit)
End Sub

Private Sub txt指导售价_GotFocus()
    Me.txt指导售价.SelStart = 0: Me.txt指导售价.SelLength = 100
End Sub

Private Sub txt指导售价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt指导售价_LostFocus()
    Me.txt指导售价.Text = FormatEx(Val(txt指导售价), mintPriceDigit)
End Sub

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
