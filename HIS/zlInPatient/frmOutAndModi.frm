VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOutAndModi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "出院及调整出院"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10800
   Icon            =   "frmOutAndModi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10800
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd接收 
      Caption         =   "接收确认(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5920
      TabIndex        =   160
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chk接收 
      Caption         =   "病历已接收"
      Height          =   255
      Left            =   4680
      TabIndex        =   159
      Top             =   195
      Width           =   1215
   End
   Begin VB.CommandButton cmd修改 
      Caption         =   "修  改(&M)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8360
      TabIndex        =   4
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmd出院 
      Caption         =   "出  院(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.TextBox txt状态 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txt住院次数 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txt住院号 
      Height          =   375
      Left            =   720
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退  出(&X)"
      Height          =   350
      Left            =   9480
      TabIndex        =   5
      Top             =   120
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      TabHeight       =   882
      TabMaxWidth     =   3528
      MouseIcon       =   "frmOutAndModi.frx":0442
      TabCaption(0)   =   "基本信息(&1)"
      TabPicture(0)   =   "frmOutAndModi.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl其他证件"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label36"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl登记时间"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label27"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label33"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label32"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl单位帐号"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl单位开户行"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl单位邮编"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl单位电话"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl工作单位"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl联系人电话"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl联系人地址"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl联系人关系"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl联系人姓名"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl家庭地址邮编"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl家庭电话"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl家庭地址"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lvl婚姻状况"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl学历"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl国籍"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl民族"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl职业"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl身份"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl身份证号"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl出生地点"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl出生日期"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl医疗付款"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label13"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl门诊号"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl姓名"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl性别"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl年龄"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label9"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label8"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label7"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl区域"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lbl籍贯"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lbl户口地址邮编"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lbl户口地址"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "lbl联系人身份证号"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt(65)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt(6)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt(16)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt(12)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt(18)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt(24)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt(26)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txt(29)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt(30)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt(60)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt(63)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt(64)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txt(41)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt(59)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt(33)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt(34)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt(31)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt(32)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt(17)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt(21)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt(28)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txt(27)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt(25)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txt(22)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txt(20)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txt(19)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txt(11)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txt(10)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txt(13)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txt(14)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txt(15)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txt(23)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txt(7)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txt(4)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txt(5)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txt(2)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txt(0)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txt(1)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txt(3)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txt(8)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txt(9)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "cbo(1)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "chk担保"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txt(74)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txt(73)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "txt(72)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "txt(75)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).ControlCount=   92
      TabCaption(1)   =   "住院信息(&2)"
      TabPicture(1)   =   "frmOutAndModi.frx":0D38
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label39"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label38"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label37"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label35"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label34"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label31"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label30"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label28"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label26"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label16"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label24"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label23"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label22"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label21"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label29"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label25"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label20"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label19"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label18"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label17"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label14"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label15"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lbl费别"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label4"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label3"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label6"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label40"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label41"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label42"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txt(68)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txt(67)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txt(66)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txt(53)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txt(46)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txt(52)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txt(54)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txt(50)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txt(51)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txt(38)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "txt(62)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "txt(61)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "txt(57)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "cbo(0)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "txt(39)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "txt(56)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txt(47)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "txt(55)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "txt(48)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "txt(49)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "txt(42)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "txt(43)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "txt(44)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "txt(35)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "txt(58)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "txt(40)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "txt(36)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "txt(37)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "txt(45)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "txt(69)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "txt(70)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "txt(71)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).ControlCount=   63
      TabCaption(2)   =   "合并记录(&3)"
      TabPicture(2)   =   "frmOutAndModi.frx":1052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "msfMerge"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   75
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   68
         Top             =   5250
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   72
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   163
         Top             =   3840
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   73
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   162
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   74
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   161
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   71
         Left            =   -70935
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   70
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   69
         Left            =   -67160
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   4800
         Width           =   2355
      End
      Begin VB.CheckBox chk担保 
         BackColor       =   &H8000000A&
         Caption         =   "临时"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         MaskColor       =   &H00000000&
         TabIndex        =   79
         Top             =   5670
         Width           =   735
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   4575
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   1009
         Width           =   580
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1358
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1009
         Width           =   680
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1009
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1009
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1358
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   23
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   5250
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   15
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2415
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   14
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2056
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   13
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2056
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   3120
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   3885
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   60
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   22
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   4545
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   25
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   58
         Top             =   5250
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   27
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         Top             =   3105
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   28
         Left            =   8970
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   56
         Top             =   3105
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   21
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   55
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   17
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   54
         Top             =   2056
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   32
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   5700
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   31
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   5700
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   34
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   51
         Top             =   5700
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   33
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   50
         Top             =   5700
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   45
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2415
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   37
         Left            =   -70935
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   36
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   40
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   58
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   35
         Left            =   -68475
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   44
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   43
         Left            =   -70935
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   42
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   49
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   48
         Left            =   -70935
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   55
         Left            =   -70515
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   47
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   56
         Left            =   -68310
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   39
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1605
         Width           =   4065
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   -66720
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   59
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4530
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   57
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   61
         Left            =   -70935
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   41
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4180
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   64
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1358
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   60
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4890
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   30
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   3825
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   29
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3480
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2760
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   24
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4905
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2760
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   16
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2415
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   62
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   38
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   51
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3630
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   50
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   54
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5160
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   52
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4020
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   46
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2415
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   53
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4410
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   65
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2415
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   66
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   67
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4780
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   68
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5520
         Width           =   2355
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfMerge 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   80
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483635
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         MouseIcon       =   "frmOutAndModi.frx":13EC
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl联系人身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人身份证"
         Height          =   180
         Left            =   5190
         TabIndex        =   167
         Top             =   5310
         Width           =   1080
      End
      Begin VB.Label lbl户口地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   180
         Left            =   420
         TabIndex        =   166
         Top             =   3900
         Width           =   720
      End
      Begin VB.Label lbl户口地址邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址邮编"
         Height          =   180
         Left            =   60
         TabIndex        =   165
         Top             =   4260
         Width           =   1080
      End
      Begin VB.Label lbl籍贯 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "籍贯"
         Height          =   180
         Left            =   3480
         TabIndex        =   164
         Top             =   4260
         Width           =   360
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   -71370
         TabIndex        =   158
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         Height          =   180
         Left            =   -74340
         TabIndex        =   157
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院方式"
         Height          =   180
         Left            =   -68040
         TabIndex        =   151
         Top             =   4840
         Width           =   720
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   180
         Left            =   780
         TabIndex        =   149
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保险类别"
         Height          =   180
         Left            =   5535
         TabIndex        =   148
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   600
         TabIndex        =   147
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊费别"
         Height          =   180
         Left            =   5535
         TabIndex        =   146
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3480
         TabIndex        =   145
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   780
         TabIndex        =   144
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   3480
         TabIndex        =   143
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   5715
         TabIndex        =   142
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   600
         TabIndex        =   141
         Top             =   1425
         Width           =   540
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付费方式"
         Height          =   180
         Left            =   8205
         TabIndex        =   140
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   3120
         TabIndex        =   139
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   420
         TabIndex        =   138
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   5535
         TabIndex        =   137
         Top             =   2116
         Width           =   720
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   780
         TabIndex        =   136
         Top             =   2475
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   3480
         TabIndex        =   135
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   5895
         TabIndex        =   134
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   3480
         TabIndex        =   133
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   8565
         TabIndex        =   132
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   420
         TabIndex        =   131
         Top             =   2115
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   180
         Left            =   600
         TabIndex        =   130
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   420
         TabIndex        =   129
         Top             =   3540
         Width           =   720
      End
      Begin VB.Label lbl家庭地址邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址邮编"
         Height          =   180
         Left            =   2760
         TabIndex        =   128
         Top             =   3540
         Width           =   1080
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   240
         TabIndex        =   127
         Top             =   4605
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   2940
         TabIndex        =   126
         Top             =   5310
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   240
         TabIndex        =   125
         Top             =   4965
         Width           =   900
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   240
         TabIndex        =   124
         Top             =   5310
         Width           =   900
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   5535
         TabIndex        =   123
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   5535
         TabIndex        =   122
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   8205
         TabIndex        =   121
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   5355
         TabIndex        =   120
         Top             =   3540
         Width           =   900
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   5535
         TabIndex        =   119
         Top             =   3885
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   3120
         TabIndex        =   118
         Top             =   5760
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   420
         TabIndex        =   117
         Top             =   5760
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Index           =   1
         Left            =   5715
         TabIndex        =   116
         Top             =   5760
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   -66870
         TabIndex        =   115
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   -69285
         TabIndex        =   114
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   -74340
         TabIndex        =   113
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   -68925
         TabIndex        =   112
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         Height          =   180
         Left            =   -69240
         TabIndex        =   111
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院天数"
         Height          =   180
         Left            =   -66885
         TabIndex        =   110
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转科信息"
         Height          =   180
         Left            =   -74520
         TabIndex        =   109
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院病况"
         Height          =   180
         Left            =   -69285
         TabIndex        =   108
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院科室"
         Height          =   180
         Left            =   -71730
         TabIndex        =   107
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   -69105
         TabIndex        =   106
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   -71730
         TabIndex        =   105
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   -74520
         TabIndex        =   104
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   -71280
         TabIndex        =   103
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   -74520
         TabIndex        =   102
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院科室"
         Height          =   180
         Left            =   -69120
         TabIndex        =   101
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   -71370
         TabIndex        =   100
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院目的"
         Height          =   180
         Left            =   -74520
         TabIndex        =   99
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院中医诊断"
         Height          =   180
         Left            =   -74880
         TabIndex        =   98
         Top             =   5220
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   -74520
         TabIndex        =   97
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院中医诊断"
         Height          =   180
         Left            =   -74880
         TabIndex        =   96
         Top             =   3690
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   -68925
         TabIndex        =   95
         Top             =   1665
         Width           =   360
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊诊断"
         Height          =   180
         Left            =   5535
         TabIndex        =   94
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊中医诊断"
         Height          =   180
         Left            =   5175
         TabIndex        =   93
         Top             =   4950
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主任(副主任)医师"
         Height          =   180
         Left            =   -72450
         TabIndex        =   92
         Top             =   2460
         Width           =   1440
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   -74520
         TabIndex        =   91
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊医师"
         Height          =   180
         Left            =   5535
         TabIndex        =   90
         Top             =   4240
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病况"
         Height          =   180
         Left            =   -66840
         TabIndex        =   89
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   7665
         TabIndex        =   88
         Top             =   5760
         Width           =   540
      End
      Begin VB.Label lbl登记时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
         Height          =   180
         Left            =   8205
         TabIndex        =   87
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡号"
         Height          =   180
         Left            =   8205
         TabIndex        =   86
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   -74520
         TabIndex        =   85
         Top             =   4470
         Width           =   720
      End
      Begin VB.Label lbl其他证件 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他证件"
         Height          =   180
         Left            =   5520
         TabIndex        =   84
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型"
         Height          =   180
         Left            =   -66870
         TabIndex        =   83
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   -74520
         TabIndex        =   82
         Top             =   4840
         Width           =   720
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医出院情况"
         Height          =   180
         Left            =   -74880
         TabIndex        =   81
         Top             =   5560
         Width           =   1080
      End
   End
   Begin VB.Label lbl状态 
      BackStyle       =   0  'Transparent
      Caption         =   "状态"
      Height          =   255
      Left            =   3360
      TabIndex        =   154
      Top             =   195
      Width           =   495
   End
   Begin VB.Label lbl住院次数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "次数"
      Height          =   180
      Left            =   2160
      TabIndex        =   153
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl住院号 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Height          =   180
      Left            =   120
      TabIndex        =   152
      Top             =   195
      Width           =   540
   End
End
Attribute VB_Name = "frmOutAndModi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long '要查看的病人ID
Private mlng主页ID As Long '住院病人时传入主页ID
Private mstr状态 As String
Private mstr住院号 As String
Private mstrPrivs As String
Private mlngModul As Long
Private mint病历接收 As Integer

Private Enum txtName
    '要求和SQL的字段对应
    病人ID = 0
    姓名 = 1
    性别 = 2
    年龄 = 3
    门诊号 = 4
    费别 = 5
    医疗付款方式 = 6
    医保号 = 7
    险类 = 8
    区域 = 9
    国籍 = 10
    民族 = 11
    学历 = 12
    婚姻状况 = 13
    职业 = 14
    身份 = 15
    出生日期 = 16
    身份证号 = 17
    出生地点 = 18
    家庭地址 = 19
    家庭地址邮编 = 20
    家庭电话 = 21
    联系人姓名 = 22
    联系人关系 = 23
    联系人地址 = 24
    联系人电话 = 25
    工作单位 = 26
    单位电话 = 27
    单位邮编 = 28
    单位开户行 = 29
    单位帐号 = 30
    预交余额 = 31
    费用余额 = 32
    担保人 = 33
    担保额 = 34
    
    住院次数 = 35
    住院号 = 36
    出院病床 = 37
    备注 = 38
    住院目的 = 39
    住院费别 = 40
    门诊医师 = 41
    住院医师 = 42
    主治医师 = 57
    主任医师 = 61
    责任护士 = 43
    登记人 = 44
    床位等级 = 45
    护理等级 = 46
    入院日期 = 47
    入院科室 = 48
    入院病况 = 49
    当前病况 = 62
    转科信息 = 52
    出院日期 = 55
    出院科室 = 56
    住院天数 = 58
    
    门诊诊断 = 59
    门诊中医诊断 = 60
    入院诊断 = 50
    入院中医诊断 = 51
    出院诊断 = 53
    出院中医诊断 = 54
    出院方式 = 69
    
    出院情况 = 67
    中医出院情况 = 68
    '问题51167,刘鹏飞,2012-07-09,增加"联系人身份证号"
    联系人身份证号 = 75
End Enum

Private Enum cboName
    主页ID = 0
    年龄单位 = 1
End Enum

Private Function ReadCard(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln查看某次住院 As Boolean) As Boolean
'功能：读取指定病人信息,并显示在界面上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTxt As String, strTmp As String, strHead As String
    Dim i As Integer, j As Integer, arrTxt As Variant
    Dim str出院情况 As String
    Dim str其它诊断 As String
    Dim blnPassShowCard  As Boolean, strCard As String
    
    On Error GoTo errH
    
    '51572:刘鹏飞,2013-11-04,就诊卡是否密文显示
    strSQL = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        blnPassShowCard = Nvl(rsTmp!卡号密文) <> ""
    End If
    
    If blnPassShowCard = True Then
        strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡号,"
    Else
        strCard = "A.就诊卡号 as 就诊卡号,"
    End If
        
    '问题51167,刘鹏飞,2012-07-09,增加"联系人身份证号"
    strSQL = "Select a.病人id, NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄, a.门诊号, a.费别, a.医疗付款方式, a.险类, a.籍贯, a.区域, a.国籍, a.民族, a.学历," & vbNewLine & _
            "            a.婚姻状况, a.职业, a.身份, Decode(To_Date(To_Char(出生日期, 'YYYY-MM-DD HH24:MI'), 'YYYY-MM-DD HH24:MI') - Trunc(出生日期), 0, To_Char(出生日期, 'YYYY-MM-DD'),To_char(出生日期,'YYYY-MM-DD HH24:MI')) 出生日期, " & _
            "            a.身份证号, a.出生地点, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.户口地址, a.户口地址邮编, a.联系人姓名," & vbNewLine & _
            "            a.联系人关系, a.联系人地址, a.联系人电话,a.联系人身份证号, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人," & vbNewLine & _
            "            a.担保额, a.担保性质, a.住院次数,a.主页ID 就诊次数, b.住院号, To_char(a.登记时间,'yyyy-mm-dd hh24:mi:ss') As 登记时间," & strCard & "b.出院病床, b.备注, b.住院目的, b.门诊医师, b.住院医师," & vbNewLine & _
            "            b.责任护士, b.登记人, b.入院病况, b.当前病况, b.住院天数, b.费别 As 住院费别, c.预交余额, c.费用余额," & vbNewLine & _
            "            Nvl(A.医保号,d.信息值) 医保号, e.名称 As 护理等级, g.名称 As 床位等级, m.名称 As 入院科室, n.名称 As 出院科室," & vbNewLine & _
            "            To_char(b.入院日期,'yyyy-mm-dd hh24:mi:ss') 入院日期, To_char(b.出院日期,'yyyy-mm-dd hh24:mi:ss') 出院日期,A.其他证件,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型,B.出院方式 " & vbNewLine & _
            "From 病人信息 a, 病案主页 b, (select 病人ID,性质,Nvl(sum(预交余额),0) 预交余额,Nvl(sum(费用余额),0) 费用余额 from 病人余额 where 病人ID=[1] and 类型=2 group by 病人ID,性质) c, 病案主页从表 d, 收费项目目录 e, 床位状况记录 f, 收费项目目录 g, 部门表 m, 部门表 n" & vbNewLine & _
            "Where a.病人id = b.病人id(+) And " & IIf(lng主页ID = 0, "Nvl(a.主页ID,0)", "[2]") & "=b.主页ID(+) And a.病人ID=[1] And a.病人id = c.病人id(+) And" & vbNewLine & _
            "           c.性质(+) = 1 And b.病人id = d.病人id(+) And" & vbNewLine & _
            "           b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.入院科室id = m.Id(+) And b.出院科室id = n.Id(+) And" & vbNewLine & _
            "           b.护理等级id = e.ID(+) And b.当前病区id = f.病区id(+) And b.出院病床 = f.床号(+) And f.等级id = g.ID(+)"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTmp.EOF Then Exit Function
        
'    If bln查看某次住院 Then
'       strTxt = "出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
'                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53,出院中医诊断=54,病人类型=66"
'    Else
        strTxt = "病人ID=0,姓名=1,性别=2,年龄=3,门诊号=4,费别=5,医疗付款方式=6,医保号=7,险类=8,区域=9,国籍=10,民族=11,学历=12,婚姻状况=13,职业=14," & _
                " 身份=15,出生日期=16,身份证号=17,出生地点=18,家庭地址=19,家庭地址邮编=20,家庭电话=21,联系人姓名=22,联系人关系=23,联系人地址=24,联系人电话=25," & _
                " 工作单位=26,单位电话=27,单位邮编=28,单位开户行=29,单位帐号=30,预交余额=31,费用余额=32,担保人=33,担保额=34,住院次数=35,住院号=36," & _
                " 出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53," & _
                " 出院中医诊断=54,登记时间=63,就诊卡号=64,其他证件=65,病人类型=66,出院方式=69,户口地址=72,户口地址邮编=73,籍贯=74,联系人身份证号=75"
'    End If
    
    arrTxt = Split(strTxt, ",")
    
    For i = 0 To UBound(arrTxt)
        strTmp = Trim(arrTxt(i))
        
        If strTmp <> "" Then
            '排开暂不处理的字段
            If InStr(1, ",门诊诊断,门诊中医诊断,入院诊断,入院中医诊断,出院诊断,出院中医诊断,转科信息,", "," & Trim(Split(strTmp, "=")(0)) & ",") = 0 Then
                If InStr(1, ",费用余额,预交余额,", "," & Trim(Split(strTmp, "=")(0)) & ",") > 0 Then
                    txt(Trim(Split(strTmp, "=")(1))).Text = Format(Val("" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))), "0.00")
                Else
                    txt(Trim(Split(strTmp, "=")(1))).Text = "" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))
                End If
            End If
        End If
    Next
    
    txt(70).Text = txt(0).Text
    txt(71).Text = txt(1).Text
    
    '其它专门处理
    '----------------------------------------------
    Call LoadOldData("" & rsTmp!年龄, txt(txtName.年龄), cbo(cboName.年龄单位))
    If cbo(cboName.年龄单位).ListIndex = -1 Then txt(txtName.年龄).width = txt(txtName.年龄).width + cbo(cboName.年龄单位).width
    chk担保.Value = Val("" & rsTmp!担保性质)
    
    
    '住院信息
    '----------------------------------------------
'    If Not bln查看某次住院 Then lng主页Id = Val(Nvl(rsTmp!就诊次数, 0))
    '住院病人的诊断情况
    If lng主页ID > 0 Then
        strSQL = "Select 诊断类型,疾病ID,诊断描述,出院情况 From 病人诊断记录 Where 诊断次序=1 And 记录来源=2 And 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!诊断类型
                    Case 1
                        j = txtName.门诊诊断
                    Case 11
                        j = txtName.门诊中医诊断
                    Case 2
                        j = txtName.入院诊断
                    Case 12
                        j = txtName.入院中医诊断
                    Case 3
                        j = txtName.出院诊断
                        '问题27832 by lesfeng 2010-01-18 增加“出院情况
                        str出院情况 = IIf(IsNull(rsTmp!出院情况), "", rsTmp!出院情况)
                        txt(txtName.出院情况).Text = str出院情况
                    Case 13
                        j = txtName.出院中医诊断
                        '问题27832 by lesfeng 2010-01-18 增加“出院情况
                        str出院情况 = IIf(IsNull(rsTmp!出院情况), "", rsTmp!出院情况)
                        txt(txtName.中医出院情况).Text = str出院情况
                    Case Else
                        j = 0
                End Select
                If j <> 0 Then txt(j).Text = IIf(IsNull(rsTmp!疾病ID), "", "(" & rsTmp!疾病ID & ")") & rsTmp!诊断描述
                
                rsTmp.MoveNext
            Next
        Else
            txt(txtName.门诊诊断).Text = ""
            txt(txtName.门诊中医诊断).Text = ""
            txt(txtName.入院诊断).Text = ""
            txt(txtName.入院中医诊断).Text = ""
            txt(txtName.出院诊断).Text = ""
            txt(txtName.出院中医诊断).Text = ""
            '问题27832 by lesfeng 2010-01-18 增加“出院情况
            txt(txtName.出院情况).Text = ""
            txt(txtName.中医出院情况).Text = ""
        End If
        
        '问题28139 by lesfeng 2010-02-04
        strSQL = "Select 诊断类型,疾病ID,诊断描述,出院情况 From 病人诊断记录 Where 诊断类型 in (3,13) and 诊断次序>1 And 记录来源=2 And 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!诊断类型
                    Case 3
                        j = txtName.出院诊断
                    Case 13
                        j = txtName.出院中医诊断
                    Case Else
                        j = 0
                End Select
                                
                If j <> 0 Then
                    strTmp = txt(j).Text
                    str其它诊断 = IIf(IsNull(rsTmp!疾病ID), "", "(" & rsTmp!疾病ID & ")") & rsTmp!诊断描述 & IIf(IsNull(rsTmp!出院情况), "", "(" & rsTmp!出院情况 & ")")
                    If InStr(1, strTmp, "其它诊断:") > 0 Then
                        txt(j).Text = strTmp & "," & str其它诊断
                    Else
                        txt(j).Text = strTmp & ";其它诊断:" & str其它诊断
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        '转科信息
        txt(txtName.转科信息).Text = ""
        strSQL = _
            " Select Distinct 1 as 开始原因,To_Date('1900-01-01','YYYY-MM-DD') as 开始时间,B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.科室ID=B.ID And A.开始时间 is Not NULL And A.开始原因 IN(1,2)" & _
            " And A.病人ID=[1] And 主页ID=[2]" & _
            " Union ALL " & _
            " Select A.开始原因,A.开始时间,B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.科室ID=B.ID And A.开始时间 is Not NULL And A.开始原因=3" & _
            " And A.病人ID=[1] And 主页ID=[2]" & _
            " Order by 开始时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        rsTmp.Filter = "开始原因=3"
        If Not rsTmp.EOF Then
            rsTmp.Filter = 0
            Do While Not rsTmp.EOF
                txt(txtName.转科信息).Text = txt(txtName.转科信息).Text & " ─→ " & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txt(txtName.转科信息).Text = Mid(txt(txtName.转科信息).Text, 5)
        End If
        
        '病案主页从表
        txt(txtName.主治医师).Text = ""
        txt(txtName.主任医师).Text = ""
        strSQL = " Select 信息名,信息值 From 病案主页从表 Where (信息名='主治医师' Or 信息名='主任医师') And 病人ID=[1] And 主页ID=[2]"
        rsTmp.Filter = ""
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        rsTmp.Filter = "信息名='主治医师'"
        If Not rsTmp.EOF Then txt(txtName.主治医师).Text = "" & rsTmp!信息值
        rsTmp.Filter = "信息名='主任医师'"
        If Not rsTmp.EOF Then txt(txtName.主任医师).Text = "" & rsTmp!信息值
    End If
    
    
    '3.病人合并信息
    If Not bln查看某次住院 Then
        
        strSQL = "Select 原信息,合并原因,操作员姓名,合并时间 From 病人合并记录 Where 病人ID=[1] Order by 合并时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
                
        strHead = "合并时间,1,1800|操作员,4,800|合并原因,1,1800|" & _
                "病人ID,1,800|门诊号,1,900|住院号,1,800|就诊卡号,1,900|姓名,4,800|" & _
                "性别,4,500|年龄,4,800|出生日期,1,1000|身份证号,1,1800|婚姻状况,4,900|职业,1,1000|家庭地址,1,4200"
        With msfMerge
            .Redraw = False
            .Rows = rsTmp.RecordCount + 1
            
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
                .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
                .ColAlignmentFixed(i) = 4
            Next
            
            Call RestoreFlexState(msfMerge, App.ProductName & "\" & Me.Name)
            .RowHeight(0) = 320
            .Col = 0: .ColSel = .Cols - 1
            
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!合并时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 1) = "" & rsTmp!操作员姓名
                .TextMatrix(i, 2) = "" & rsTmp!合并原因
                
                'v_原信息:=r_InfoA.病人Id || ',' || r_InfoA.门诊号 || ',' ||  r_InfoA.住院号 || ',' ||  r_InfoA.就诊卡号 || ',' ||  r_InfoA.姓名 ||  ',' ||  r_InfoA.性别 ||  ',' ||
                '   r_InfoA.年龄 ||  ',' || to_char(r_InfoA.出生日期,'yyyy-mm-dd') ||  ',' || r_InfoA.身份证号 ||  ',' || r_InfoA.婚姻状况 ||  ',' || r_InfoA.职业 ||  ',' || r_InfoA.家庭地址;
                arrTxt = Split(rsTmp!原信息, ",")
                If UBound(arrTxt) >= 11 Then
                    .TextMatrix(i, 3) = arrTxt(0)
                    .TextMatrix(i, 4) = arrTxt(1)
                    .TextMatrix(i, 5) = arrTxt(2)
                    If blnPassShowCard = True Then
                        .TextMatrix(i, 6) = String(Len(Trim(arrTxt(3))), "*")
                    Else
                        .TextMatrix(i, 6) = arrTxt(3)
                    End If
                    .TextMatrix(i, 7) = arrTxt(4)
                    .TextMatrix(i, 8) = arrTxt(5)
                    .TextMatrix(i, 9) = arrTxt(6)
                    .TextMatrix(i, 10) = arrTxt(7)
                    .TextMatrix(i, 11) = arrTxt(8)
                    .TextMatrix(i, 12) = arrTxt(9)
                    .TextMatrix(i, 13) = arrTxt(10)
                    .TextMatrix(i, 14) = arrTxt(11)
                End If
                rsTmp.MoveNext
            Next
            
            
            If rsTmp.RecordCount = 0 Then .Rows = 2: .Row = 1: .FixedRows = 1
            .Redraw = True
        End With
        
    End If
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo_Click(Index As Integer)
    If Index = cboName.主页ID Then      '启动加载住院次数时不调用
        If cbo(cboName.主页ID).Visible Then Call ReadCard(mlng病人ID, cbo(cboName.主页ID).ItemData(cbo(cboName.主页ID).ListIndex), True)
        If Trim(txt(55).Text) = "" Then txt状态.Text = "在院"
        If IsDate(Trim(txt(55).Text)) Then txt状态.Text = "出院"
        
        mstr状态 = Trim(txt状态.Text)
        mlng主页ID = cbo(cboName.主页ID).ItemData(cbo(cboName.主页ID).ListIndex)
        mstr住院号 = txt(36).Text
        txt住院号 = mstr住院号
        txt住院次数 = mlng主页ID
        
        If Trim(txt状态.Text) = "在院" Then
            cmd出院.Enabled = True
            cmd修改.Enabled = False
        Else
            cmd出院.Enabled = False
            cmd修改.Enabled = True
        End If
    End If
End Sub

Private Sub chk接收_Click()
    If mint病历接收 <> Val(chk接收.Value) Then
        cmd接收.Enabled = True
    Else
        cmd接收.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmd出院_Click()
    Dim lng病人ID As Long, lng主页ID As Long

    lng病人ID = mlng病人ID
    lng主页ID = mlng主页ID

    Call ExecPatiChange(EFun.E出院, Me, mstrPrivs, lng病人ID, lng主页ID)
    If gblnOK Then
        cmd出院.Enabled = False
        mlng病人ID = 0
        mlng主页ID = 0
        txt住院号 = ""
        txt住院次数 = ""
        txt状态 = ""
        txt住院号.SetFocus
    End If
End Sub

Private Sub cmd接收_Click()
    Dim int病历接收 As Integer
    Dim strSQL As String
    
    If InStr(mstrPrivs, "调整出院时间") <> 0 Then
    Else
        MsgBox "你没有‘调整出院时间’权限，不能更新病人病历接收信息！", vbInformation, gstrSysName
        Exit Sub
    End If
    int病历接收 = chk接收.Value
    If mlng病人ID = 0 Then Exit Sub
    If mlng主页ID = 0 Then Exit Sub
    
    On Error GoTo errH
    strSQL = "Zl_病案主页_病历接收(" & mlng病人ID & "," & mlng主页ID & "," & int病历接收 & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mint病历接收 = int病历接收
    cmd接收.Enabled = False
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd修改_Click()
    Dim lng病人ID As Long, lng主页ID As Long

    lng病人ID = mlng病人ID
    lng主页ID = mlng主页ID

    Call ExecPatiChange(EFun.E修改出院时间, Me, mstrPrivs, lng病人ID, lng主页ID)
    If gblnOK Then
        cmd修改.Enabled = False
        mlng病人ID = 0
        mlng主页ID = 0
        txt住院号 = ""
        txt住院次数 = ""
        txt状态 = ""
        txt住院号.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    '固定信息处理
    With cbo(cboName.年龄单位)
        .AddItem "岁"
        .AddItem "月"
        .AddItem "天"
        .ListIndex = 0
    End With
    mlng病人ID = 0
    mlng主页ID = 0
    mint病历接收 = 0
    chk接收.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng病人ID = 0
    mlng主页ID = 0
    Call SaveFlexState(msfMerge, App.ProductName & "\" & Me.Name)
End Sub

Private Sub txt住院次数_GotFocus()
    zlControl.TxtSelAll txt住院次数
    zlCommFun.OpenIme False
End Sub

Private Sub txt住院次数_KeyPress(KeyAscii As Integer)
    Dim lng次数 As Long
    
    If KeyAscii <> vbKeyReturn Then
        zlControl.TxtCheckKeyPress txt住院次数, KeyAscii, m数字式
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If Trim(txt住院次数.Text) <> "" Then
            lng次数 = Val(txt住院次数.Text)
            If Check住院号(mstr住院号, lng次数) Then
                If ReadCard(mlng病人ID, mlng主页ID) Then
                    If Trim(txt状态.Text) = "在院" Then
                        cmd出院.Enabled = True
                        cmd修改.Enabled = False
                    Else
                        cmd出院.Enabled = False
                        cmd修改.Enabled = True
                    End If
                End If
                zlCommFun.PressKey (vbKeyTab)
            Else
                MsgBox "输入的次数不存在住院信息，请正确输入！", vbInformation, gstrSysName
'               txt住院次数.Text = ""
                txt住院次数.SetFocus
            End If
        Else
            txt住院次数.Text = ""
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
    zlCommFun.OpenIme False
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    Dim str住院号 As String
    
    If KeyAscii <> vbKeyReturn Then
        zlControl.TxtCheckKeyPress txt住院号, KeyAscii, m数字式
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If Trim(txt住院号.Text) <> "" Then
            str住院号 = txt住院号.Text
            If Check住院号(str住院号) Then
                If mlng主页ID > 0 And Check住院号(str住院号, mlng主页ID) Then
                    If ReadCard(mlng病人ID, mlng主页ID) Then
                        If Trim(txt状态.Text) = "在院" Then
                            cmd出院.Enabled = True
                            cmd修改.Enabled = False
                        Else
                            cmd出院.Enabled = False
                            cmd修改.Enabled = True
                        End If
                    End If
                    zlCommFun.PressKey (vbKeyTab)
                Else
                    MsgBox "输入的次数不存在住院信息，请正确输入！", vbInformation, gstrSysName
'                    txt住院次数.Text = ""
                    txt住院次数.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "输入的住院号不存在住院病人信息，请正确输入住院号！", vbInformation, gstrSysName
                txt住院号.SetFocus
                Exit Sub
            End If
            mstr住院号 = str住院号
        Else
            txt住院号.Text = ""
            txt住院次数.Text = ""
            zlCommFun.PressKey (vbKeyTab)
'            txt住院号.SetFocus
        End If
    End If
End Sub

Private Function Check住院号(ByVal str住院号 As String, Optional lng主页ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------
    '检查输入住院号是否存在
    '参数:bln病人-是否对当前的病人的住院号不进行判断
    '返回:存在住院号返回true,否则返回False
    '编制:
    '-----------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnLimitUnit As Boolean, strUnitIDs As String, strWhere As String
'    当前病区ID
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits & ","
        strWhere = " And instr([3],',' || 当前病区ID || ',')>0 "
    Else
        strWhere = ""
        strUnitIDs = ""
    End If
    
    On Error GoTo errHandle
    
    '问题30031 by lesfeng 2010-05-19 增加 病历接收
    If lng主页ID = 0 Then
        strSQL = "select 病人ID,住院号,max(主页ID) as 主页ID from 病案主页 where 住院号=[1] " & strWhere & " Group By 病人ID,住院号"
    Else
        strSQL = "select 病人ID,住院号,主页ID,decode(出院日期,Null,'在院','出院') as 状态 from 病案主页 where 住院号=[1] and 主页id=[2] " & strWhere
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str住院号, lng主页ID, strUnitIDs)
    If Not rsTemp.EOF Then
        If rsTemp.RecordCount = 0 Then
            Check住院号 = False
            rsTemp.Close
            Exit Function
        End If
        If lng主页ID = 0 Then
            mlng病人ID = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
            mlng主页ID = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
            mstr状态 = ""
        Else
            mlng病人ID = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
            mlng主页ID = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
            mstr状态 = IIf(IsNull(rsTemp!状态), "", rsTemp!状态)
        End If
        If mlng病人ID = 0 Then
            Check住院号 = False
            rsTemp.Close
            Exit Function
        End If
        txt住院次数.Text = mlng主页ID
        txt状态.Text = mstr状态
        Call Check病历(mlng病人ID, mlng主页ID)
    Else
        Check住院号 = False
        chk接收.Value = 0
        cmd接收.Enabled = False
        chk接收.Enabled = False
        rsTemp.Close
        Exit Function
    End If
    Check住院号 = True
    Call Get住院次数
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check病历(ByVal lng病人ID As Long, Optional lng主页ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    '问题30031 by lesfeng 2010-05-19 增加 病历接收
    chk接收.Value = 0
    mint病历接收 = 0
    On Error GoTo errHandle
    strSQL = "select 信息值 from 病案主页从表 where 病人id=[1] and 主页id = [2] and 信息名 = '病历接收'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        If rsTemp.RecordCount = 0 Then
            Check病历 = False
            rsTemp.Close
            Exit Function
        End If
        mint病历接收 = Val(IIf(IsNull(rsTemp!信息值), 0, rsTemp!信息值))
        chk接收.Value = mint病历接收
        cmd接收.Enabled = False
        chk接收.Enabled = True
    Else
        chk接收.Enabled = True
        Check病历 = False
        rsTemp.Close
        Exit Function
    End If
    Check病历 = True

    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get住院次数() As Boolean
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    If mlng主页ID > 0 Then
        '获取就诊次数(主页ID中的内容可能存在不连续)
        strSQL = "Select 主页ID,病人性质 From 病案主页 Where 病人ID=[1] And NVL(主页ID,0)<>0 Order by 主页ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取就诊次数", mlng病人ID)
        
        If rsTemp.RecordCount > 0 Then
            With cbo(cboName.主页ID)
                .Clear
                For i = 1 To rsTemp.RecordCount
                    .AddItem "第" & rsTemp!主页ID & "次住院" & IIf(Nvl(rsTemp!病人性质, 0) = 1, "(门诊留观)", IIf(Nvl(rsTemp!病人性质, 0) = 2, "(住院留观)", ""))
                    .ItemData(.NewIndex) = Val(rsTemp!主页ID)
                    If Val(rsTemp!主页ID) = mlng主页ID Then .ListIndex = .NewIndex '不会再调用readcard
                    rsTemp.MoveNext
                Next
            End With
            cbo(cboName.主页ID).Enabled = True
            cbo(cboName.主页ID).Locked = False
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
