VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   63
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      TabHeight       =   882
      TabMaxWidth     =   3528
      MouseIcon       =   "frmDegreeCard.frx":0E42
      TabCaption(0)   =   "基本信息(&1)"
      TabPicture(0)   =   "frmDegreeCard.frx":0E5E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl区域"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl年龄"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl性别"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl姓名"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl门诊号"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl医疗付款"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl出生日期"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl出生地点"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl身份证号"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl身份"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl职业"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl民族"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl国籍"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl学历"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lvl婚姻状况"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl家庭地址"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl家庭电话"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl家庭地址邮编"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl联系人姓名"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl联系人关系"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl联系人地址"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl联系人电话"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl工作单位"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl单位电话"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl单位邮编"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl单位开户行"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl单位帐号"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label11"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label10"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label2(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label32"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label33"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label27"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl登记时间"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label36"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl其他证件"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label40"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label41"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lbl户口地址"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "lbl户口地址邮编"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "lbl籍贯"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "lbl联系人身份证"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "lbl监护人"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lblMobile"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "chk担保"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cbo(1)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt(9)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txt(8)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt(3)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt(1)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt(0)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt(2)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txt(5)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt(4)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt(7)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt(23)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt(15)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt(14)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt(13)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt(10)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt(11)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txt(19)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt(20)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txt(22)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txt(25)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txt(27)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txt(28)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txt(21)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txt(17)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txt(32)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txt(31)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txt(34)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txt(33)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txt(59)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txt(41)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txt(64)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txt(63)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txt(60)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txt(30)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txt(29)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txt(26)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "txt(24)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txt(18)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txt(12)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txt(16)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "txt(6)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "txt(65)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "txt(69)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "txt(70)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "txt(71)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "txt(72)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "txt(73)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "txt(74)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "txt(75)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "txt(76)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "txt(77)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).ControlCount=   101
      TabCaption(1)   =   "住院信息(&2)"
      TabPicture(1)   =   "frmDegreeCard.frx":1738
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt(58)"
      Tab(1).Control(1)=   "txt(55)"
      Tab(1).Control(2)=   "txt(56)"
      Tab(1).Control(3)=   "txt(68)"
      Tab(1).Control(4)=   "txt(67)"
      Tab(1).Control(5)=   "txt(66)"
      Tab(1).Control(6)=   "txt(53)"
      Tab(1).Control(7)=   "txt(46)"
      Tab(1).Control(8)=   "txt(52)"
      Tab(1).Control(9)=   "txt(54)"
      Tab(1).Control(10)=   "txt(50)"
      Tab(1).Control(11)=   "txt(51)"
      Tab(1).Control(12)=   "txt(38)"
      Tab(1).Control(13)=   "txt(62)"
      Tab(1).Control(14)=   "txt(61)"
      Tab(1).Control(15)=   "txt(57)"
      Tab(1).Control(16)=   "cbo(0)"
      Tab(1).Control(17)=   "txt(39)"
      Tab(1).Control(18)=   "txt(47)"
      Tab(1).Control(19)=   "txt(48)"
      Tab(1).Control(20)=   "txt(49)"
      Tab(1).Control(21)=   "txt(42)"
      Tab(1).Control(22)=   "txt(43)"
      Tab(1).Control(23)=   "txt(44)"
      Tab(1).Control(24)=   "txt(35)"
      Tab(1).Control(25)=   "txt(40)"
      Tab(1).Control(26)=   "txt(36)"
      Tab(1).Control(27)=   "txt(37)"
      Tab(1).Control(28)=   "txt(45)"
      Tab(1).Control(29)=   "Label39"
      Tab(1).Control(30)=   "Label38"
      Tab(1).Control(31)=   "Label37"
      Tab(1).Control(32)=   "Label12"
      Tab(1).Control(33)=   "Label35"
      Tab(1).Control(34)=   "Label34"
      Tab(1).Control(35)=   "Label5"
      Tab(1).Control(36)=   "Label31"
      Tab(1).Control(37)=   "Label30"
      Tab(1).Control(38)=   "Label28"
      Tab(1).Control(39)=   "Label26"
      Tab(1).Control(40)=   "Label16"
      Tab(1).Control(41)=   "Label24"
      Tab(1).Control(42)=   "Label23"
      Tab(1).Control(43)=   "Label22"
      Tab(1).Control(44)=   "Label21"
      Tab(1).Control(45)=   "Label29"
      Tab(1).Control(46)=   "Label25"
      Tab(1).Control(47)=   "Label20"
      Tab(1).Control(48)=   "Label19"
      Tab(1).Control(49)=   "Label18"
      Tab(1).Control(50)=   "Label17"
      Tab(1).Control(51)=   "Label14"
      Tab(1).Control(52)=   "Label15"
      Tab(1).Control(53)=   "lbl费别"
      Tab(1).Control(54)=   "Label4"
      Tab(1).Control(55)=   "Label3"
      Tab(1).Control(56)=   "Label6"
      Tab(1).ControlCount=   57
      TabCaption(2)   =   "合并记录(&3)"
      TabPicture(2)   =   "frmDegreeCard.frx":1A52
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "msfMerge"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   77
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         Top             =   5940
         Width           =   1950
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   76
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   46
         Top             =   5940
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   75
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   5250
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   74
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   38
         Top             =   5250
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   73
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   72
         Left            =   1335
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   28
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   71
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3840
         Width           =   3705
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   70
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   5940
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   69
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5940
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   58
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   55
         Left            =   -70515
         Locked          =   -1  'True
         TabIndex        =   151
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   56
         Left            =   -68310
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   5520
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   68
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   148
         Top             =   5520
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   67
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   143
         Top             =   4780
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   66
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   145
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   65
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2415
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   53
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   142
         Top             =   4410
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   46
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   2415
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   52
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   4020
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   54
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   5150
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   50
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   3240
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   51
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   3630
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   38
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   135
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   62
         Left            =   -66050
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   16
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2415
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2760
         Width           =   3705
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   24
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4905
         Width           =   3705
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2760
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
         Index           =   30
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3825
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   60
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4890
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   64
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1358
         Width           =   1275
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
         Index           =   61
         Left            =   -70965
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   57
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2400
         Width           =   1245
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
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   -71865
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   780
         Width           =   2160
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   39
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1605
         Width           =   4035
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   47
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   48
         Left            =   -70965
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   49
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   42
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   43
         Left            =   -70965
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   44
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   35
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   40
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   36
         Left            =   -73760
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   37
         Left            =   -70965
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   45
         Left            =   -68480
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2415
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   33
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   41
         Top             =   5595
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
         TabIndex        =   42
         Top             =   5595
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   31
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   5595
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   32
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   5595
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   17
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   14
         Top             =   2056
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   21
         Left            =   1335
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   28
         Left            =   8970
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   22
         Top             =   3105
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   27
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Top             =   3105
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   25
         Left            =   3765
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   32
         Top             =   4530
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   22
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4545
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   3765
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3120
         Width           =   3705
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1707
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   13
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2056
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   14
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2056
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   15
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2415
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   23
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5250
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1358
         Width           =   3705
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1009
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1009
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1009
         Width           =   680
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1358
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1707
         Width           =   1275
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   4455
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1009
         Width           =   580
      End
      Begin VB.CheckBox chk担保 
         BackColor       =   &H8000000A&
         Caption         =   "临时"
         Enabled         =   0   'False
         Height          =   345
         Left            =   9600
         MaskColor       =   &H00000000&
         TabIndex        =   43
         Top             =   5558
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfMerge 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   127
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9551
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
         MouseIcon       =   "frmDegreeCard.frx":1DEC
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblMobile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手机号"
         Height          =   180
         Left            =   7665
         TabIndex        =   160
         Top             =   6000
         Width           =   540
      End
      Begin VB.Label lbl监护人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "监护人"
         Height          =   180
         Left            =   5715
         TabIndex        =   159
         Top             =   6000
         Width           =   540
      End
      Begin VB.Label lbl联系人身份证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人身份证"
         Height          =   180
         Left            =   5175
         TabIndex        =   158
         Top             =   5310
         Width           =   1080
      End
      Begin VB.Label lbl籍贯 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "籍贯"
         Height          =   180
         Left            =   3360
         TabIndex        =   157
         Top             =   4260
         Width           =   360
      End
      Begin VB.Label lbl户口地址邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址邮编"
         Height          =   180
         Left            =   180
         TabIndex        =   156
         Top             =   4260
         Width           =   1080
      End
      Begin VB.Label lbl户口地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   180
         Left            =   540
         TabIndex        =   155
         Top             =   3900
         Width           =   720
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院未结费用"
         Height          =   180
         Left            =   2640
         TabIndex        =   154
         Top             =   6000
         Width           =   1080
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院预交余额"
         Height          =   180
         Left            =   180
         TabIndex        =   153
         Top             =   6000
         Width           =   1080
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医出院情况"
         Height          =   180
         Left            =   -74880
         TabIndex        =   149
         Top             =   5560
         Width           =   1080
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   -74520
         TabIndex        =   147
         Top             =   4840
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型"
         Height          =   180
         Left            =   -66870
         TabIndex        =   146
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lbl其他证件 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他证件"
         Height          =   180
         Left            =   5520
         TabIndex        =   144
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   -74520
         TabIndex        =   141
         Top             =   4460
         Width           =   720
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡号"
         Height          =   180
         Left            =   8205
         TabIndex        =   129
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lbl登记时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
         Height          =   180
         Left            =   8205
         TabIndex        =   128
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   7665
         TabIndex        =   98
         Top             =   5655
         Width           =   540
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病况"
         Height          =   180
         Left            =   -66840
         TabIndex        =   126
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊医师"
         Height          =   180
         Left            =   5535
         TabIndex        =   125
         Top             =   4240
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   -74520
         TabIndex        =   124
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主任(副主任)医师"
         Height          =   180
         Left            =   -72450
         TabIndex        =   123
         Top             =   2460
         Width           =   1440
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊中医诊断"
         Height          =   180
         Left            =   5175
         TabIndex        =   122
         Top             =   4950
         Width           =   1080
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊诊断"
         Height          =   180
         Left            =   5535
         TabIndex        =   121
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   -68925
         TabIndex        =   120
         Top             =   1665
         Width           =   360
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院中医诊断"
         Height          =   180
         Left            =   -74880
         TabIndex        =   119
         Top             =   3690
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   -74520
         TabIndex        =   118
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院中医诊断"
         Height          =   180
         Left            =   -74880
         TabIndex        =   117
         Top             =   5220
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院目的"
         Height          =   180
         Left            =   -74520
         TabIndex        =   116
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   -71370
         TabIndex        =   115
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院科室"
         Height          =   180
         Left            =   -69120
         TabIndex        =   114
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   -74520
         TabIndex        =   113
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   -71280
         TabIndex        =   112
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   -74520
         TabIndex        =   111
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   -71730
         TabIndex        =   110
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   -69105
         TabIndex        =   109
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院科室"
         Height          =   180
         Left            =   -71730
         TabIndex        =   108
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院病况"
         Height          =   180
         Left            =   -69285
         TabIndex        =   107
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转科信息"
         Height          =   180
         Left            =   -74520
         TabIndex        =   106
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院天数"
         Height          =   180
         Left            =   -66885
         TabIndex        =   105
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         Height          =   180
         Left            =   -74520
         TabIndex        =   104
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   -68925
         TabIndex        =   103
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   -74340
         TabIndex        =   102
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   -69285
         TabIndex        =   101
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   -66870
         TabIndex        =   100
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Index           =   1
         Left            =   5715
         TabIndex        =   99
         Top             =   5655
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊预交余额"
         Height          =   180
         Left            =   180
         TabIndex        =   97
         Top             =   5655
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊未结费用"
         Height          =   180
         Left            =   2640
         TabIndex        =   96
         Top             =   5655
         Width           =   1080
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   5535
         TabIndex        =   95
         Top             =   3885
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   5355
         TabIndex        =   94
         Top             =   3540
         Width           =   900
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   8205
         TabIndex        =   93
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   5535
         TabIndex        =   92
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   5535
         TabIndex        =   91
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   2820
         TabIndex        =   90
         Top             =   4590
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   360
         TabIndex        =   89
         Top             =   4965
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   360
         TabIndex        =   88
         Top             =   5310
         Width           =   900
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   360
         TabIndex        =   87
         Top             =   4605
         Width           =   900
      End
      Begin VB.Label lbl家庭地址邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址邮编"
         Height          =   180
         Left            =   2640
         TabIndex        =   86
         Top             =   3540
         Width           =   1080
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   540
         TabIndex        =   85
         Top             =   3540
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   180
         Left            =   720
         TabIndex        =   84
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   540
         TabIndex        =   83
         Top             =   2115
         Width           =   720
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   8565
         TabIndex        =   82
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   3360
         TabIndex        =   81
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   5895
         TabIndex        =   80
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   3360
         TabIndex        =   79
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   900
         TabIndex        =   78
         Top             =   2475
         Width           =   360
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   5535
         TabIndex        =   77
         Top             =   2116
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   540
         TabIndex        =   76
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   3000
         TabIndex        =   75
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付费方式"
         Height          =   180
         Left            =   8205
         TabIndex        =   74
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   720
         TabIndex        =   73
         Top             =   1425
         Width           =   540
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   5715
         TabIndex        =   72
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   3360
         TabIndex        =   71
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   900
         TabIndex        =   70
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3360
         TabIndex        =   69
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊费别"
         Height          =   180
         Left            =   5535
         TabIndex        =   68
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   720
         TabIndex        =   67
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保险类别"
         Height          =   180
         Left            =   5535
         TabIndex        =   66
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   180
         Left            =   900
         TabIndex        =   65
         Top             =   1770
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退  出(&X)"
      Height          =   450
      Left            =   8975
      TabIndex        =   0
      Top             =   6600
      Width           =   1600
   End
End
Attribute VB_Name = "frmDegreeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long '要查看的病人ID
Public mlng主页ID As Long '住院病人时传入主页ID

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
    门诊预交余额 = 31
    门诊费用余额 = 32
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
    '问题27832 by lesfeng 2010-01-18 增加“出院情况
    出院情况 = 67
    中医出院情况 = 68
    
    住院预交余额 = 69
    住院费用余额 = 70
    
    户口地址 = 71
    户口地址邮编 = 72
    籍贯 = 73
    '51163,刘鹏飞,2012-07-09,增加"联系人身份证号"
    联系人身份证号 = 74
    联系人附加信息 = 75
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
    Dim strPlus As String
    Dim blnPassShowCard As Boolean, strCard As String
    
    On Error GoTo errH
    
    '51572:刘鹏飞,2013-11-04,就诊卡是否密文显示
    strSQL = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        blnPassShowCard = Nvl(rsTmp!卡号密文) <> ""
    End If
    
    If blnPassShowCard = True Then
        strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡号,"
    Else
        strCard = "A.就诊卡号 as 就诊卡号,"
    End If
    
    '问题28365 by lesfeng 2010-03-04 a.住院号 修改为 b.住院号
    strSQL = "Select a.病人id," & IIf(lng主页ID = 0, " a.姓名", " NVL(B.姓名,a.姓名) 姓名") & "," & IIf(lng主页ID = 0, " a.性别", " NVL(b.性别,a.性别) 性别") & "," & IIf(lng主页ID = 0, " a.年龄", " NVL(b.年龄,a.年龄) 年龄") & ", a.门诊号, a.费别, a.医疗付款方式, a.险类, a.区域, a.国籍, a.民族, a.学历," & vbNewLine & _
            "            a.婚姻状况, a.职业, a.身份, Decode(To_Date(To_Char(出生日期, 'YYYY-MM-DD HH24:MI'), 'YYYY-MM-DD HH24:MI') - Trunc(出生日期), 0, To_Char(出生日期, 'YYYY-MM-DD'),To_char(出生日期,'YYYY-MM-DD HH24:MI')) 出生日期, " & _
            "            a.身份证号, a.出生地点, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.户口地址, a.户口地址邮编, a.籍贯,a.监护人, a.联系人姓名,a.联系人身份证号," & vbNewLine & _
            "            a.联系人关系, a.联系人地址, a.联系人电话, a.工作单位, a.单位电话, a.手机号, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人," & vbNewLine & _
            "            a.担保额, a.担保性质, a.住院次数,a.主页ID 就诊次数, b.住院号, To_char(a.登记时间,'yyyy-mm-dd hh24:mi:ss') As 登记时间," & strCard & " b.出院病床, b.备注, b.住院目的, b.门诊医师, b.住院医师," & vbNewLine & _
            "            b.责任护士, b.登记人, b.入院病况, b.当前病况, b.住院天数, b.费别 As 住院费别, c.门诊预交余额, c.门诊费用余额, c.住院预交余额, c.住院费用余额," & vbNewLine & _
            "            Nvl(A.医保号,d.信息值) 医保号, e.名称 As 护理等级, g.名称 As 床位等级, m.名称 As 入院科室, n.名称 As 出院科室," & vbNewLine & _
            "            To_char(b.入院日期,'yyyy-mm-dd hh24:mi:ss') 入院日期, To_char(b.出院日期,'yyyy-mm-dd hh24:mi:ss') 出院日期,A.其他证件,Nvl(Nvl(A.病人类型,B.病人类型),Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
            "From 病人信息 a, 病案主页 b, (select 病人ID, 性质,NVL(sum(门诊预交余额),0) 门诊预交余额,NVL(sum(门诊费用余额),0) 门诊费用余额,NVL(sum(住院预交余额),0) 住院预交余额, NVL(sum(住院费用余额),0) 住院费用余额" & vbNewLine & _
            "           from (select 病人ID, 性质, 类型, case when 类型 = 1 then 预交余额 end 门诊预交余额, case when 类型 = 2 then 预交余额 end 住院预交余额, case when 类型 = 1 then 费用余额 end 门诊费用余额," & vbNewLine & _
            "           case when 类型 = 2 then 费用余额 end 住院费用余额 from 病人余额 where 病人ID = [1]) group by 病人ID, 性质) c, 病案主页从表 d, 收费项目目录 e, 床位状况记录 f, 收费项目目录 g, 部门表 m, 部门表 n" & vbNewLine & _
            "Where a.病人id = b.病人id(+) And " & IIf(lng主页ID = 0, "Nvl(a.主页ID,0)", "[2]") & "=b.主页ID(+) And a.病人ID=[1] And a.病人id = c.病人id(+) And" & vbNewLine & _
            "           c.性质(+) = 1 And b.病人id = d.病人id(+) And" & vbNewLine & _
            "           b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.入院科室id = m.Id(+) And b.出院科室id = n.Id(+) And" & vbNewLine & _
            "           b.护理等级id = e.ID(+) And b.当前病区id = f.病区id(+) And b.出院病床 = f.床号(+) And f.等级id = g.ID(+)"
        
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTmp.EOF Then Exit Function
        
    If bln查看某次住院 Then
       strTxt = "住院号=36,出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53,出院中医诊断=54,病人类型=66"
    Else
        strTxt = "病人ID=0,姓名=1,性别=2,年龄=3,门诊号=4,费别=5,医疗付款方式=6,医保号=7,险类=8,区域=9,国籍=10,民族=11,学历=12,婚姻状况=13,职业=14," & _
                " 身份=15,出生日期=16,身份证号=17,出生地点=18,家庭地址=19,家庭地址邮编=20,家庭电话=21,联系人姓名=22,联系人关系=23,联系人地址=24,联系人电话=25," & _
                " 工作单位=26,单位电话=27,单位邮编=28,单位开户行=29,单位帐号=30,门诊预交余额=31,门诊费用余额=32,担保人=33,担保额=34,住院次数=35,住院号=36," & _
                " 出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53," & _
                " 出院中医诊断=54,登记时间=63,就诊卡号=64,其他证件=65,病人类型=66,住院预交余额=69,住院费用余额=70,户口地址=71,户口地址邮编=72,籍贯=73,联系人身份证号=74,联系人附加信息=75,监护人=76,手机号=77"
    End If
    
    arrTxt = Split(strTxt, ",")
    
    For i = 0 To UBound(arrTxt)
        strTmp = Trim(arrTxt(i))
        
        If strTmp <> "" Then
            '排开暂不处理的字段
            If InStr(1, ",门诊诊断,门诊中医诊断,入院诊断,入院中医诊断,出院诊断,出院中医诊断,转科信息,联系人附加信息,", "," & Trim(Split(strTmp, "=")(0)) & ",") = 0 Then
                If InStr(1, ",门诊费用余额,门诊预交余额,住院费用余额,住院预交余额,", "," & Trim(Split(strTmp, "=")(0)) & ",") > 0 Then
                    txt(Trim(Split(strTmp, "=")(1))).Text = Format(Val("" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))), "0.00")
                Else
                    txt(Trim(Split(strTmp, "=")(1))).Text = "" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))
                End If
            End If
        End If
    Next
    '74426:李南春,2014-7-9,病人姓名颜色处理
    txt(txtName.姓名).ForeColor = gobjDatabase.GetPatiColor(Nvl(rsTmp!病人类型), True)
    '其它专门处理
    '----------------------------------------------
    '功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Call gobjControl.LoadOldData("" & rsTmp!年龄, txt(txtName.年龄), cbo(cboName.年龄单位))
    If cbo(cboName.年龄单位).ListIndex = -1 Then txt(txtName.年龄).Width = txt(txtName.年龄).Width + cbo(cboName.年龄单位).Width
    chk担保.Value = Val("" & rsTmp!担保性质)
    
    '病人信息从表
    strPlus = ""
    If txt(txtName.联系人关系).Text = "其他" Then
        txt(txtName.联系人附加信息).Visible = True
        strPlus = strPlus & "," & "联系人附加信息"
    Else
        txt(txtName.联系人附加信息).Visible = False
    End If
    If txt(txtName.身份证号).Text = "" Then
        strPlus = strPlus & "," & "身份证号状态"
    End If
    If strPlus <> "" Then strPlus = Mid(strPlus, 2)
    
    If strPlus <> "" Then
        Set rsTmp = Get病人信息从表(lng病人ID, strPlus)
        rsTmp.Filter = "信息名='联系人附加信息'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!信息值) Then txt(txtName.联系人附加信息).Text = rsTmp!信息值
        End If
        rsTmp.Filter = "信息名='身份证号状态'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!信息值) Then
                txt(txtName.身份证号).Text = zlCommFun.GetNeedName(rsTmp!信息值)
            End If
        End If
    End If
    
    '住院信息
    '----------------------------------------------
'    If Not bln查看某次住院 Then lng主页Id =Val(Nvl(rsTmp!就诊次数, 0))
    '住院病人的诊断情况
    If lng主页ID > 0 Then
        '问题27832 by lesfeng 2010-01-18 增加“出院情况”
        strSQL = "Select 诊断类型,疾病ID,诊断描述,出院情况 From 病人诊断记录 Where 诊断次序=1 And 记录来源 In(2,3)  And NVL(编码序号,1) = 1 And 病人ID=[1] And 主页ID=[2] Order By 记录来源 Desc"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        
        txt(txtName.门诊诊断).Text = ""
        txt(txtName.门诊中医诊断).Text = ""
        txt(txtName.入院诊断).Text = ""
        txt(txtName.入院中医诊断).Text = ""
        txt(txtName.出院诊断).Text = ""
        txt(txtName.出院中医诊断).Text = ""
        '问题27832 by lesfeng 2010-01-18 增加“出院情况
        txt(txtName.出院情况).Text = ""
        txt(txtName.中医出院情况).Text = ""
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
                'If j <> 0 And Trim(txt(j).Text) = "" Then txt(j).Text = IIf(IsNull(rsTmp!疾病ID), "", "(" & rsTmp!疾病ID & ")") & rsTmp!诊断描述
                If j <> 0 Then txt(j).Text = Nvl(rsTmp!诊断描述)
                
                rsTmp.MoveNext
            Next
        End If
        
        '问题28139 by lesfeng 2010-02-04
        strSQL = "Select 诊断类型,疾病ID,诊断描述,出院情况 From 病人诊断记录 Where 诊断类型 in (3,13) and 诊断次序>1  And NVL(编码序号,1) = 1 And 记录来源=3 And 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 诊断类型,疾病ID,诊断描述,出院情况 From 病人诊断记录 Where 诊断类型 in (3,13) and 诊断次序>1  And NVL(编码序号,1) = 1 And 记录来源=2 And 病人ID=[1] And 主页ID=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        End If

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
                    'str其它诊断 = IIf(IsNull(rsTmp!疾病ID), "", "(" & rsTmp!疾病ID & ")") & rsTmp!诊断描述 & IIf(IsNull(rsTmp!出院情况), "", "(" & rsTmp!出院情况 & ")")
                    str其它诊断 = Nvl(rsTmp!诊断描述) & IIf(IsNull(rsTmp!出院情况), "", "(" & rsTmp!出院情况 & ")")
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
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
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
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        rsTmp.Filter = "信息名='主治医师'"
        If Not rsTmp.EOF Then txt(txtName.主治医师).Text = "" & rsTmp!信息值
        rsTmp.Filter = "信息名='主任医师'"
        If Not rsTmp.EOF Then txt(txtName.主任医师).Text = "" & rsTmp!信息值
    End If
    
    
    '3.病人合并信息
    If Not bln查看某次住院 Then
        
        strSQL = "Select 原信息,合并原因,操作员姓名,合并时间 From 病人合并记录 Where 病人ID=[1] Order by 合并时间 Desc"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
                
        strHead = "合并时间,1,1800|操作员,4,800|合并原因,1,1800|" & _
                "病人ID,1,800|门诊号,1,900|住院号,1,800|就诊卡号,1,900|姓名,4,800|" & _
                "性别,4,500|年龄,4,800|出生日期,1,1000|身份证号,1,1800|婚姻状况,4,900|职业,1,1000|家庭地址,1,4200"
        With msfMerge
            .Redraw = False
            .Rows = rsTmp.RecordCount + 1
            
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
                .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
                .ColAlignmentFixed(i) = 4
            Next
            
            Call gobjComlib.RestoreFlexState(msfMerge, App.ProductName & "\" & Me.Name)
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub cbo_Click(Index As Integer)
    If Index = cboName.主页ID Then      '启动加载就诊次数时不调用
        If cbo(cboName.主页ID).Visible Then Call ReadCard(mlng病人ID, cbo(cboName.主页ID).ItemData(cbo(cboName.主页ID).ListIndex), True)
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    '固定信息处理
    With cbo(cboName.年龄单位)
        .AddItem "岁"
        .AddItem "月"
        .AddItem "天"
        .ListIndex = 0
    End With
    
    If Not ReadCard(mlng病人ID, mlng主页ID) Then
        MsgBox "不能正确读取病人信息,请与系统管理员联系！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    '住院信息
    '获取就诊次数(主页ID中的内容可能存在不连续)
    strSQL = "Select 主页ID,病人性质 From 病案主页 Where 病人ID=[1] And NVL(主页ID,0)<>0 Order by 主页ID"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取就诊次数", mlng病人ID)
    
    If rsTemp.RecordCount > 0 Then
        With cbo(cboName.主页ID)
            For i = 1 To rsTemp.RecordCount
                .AddItem "第" & rsTemp!主页ID & "次住院" & IIf(Nvl(rsTemp!病人性质, 0) = 1, "(门诊留观)", IIf(Nvl(rsTemp!病人性质, 0) = 2, "(住院留观)", ""))
                .ItemData(.NewIndex) = Val(rsTemp!主页ID)
                If Val(rsTemp!主页ID) = mlng主页ID Then .ListIndex = .NewIndex '不会再调用readcard
                rsTemp.MoveNext
            Next
        End With
        cbo(cboName.主页ID).Enabled = True
        cbo(cboName.主页ID).Locked = False
    Else
        cbo(cboName.主页ID).Enabled = False
        cbo(cboName.主页ID).Locked = True
        SSTab1.TabVisible(1) = False
        SSTab1.TabCaption(2) = "合并记录(&2)"
    End If
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mlng病人ID = 0
    mlng主页ID = 0
    
    Call gobjComlib.SaveFlexState(msfMerge, App.ProductName & "\" & Me.Name)
End Sub

