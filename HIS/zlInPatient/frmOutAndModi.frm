VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOutAndModi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ժ��������Ժ"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "����ȷ��(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5920
      TabIndex        =   160
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�����ѽ���"
      Height          =   255
      Left            =   4680
      TabIndex        =   159
      Top             =   195
      Width           =   1215
   End
   Begin VB.CommandButton cmd�޸� 
      Caption         =   "��  ��(&M)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8360
      TabIndex        =   4
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmd��Ժ 
      Caption         =   "��  Ժ(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.TextBox txt״̬ 
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.TextBox txtסԺ���� 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtסԺ�� 
      Height          =   375
      Left            =   720
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "��  ��(&X)"
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
      TabCaption(0)   =   "������Ϣ(&1)"
      TabPicture(0)   =   "frmOutAndModi.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����֤��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label36"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl�Ǽ�ʱ��"
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
      Tab(0).Control(10)=   "lbl��λ�ʺ�"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl��λ������"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl��λ�ʱ�"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl��λ�绰"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl������λ"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl��ϵ�˵绰"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl��ϵ�˵�ַ"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl��ϵ�˹�ϵ"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl��ϵ������"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl��ͥ��ַ�ʱ�"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl��ͥ�绰"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl��ͥ��ַ"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lvl����״��"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblѧ��"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl����"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl����"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblְҵ"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl���"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl���֤��"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl�����ص�"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl��������"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblҽ�Ƹ���"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label13"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl�����"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl����"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl�Ա�"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl����"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label9"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label8"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label7"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl����"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lbl����"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lbl���ڵ�ַ�ʱ�"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lbl���ڵ�ַ"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "lbl��ϵ�����֤��"
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
      Tab(0).Control(87)=   "chk����"
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
      TabCaption(1)   =   "סԺ��Ϣ(&2)"
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
      Tab(1).Control(24)=   "lbl�ѱ�"
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
      TabCaption(2)   =   "�ϲ���¼(&3)"
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
      Begin VB.CheckBox chk���� 
         BackColor       =   &H8000000A&
         Caption         =   "��ʱ"
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
      Begin VB.Label lbl��ϵ�����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�����֤"
         Height          =   180
         Left            =   5190
         TabIndex        =   167
         Top             =   5310
         Width           =   1080
      End
      Begin VB.Label lbl���ڵ�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ"
         Height          =   180
         Left            =   420
         TabIndex        =   166
         Top             =   3900
         Width           =   720
      End
      Begin VB.Label lbl���ڵ�ַ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ�ʱ�"
         Height          =   180
         Left            =   60
         TabIndex        =   165
         Top             =   4260
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3480
         TabIndex        =   164
         Top             =   4260
         Width           =   360
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   -71370
         TabIndex        =   158
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   180
         Left            =   -74340
         TabIndex        =   157
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         Height          =   180
         Left            =   -68040
         TabIndex        =   151
         Top             =   4840
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   780
         TabIndex        =   149
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   5535
         TabIndex        =   148
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
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
         Caption         =   "����ѱ�"
         Height          =   180
         Left            =   5535
         TabIndex        =   146
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3480
         TabIndex        =   145
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   780
         TabIndex        =   144
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3480
         TabIndex        =   143
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   5715
         TabIndex        =   142
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   600
         TabIndex        =   141
         Top             =   1425
         Width           =   540
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ѷ�ʽ"
         Height          =   180
         Left            =   8205
         TabIndex        =   140
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   3120
         TabIndex        =   139
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   420
         TabIndex        =   138
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   5535
         TabIndex        =   137
         Top             =   2116
         Width           =   720
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   780
         TabIndex        =   136
         Top             =   2475
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   3480
         TabIndex        =   135
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5895
         TabIndex        =   134
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3480
         TabIndex        =   133
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   8565
         TabIndex        =   132
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   420
         TabIndex        =   131
         Top             =   2115
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ"
         Height          =   180
         Left            =   600
         TabIndex        =   130
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   420
         TabIndex        =   129
         Top             =   3540
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ�ʱ�"
         Height          =   180
         Left            =   2760
         TabIndex        =   128
         Top             =   3540
         Width           =   1080
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   240
         TabIndex        =   127
         Top             =   4605
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   2940
         TabIndex        =   126
         Top             =   5310
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   240
         TabIndex        =   125
         Top             =   4965
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   240
         TabIndex        =   124
         Top             =   5310
         Width           =   900
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   5535
         TabIndex        =   123
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   5535
         TabIndex        =   122
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   8205
         TabIndex        =   121
         Top             =   3165
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   5355
         TabIndex        =   120
         Top             =   3540
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   5535
         TabIndex        =   119
         Top             =   3885
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   3120
         TabIndex        =   118
         Top             =   5760
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   420
         TabIndex        =   117
         Top             =   5760
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
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
         Caption         =   "����ȼ�"
         Height          =   180
         Left            =   -66870
         TabIndex        =   115
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   -69285
         TabIndex        =   114
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   -74340
         TabIndex        =   113
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   -68925
         TabIndex        =   112
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   -69240
         TabIndex        =   111
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   -66885
         TabIndex        =   110
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת����Ϣ"
         Height          =   180
         Left            =   -74520
         TabIndex        =   109
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   -69285
         TabIndex        =   108
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   -71730
         TabIndex        =   107
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���"
         Height          =   180
         Left            =   -69105
         TabIndex        =   106
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���λ�ʿ"
         Height          =   180
         Left            =   -71730
         TabIndex        =   105
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺҽʦ"
         Height          =   180
         Left            =   -74520
         TabIndex        =   104
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   -71280
         TabIndex        =   103
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   -74520
         TabIndex        =   102
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   -69120
         TabIndex        =   101
         Top             =   5560
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   -71370
         TabIndex        =   100
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺĿ��"
         Height          =   180
         Left            =   -74520
         TabIndex        =   99
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ҽ���"
         Height          =   180
         Left            =   -74880
         TabIndex        =   98
         Top             =   5220
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   -74520
         TabIndex        =   97
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ҽ���"
         Height          =   180
         Left            =   -74880
         TabIndex        =   96
         Top             =   3690
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Left            =   -68925
         TabIndex        =   95
         Top             =   1665
         Width           =   360
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   5535
         TabIndex        =   94
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ҽ���"
         Height          =   180
         Left            =   5175
         TabIndex        =   93
         Top             =   4950
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(������)ҽʦ"
         Height          =   180
         Left            =   -72450
         TabIndex        =   92
         Top             =   2460
         Width           =   1440
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   -74520
         TabIndex        =   91
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   5535
         TabIndex        =   90
         Top             =   4240
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   -66840
         TabIndex        =   89
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   7665
         TabIndex        =   88
         Top             =   5760
         Width           =   540
      End
      Begin VB.Label lbl�Ǽ�ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��"
         Height          =   180
         Left            =   8205
         TabIndex        =   87
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��"
         Height          =   180
         Left            =   8205
         TabIndex        =   86
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   -74520
         TabIndex        =   85
         Top             =   4470
         Width           =   720
      End
      Begin VB.Label lbl����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����֤��"
         Height          =   180
         Left            =   5520
         TabIndex        =   84
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   -66870
         TabIndex        =   83
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   -74520
         TabIndex        =   82
         Top             =   4840
         Width           =   720
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ��Ժ���"
         Height          =   180
         Left            =   -74880
         TabIndex        =   81
         Top             =   5560
         Width           =   1080
      End
   End
   Begin VB.Label lbl״̬ 
      BackStyle       =   0  'Transparent
      Caption         =   "״̬"
      Height          =   255
      Left            =   3360
      TabIndex        =   154
      Top             =   195
      Width           =   495
   End
   Begin VB.Label lblסԺ���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2160
      TabIndex        =   153
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lblסԺ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ��"
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
Option Explicit 'Ҫ���������
Private mlng����ID As Long 'Ҫ�鿴�Ĳ���ID
Private mlng��ҳID As Long 'סԺ����ʱ������ҳID
Private mstr״̬ As String
Private mstrסԺ�� As String
Private mstrPrivs As String
Private mlngModul As Long
Private mint�������� As Integer

Private Enum txtName
    'Ҫ���SQL���ֶζ�Ӧ
    ����ID = 0
    ���� = 1
    �Ա� = 2
    ���� = 3
    ����� = 4
    �ѱ� = 5
    ҽ�Ƹ��ʽ = 6
    ҽ���� = 7
    ���� = 8
    ���� = 9
    ���� = 10
    ���� = 11
    ѧ�� = 12
    ����״�� = 13
    ְҵ = 14
    ��� = 15
    �������� = 16
    ���֤�� = 17
    �����ص� = 18
    ��ͥ��ַ = 19
    ��ͥ��ַ�ʱ� = 20
    ��ͥ�绰 = 21
    ��ϵ������ = 22
    ��ϵ�˹�ϵ = 23
    ��ϵ�˵�ַ = 24
    ��ϵ�˵绰 = 25
    ������λ = 26
    ��λ�绰 = 27
    ��λ�ʱ� = 28
    ��λ������ = 29
    ��λ�ʺ� = 30
    Ԥ����� = 31
    ������� = 32
    ������ = 33
    ������ = 34
    
    סԺ���� = 35
    סԺ�� = 36
    ��Ժ���� = 37
    ��ע = 38
    סԺĿ�� = 39
    סԺ�ѱ� = 40
    ����ҽʦ = 41
    סԺҽʦ = 42
    ����ҽʦ = 57
    ����ҽʦ = 61
    ���λ�ʿ = 43
    �Ǽ��� = 44
    ��λ�ȼ� = 45
    ����ȼ� = 46
    ��Ժ���� = 47
    ��Ժ���� = 48
    ��Ժ���� = 49
    ��ǰ���� = 62
    ת����Ϣ = 52
    ��Ժ���� = 55
    ��Ժ���� = 56
    סԺ���� = 58
    
    ������� = 59
    ������ҽ��� = 60
    ��Ժ��� = 50
    ��Ժ��ҽ��� = 51
    ��Ժ��� = 53
    ��Ժ��ҽ��� = 54
    ��Ժ��ʽ = 69
    
    ��Ժ��� = 67
    ��ҽ��Ժ��� = 68
    '����51167,������,2012-07-09,����"��ϵ�����֤��"
    ��ϵ�����֤�� = 75
End Enum

Private Enum cboName
    ��ҳID = 0
    ���䵥λ = 1
End Enum

Private Function ReadCard(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln�鿴ĳ��סԺ As Boolean) As Boolean
'���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTxt As String, strTmp As String, strHead As String
    Dim i As Integer, j As Integer, arrTxt As Variant
    Dim str��Ժ��� As String
    Dim str������� As String
    Dim blnPassShowCard  As Boolean, strCard As String
    
    On Error GoTo errH
    
    '51572:������,2013-11-04,���￨�Ƿ�������ʾ
    strSQL = "Select �������� From ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        blnPassShowCard = Nvl(rsTmp!��������) <> ""
    End If
    
    If blnPassShowCard = True Then
        strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨��,"
    Else
        strCard = "A.���￨�� as ���￨��,"
    End If
        
    '����51167,������,2012-07-09,����"��ϵ�����֤��"
    strSQL = "Select a.����id, NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����, a.�����, a.�ѱ�, a.ҽ�Ƹ��ʽ, a.����, a.����, a.����, a.����, a.����, a.ѧ��," & vbNewLine & _
            "            a.����״��, a.ְҵ, a.���, Decode(To_Date(To_Char(��������, 'YYYY-MM-DD HH24:MI'), 'YYYY-MM-DD HH24:MI') - Trunc(��������), 0, To_Char(��������, 'YYYY-MM-DD'),To_char(��������,'YYYY-MM-DD HH24:MI')) ��������, " & _
            "            a.���֤��, a.�����ص�, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.��ϵ������," & vbNewLine & _
            "            a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰,a.��ϵ�����֤��, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������," & vbNewLine & _
            "            a.������, a.��������, a.סԺ����,a.��ҳID �������, b.סԺ��, To_char(a.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��," & strCard & "b.��Ժ����, b.��ע, b.סԺĿ��, b.����ҽʦ, b.סԺҽʦ," & vbNewLine & _
            "            b.���λ�ʿ, b.�Ǽ���, b.��Ժ����, b.��ǰ����, b.סԺ����, b.�ѱ� As סԺ�ѱ�, c.Ԥ�����, c.�������," & vbNewLine & _
            "            Nvl(A.ҽ����,d.��Ϣֵ) ҽ����, e.���� As ����ȼ�, g.���� As ��λ�ȼ�, m.���� As ��Ժ����, n.���� As ��Ժ����," & vbNewLine & _
            "            To_char(b.��Ժ����,'yyyy-mm-dd hh24:mi:ss') ��Ժ����, To_char(b.��Ժ����,'yyyy-mm-dd hh24:mi:ss') ��Ժ����,A.����֤��,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������,B.��Ժ��ʽ " & vbNewLine & _
            "From ������Ϣ a, ������ҳ b, (select ����ID,����,Nvl(sum(Ԥ�����),0) Ԥ�����,Nvl(sum(�������),0) ������� from ������� where ����ID=[1] and ����=2 group by ����ID,����) c, ������ҳ�ӱ� d, �շ���ĿĿ¼ e, ��λ״����¼ f, �շ���ĿĿ¼ g, ���ű� m, ���ű� n" & vbNewLine & _
            "Where a.����id = b.����id(+) And " & IIf(lng��ҳID = 0, "Nvl(a.��ҳID,0)", "[2]") & "=b.��ҳID(+) And a.����ID=[1] And a.����id = c.����id(+) And" & vbNewLine & _
            "           c.����(+) = 1 And b.����id = d.����id(+) And" & vbNewLine & _
            "           b.��ҳid = d.��ҳid(+) And d.��Ϣ��(+) = 'ҽ����' And b.��Ժ����id = m.Id(+) And b.��Ժ����id = n.Id(+) And" & vbNewLine & _
            "           b.����ȼ�id = e.ID(+) And b.��ǰ����id = f.����id(+) And b.��Ժ���� = f.����(+) And f.�ȼ�id = g.ID(+)"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTmp.EOF Then Exit Function
        
'    If bln�鿴ĳ��סԺ Then
'       strTxt = "��Ժ����=37,��ע=38,סԺĿ��=39,סԺ�ѱ�=40,����ҽʦ=41,סԺҽʦ=42,���λ�ʿ=43,�Ǽ���=44,��λ�ȼ�=45,����ȼ�=46,��Ժ����=47,��Ժ����=48," & _
'                " ��Ժ����=49,��ǰ����=62,ת����Ϣ=52,��Ժ����=55,��Ժ����=56,סԺ����=58,�������=59,������ҽ���=60,��Ժ���=50,��Ժ��ҽ���=51,��Ժ���=53,��Ժ��ҽ���=54,��������=66"
'    Else
        strTxt = "����ID=0,����=1,�Ա�=2,����=3,�����=4,�ѱ�=5,ҽ�Ƹ��ʽ=6,ҽ����=7,����=8,����=9,����=10,����=11,ѧ��=12,����״��=13,ְҵ=14," & _
                " ���=15,��������=16,���֤��=17,�����ص�=18,��ͥ��ַ=19,��ͥ��ַ�ʱ�=20,��ͥ�绰=21,��ϵ������=22,��ϵ�˹�ϵ=23,��ϵ�˵�ַ=24,��ϵ�˵绰=25," & _
                " ������λ=26,��λ�绰=27,��λ�ʱ�=28,��λ������=29,��λ�ʺ�=30,Ԥ�����=31,�������=32,������=33,������=34,סԺ����=35,סԺ��=36," & _
                " ��Ժ����=37,��ע=38,סԺĿ��=39,סԺ�ѱ�=40,����ҽʦ=41,סԺҽʦ=42,���λ�ʿ=43,�Ǽ���=44,��λ�ȼ�=45,����ȼ�=46,��Ժ����=47,��Ժ����=48," & _
                " ��Ժ����=49,��ǰ����=62,ת����Ϣ=52,��Ժ����=55,��Ժ����=56,סԺ����=58,�������=59,������ҽ���=60,��Ժ���=50,��Ժ��ҽ���=51,��Ժ���=53," & _
                " ��Ժ��ҽ���=54,�Ǽ�ʱ��=63,���￨��=64,����֤��=65,��������=66,��Ժ��ʽ=69,���ڵ�ַ=72,���ڵ�ַ�ʱ�=73,����=74,��ϵ�����֤��=75"
'    End If
    
    arrTxt = Split(strTxt, ",")
    
    For i = 0 To UBound(arrTxt)
        strTmp = Trim(arrTxt(i))
        
        If strTmp <> "" Then
            '�ſ��ݲ�������ֶ�
            If InStr(1, ",�������,������ҽ���,��Ժ���,��Ժ��ҽ���,��Ժ���,��Ժ��ҽ���,ת����Ϣ,", "," & Trim(Split(strTmp, "=")(0)) & ",") = 0 Then
                If InStr(1, ",�������,Ԥ�����,", "," & Trim(Split(strTmp, "=")(0)) & ",") > 0 Then
                    txt(Trim(Split(strTmp, "=")(1))).Text = Format(Val("" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))), "0.00")
                Else
                    txt(Trim(Split(strTmp, "=")(1))).Text = "" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))
                End If
            End If
        End If
    Next
    
    txt(70).Text = txt(0).Text
    txt(71).Text = txt(1).Text
    
    '����ר�Ŵ���
    '----------------------------------------------
    Call LoadOldData("" & rsTmp!����, txt(txtName.����), cbo(cboName.���䵥λ))
    If cbo(cboName.���䵥λ).ListIndex = -1 Then txt(txtName.����).width = txt(txtName.����).width + cbo(cboName.���䵥λ).width
    chk����.Value = Val("" & rsTmp!��������)
    
    
    'סԺ��Ϣ
    '----------------------------------------------
'    If Not bln�鿴ĳ��סԺ Then lng��ҳId = Val(Nvl(rsTmp!�������, 0))
    'סԺ���˵�������
    If lng��ҳID > 0 Then
        strSQL = "Select �������,����ID,�������,��Ժ��� From ������ϼ�¼ Where ��ϴ���=1 And ��¼��Դ=2 And ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!�������
                    Case 1
                        j = txtName.�������
                    Case 11
                        j = txtName.������ҽ���
                    Case 2
                        j = txtName.��Ժ���
                    Case 12
                        j = txtName.��Ժ��ҽ���
                    Case 3
                        j = txtName.��Ժ���
                        '����27832 by lesfeng 2010-01-18 ���ӡ���Ժ���
                        str��Ժ��� = IIf(IsNull(rsTmp!��Ժ���), "", rsTmp!��Ժ���)
                        txt(txtName.��Ժ���).Text = str��Ժ���
                    Case 13
                        j = txtName.��Ժ��ҽ���
                        '����27832 by lesfeng 2010-01-18 ���ӡ���Ժ���
                        str��Ժ��� = IIf(IsNull(rsTmp!��Ժ���), "", rsTmp!��Ժ���)
                        txt(txtName.��ҽ��Ժ���).Text = str��Ժ���
                    Case Else
                        j = 0
                End Select
                If j <> 0 Then txt(j).Text = IIf(IsNull(rsTmp!����ID), "", "(" & rsTmp!����ID & ")") & rsTmp!�������
                
                rsTmp.MoveNext
            Next
        Else
            txt(txtName.�������).Text = ""
            txt(txtName.������ҽ���).Text = ""
            txt(txtName.��Ժ���).Text = ""
            txt(txtName.��Ժ��ҽ���).Text = ""
            txt(txtName.��Ժ���).Text = ""
            txt(txtName.��Ժ��ҽ���).Text = ""
            '����27832 by lesfeng 2010-01-18 ���ӡ���Ժ���
            txt(txtName.��Ժ���).Text = ""
            txt(txtName.��ҽ��Ժ���).Text = ""
        End If
        
        '����28139 by lesfeng 2010-02-04
        strSQL = "Select �������,����ID,�������,��Ժ��� From ������ϼ�¼ Where ������� in (3,13) and ��ϴ���>1 And ��¼��Դ=2 And ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!�������
                    Case 3
                        j = txtName.��Ժ���
                    Case 13
                        j = txtName.��Ժ��ҽ���
                    Case Else
                        j = 0
                End Select
                                
                If j <> 0 Then
                    strTmp = txt(j).Text
                    str������� = IIf(IsNull(rsTmp!����ID), "", "(" & rsTmp!����ID & ")") & rsTmp!������� & IIf(IsNull(rsTmp!��Ժ���), "", "(" & rsTmp!��Ժ��� & ")")
                    If InStr(1, strTmp, "�������:") > 0 Then
                        txt(j).Text = strTmp & "," & str�������
                    Else
                        txt(j).Text = strTmp & ";�������:" & str�������
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        'ת����Ϣ
        txt(txtName.ת����Ϣ).Text = ""
        strSQL = _
            " Select Distinct 1 as ��ʼԭ��,To_Date('1900-01-01','YYYY-MM-DD') as ��ʼʱ��,B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=B.ID And A.��ʼʱ�� is Not NULL And A.��ʼԭ�� IN(1,2)" & _
            " And A.����ID=[1] And ��ҳID=[2]" & _
            " Union ALL " & _
            " Select A.��ʼԭ��,A.��ʼʱ��,B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=B.ID And A.��ʼʱ�� is Not NULL And A.��ʼԭ��=3" & _
            " And A.����ID=[1] And ��ҳID=[2]" & _
            " Order by ��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        rsTmp.Filter = "��ʼԭ��=3"
        If Not rsTmp.EOF Then
            rsTmp.Filter = 0
            Do While Not rsTmp.EOF
                txt(txtName.ת����Ϣ).Text = txt(txtName.ת����Ϣ).Text & " ���� " & rsTmp!����
                rsTmp.MoveNext
            Loop
            txt(txtName.ת����Ϣ).Text = Mid(txt(txtName.ת����Ϣ).Text, 5)
        End If
        
        '������ҳ�ӱ�
        txt(txtName.����ҽʦ).Text = ""
        txt(txtName.����ҽʦ).Text = ""
        strSQL = " Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where (��Ϣ��='����ҽʦ' Or ��Ϣ��='����ҽʦ') And ����ID=[1] And ��ҳID=[2]"
        rsTmp.Filter = ""
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        rsTmp.Filter = "��Ϣ��='����ҽʦ'"
        If Not rsTmp.EOF Then txt(txtName.����ҽʦ).Text = "" & rsTmp!��Ϣֵ
        rsTmp.Filter = "��Ϣ��='����ҽʦ'"
        If Not rsTmp.EOF Then txt(txtName.����ҽʦ).Text = "" & rsTmp!��Ϣֵ
    End If
    
    
    '3.���˺ϲ���Ϣ
    If Not bln�鿴ĳ��סԺ Then
        
        strSQL = "Select ԭ��Ϣ,�ϲ�ԭ��,����Ա����,�ϲ�ʱ�� From ���˺ϲ���¼ Where ����ID=[1] Order by �ϲ�ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
                
        strHead = "�ϲ�ʱ��,1,1800|����Ա,4,800|�ϲ�ԭ��,1,1800|" & _
                "����ID,1,800|�����,1,900|סԺ��,1,800|���￨��,1,900|����,4,800|" & _
                "�Ա�,4,500|����,4,800|��������,1,1000|���֤��,1,1800|����״��,4,900|ְҵ,1,1000|��ͥ��ַ,1,4200"
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
                .TextMatrix(i, 0) = Format(rsTmp!�ϲ�ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 1) = "" & rsTmp!����Ա����
                .TextMatrix(i, 2) = "" & rsTmp!�ϲ�ԭ��
                
                'v_ԭ��Ϣ:=r_InfoA.����Id || ',' || r_InfoA.����� || ',' ||  r_InfoA.סԺ�� || ',' ||  r_InfoA.���￨�� || ',' ||  r_InfoA.���� ||  ',' ||  r_InfoA.�Ա� ||  ',' ||
                '   r_InfoA.���� ||  ',' || to_char(r_InfoA.��������,'yyyy-mm-dd') ||  ',' || r_InfoA.���֤�� ||  ',' || r_InfoA.����״�� ||  ',' || r_InfoA.ְҵ ||  ',' || r_InfoA.��ͥ��ַ;
                arrTxt = Split(rsTmp!ԭ��Ϣ, ",")
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
    If Index = cboName.��ҳID Then      '��������סԺ����ʱ������
        If cbo(cboName.��ҳID).Visible Then Call ReadCard(mlng����ID, cbo(cboName.��ҳID).ItemData(cbo(cboName.��ҳID).ListIndex), True)
        If Trim(txt(55).Text) = "" Then txt״̬.Text = "��Ժ"
        If IsDate(Trim(txt(55).Text)) Then txt״̬.Text = "��Ժ"
        
        mstr״̬ = Trim(txt״̬.Text)
        mlng��ҳID = cbo(cboName.��ҳID).ItemData(cbo(cboName.��ҳID).ListIndex)
        mstrסԺ�� = txt(36).Text
        txtסԺ�� = mstrסԺ��
        txtסԺ���� = mlng��ҳID
        
        If Trim(txt״̬.Text) = "��Ժ" Then
            cmd��Ժ.Enabled = True
            cmd�޸�.Enabled = False
        Else
            cmd��Ժ.Enabled = False
            cmd�޸�.Enabled = True
        End If
    End If
End Sub

Private Sub chk����_Click()
    If mint�������� <> Val(chk����.Value) Then
        cmd����.Enabled = True
    Else
        cmd����.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmd��Ժ_Click()
    Dim lng����ID As Long, lng��ҳID As Long

    lng����ID = mlng����ID
    lng��ҳID = mlng��ҳID

    Call ExecPatiChange(EFun.E��Ժ, Me, mstrPrivs, lng����ID, lng��ҳID)
    If gblnOK Then
        cmd��Ժ.Enabled = False
        mlng����ID = 0
        mlng��ҳID = 0
        txtסԺ�� = ""
        txtסԺ���� = ""
        txt״̬ = ""
        txtסԺ��.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim int�������� As Integer
    Dim strSQL As String
    
    If InStr(mstrPrivs, "������Ժʱ��") <> 0 Then
    Else
        MsgBox "��û�С�������Ժʱ�䡯Ȩ�ޣ����ܸ��²��˲���������Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
    int�������� = chk����.Value
    If mlng����ID = 0 Then Exit Sub
    If mlng��ҳID = 0 Then Exit Sub
    
    On Error GoTo errH
    strSQL = "Zl_������ҳ_��������(" & mlng����ID & "," & mlng��ҳID & "," & int�������� & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mint�������� = int��������
    cmd����.Enabled = False
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd�޸�_Click()
    Dim lng����ID As Long, lng��ҳID As Long

    lng����ID = mlng����ID
    lng��ҳID = mlng��ҳID

    Call ExecPatiChange(EFun.E�޸ĳ�Ժʱ��, Me, mstrPrivs, lng����ID, lng��ҳID)
    If gblnOK Then
        cmd�޸�.Enabled = False
        mlng����ID = 0
        mlng��ҳID = 0
        txtסԺ�� = ""
        txtסԺ���� = ""
        txt״̬ = ""
        txtסԺ��.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    '�̶���Ϣ����
    With cbo(cboName.���䵥λ)
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .ListIndex = 0
    End With
    mlng����ID = 0
    mlng��ҳID = 0
    mint�������� = 0
    chk����.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mlng��ҳID = 0
    Call SaveFlexState(msfMerge, App.ProductName & "\" & Me.Name)
End Sub

Private Sub txtסԺ����_GotFocus()
    zlControl.TxtSelAll txtסԺ����
    zlCommFun.OpenIme False
End Sub

Private Sub txtסԺ����_KeyPress(KeyAscii As Integer)
    Dim lng���� As Long
    
    If KeyAscii <> vbKeyReturn Then
        zlControl.TxtCheckKeyPress txtסԺ����, KeyAscii, m����ʽ
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If Trim(txtסԺ����.Text) <> "" Then
            lng���� = Val(txtסԺ����.Text)
            If CheckסԺ��(mstrסԺ��, lng����) Then
                If ReadCard(mlng����ID, mlng��ҳID) Then
                    If Trim(txt״̬.Text) = "��Ժ" Then
                        cmd��Ժ.Enabled = True
                        cmd�޸�.Enabled = False
                    Else
                        cmd��Ժ.Enabled = False
                        cmd�޸�.Enabled = True
                    End If
                End If
                zlCommFun.PressKey (vbKeyTab)
            Else
                MsgBox "����Ĵ���������סԺ��Ϣ������ȷ���룡", vbInformation, gstrSysName
'               txtסԺ����.Text = ""
                txtסԺ����.SetFocus
            End If
        Else
            txtסԺ����.Text = ""
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
    zlCommFun.OpenIme False
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    Dim strסԺ�� As String
    
    If KeyAscii <> vbKeyReturn Then
        zlControl.TxtCheckKeyPress txtסԺ��, KeyAscii, m����ʽ
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If Trim(txtסԺ��.Text) <> "" Then
            strסԺ�� = txtסԺ��.Text
            If CheckסԺ��(strסԺ��) Then
                If mlng��ҳID > 0 And CheckסԺ��(strסԺ��, mlng��ҳID) Then
                    If ReadCard(mlng����ID, mlng��ҳID) Then
                        If Trim(txt״̬.Text) = "��Ժ" Then
                            cmd��Ժ.Enabled = True
                            cmd�޸�.Enabled = False
                        Else
                            cmd��Ժ.Enabled = False
                            cmd�޸�.Enabled = True
                        End If
                    End If
                    zlCommFun.PressKey (vbKeyTab)
                Else
                    MsgBox "����Ĵ���������סԺ��Ϣ������ȷ���룡", vbInformation, gstrSysName
'                    txtסԺ����.Text = ""
                    txtסԺ����.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "�����סԺ�Ų�����סԺ������Ϣ������ȷ����סԺ�ţ�", vbInformation, gstrSysName
                txtסԺ��.SetFocus
                Exit Sub
            End If
            mstrסԺ�� = strסԺ��
        Else
            txtסԺ��.Text = ""
            txtסԺ����.Text = ""
            zlCommFun.PressKey (vbKeyTab)
'            txtסԺ��.SetFocus
        End If
    End If
End Sub

Private Function CheckסԺ��(ByVal strסԺ�� As String, Optional lng��ҳID As Long = 0) As Boolean
    '-----------------------------------------------------------------------
    '�������סԺ���Ƿ����
    '����:bln����-�Ƿ�Ե�ǰ�Ĳ��˵�סԺ�Ų������ж�
    '����:����סԺ�ŷ���true,���򷵻�False
    '����:
    '-----------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnLimitUnit As Boolean, strUnitIDs As String, strWhere As String
'    ��ǰ����ID
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits & ","
        strWhere = " And instr([3],',' || ��ǰ����ID || ',')>0 "
    Else
        strWhere = ""
        strUnitIDs = ""
    End If
    
    On Error GoTo errHandle
    
    '����30031 by lesfeng 2010-05-19 ���� ��������
    If lng��ҳID = 0 Then
        strSQL = "select ����ID,סԺ��,max(��ҳID) as ��ҳID from ������ҳ where סԺ��=[1] " & strWhere & " Group By ����ID,סԺ��"
    Else
        strSQL = "select ����ID,סԺ��,��ҳID,decode(��Ժ����,Null,'��Ժ','��Ժ') as ״̬ from ������ҳ where סԺ��=[1] and ��ҳid=[2] " & strWhere
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strסԺ��, lng��ҳID, strUnitIDs)
    If Not rsTemp.EOF Then
        If rsTemp.RecordCount = 0 Then
            CheckסԺ�� = False
            rsTemp.Close
            Exit Function
        End If
        If lng��ҳID = 0 Then
            mlng����ID = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
            mlng��ҳID = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
            mstr״̬ = ""
        Else
            mlng����ID = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
            mlng��ҳID = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
            mstr״̬ = IIf(IsNull(rsTemp!״̬), "", rsTemp!״̬)
        End If
        If mlng����ID = 0 Then
            CheckסԺ�� = False
            rsTemp.Close
            Exit Function
        End If
        txtסԺ����.Text = mlng��ҳID
        txt״̬.Text = mstr״̬
        Call Check����(mlng����ID, mlng��ҳID)
    Else
        CheckסԺ�� = False
        chk����.Value = 0
        cmd����.Enabled = False
        chk����.Enabled = False
        rsTemp.Close
        Exit Function
    End If
    CheckסԺ�� = True
    Call GetסԺ����
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check����(ByVal lng����ID As Long, Optional lng��ҳID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    '����30031 by lesfeng 2010-05-19 ���� ��������
    chk����.Value = 0
    mint�������� = 0
    On Error GoTo errHandle
    strSQL = "select ��Ϣֵ from ������ҳ�ӱ� where ����id=[1] and ��ҳid = [2] and ��Ϣ�� = '��������'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        If rsTemp.RecordCount = 0 Then
            Check���� = False
            rsTemp.Close
            Exit Function
        End If
        mint�������� = Val(IIf(IsNull(rsTemp!��Ϣֵ), 0, rsTemp!��Ϣֵ))
        chk����.Value = mint��������
        cmd����.Enabled = False
        chk����.Enabled = True
    Else
        chk����.Enabled = True
        Check���� = False
        rsTemp.Close
        Exit Function
    End If
    Check���� = True

    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetסԺ����() As Boolean
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    If mlng��ҳID > 0 Then
        '��ȡ�������(��ҳID�е����ݿ��ܴ��ڲ�����)
        strSQL = "Select ��ҳID,�������� From ������ҳ Where ����ID=[1] And NVL(��ҳID,0)<>0 Order by ��ҳID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", mlng����ID)
        
        If rsTemp.RecordCount > 0 Then
            With cbo(cboName.��ҳID)
                .Clear
                For i = 1 To rsTemp.RecordCount
                    .AddItem "��" & rsTemp!��ҳID & "��סԺ" & IIf(Nvl(rsTemp!��������, 0) = 1, "(��������)", IIf(Nvl(rsTemp!��������, 0) = 2, "(סԺ����)", ""))
                    .ItemData(.NewIndex) = Val(rsTemp!��ҳID)
                    If Val(rsTemp!��ҳID) = mlng��ҳID Then .ListIndex = .NewIndex '�����ٵ���readcard
                    rsTemp.MoveNext
                Next
            End With
            cbo(cboName.��ҳID).Enabled = True
            cbo(cboName.��ҳID).Locked = False
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
