VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediHerbal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�в�ҩƷ�ֱ༭"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "ע����ҩƷ������2002��12��20�գ���2003��8��10��ͣ�á�"
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
      TabCaption(0)   =   "������Ϣ(&1)"
      TabPicture(0)   =   "frmMediHerbal.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl��λ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl��������"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl��ֵ"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl��Դ"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl����"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lbl����ְ��"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lblҽ��ְ��"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblҩƷ����"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl����"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblҩ�ⵥλChild"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblҩ���װ"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblҩ�ⵥλ"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl��ʶ��"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl����"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl����"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl�ۼ۵�λ"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "���쵥λ"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl���췧ֵ"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl���쵥λChild"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblComment"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl��ͬ��λ"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl˵��"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblҩ����λ"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl�ۼ۵�λChild"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl����ϵ��"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl���"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl����"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Lbl�ݴ�"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblҩ����λChild"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblҩ����װ"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl��ҩ����"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl��ѡ��"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblStationNo"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl�����Ա�"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "msf����"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cbo����ְ��"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt����"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt����"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cbo��λ"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt��������"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cbo��ֵ"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cbo��Դ"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cbo����"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtƴ��"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cboҽ��ְ��"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cboҩƷ����"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt���"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "chkԭ��ҩ"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtҩ���װ"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtҩ�ⵥλ"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt��ʶ��"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt����"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "chk����Ӧ��"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt���췧ֵ"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cbo���쵥λ"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmd�ο�"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt�ο�"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cmd��ͬ��λ"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt��ͬ��λ"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt˵��"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt�ۼ۵�λ"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtҩ����λ"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt����ϵ��"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txt���"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txt����"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cmd����"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cbo�ݴ�"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtҩ����װ"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cbo��ҩ����"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txt��ѡ��"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "cmbStationNo"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "chk���ҩ"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "cbo�����Ա�"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).ControlCount=   78
      TabCaption(1)   =   "ҩ����Ϣ(&2)"
      TabPicture(1)   =   "frmMediHerbal.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblҩ������"
      Tab(1).Control(1)=   "lbl���۵�λ(1)"
      Tab(1).Control(2)=   "lbl���۵�λ(0)"
      Tab(1).Control(3)=   "lblҩ�ۼ���"
      Tab(1).Control(4)=   "lbl�����"
      Tab(1).Control(5)=   "lbl����"
      Tab(1).Control(6)=   "lblPercent(0)"
      Tab(1).Control(7)=   "lbl�������"
      Tab(1).Control(8)=   "lbl��������"
      Tab(1).Control(9)=   "lbl��ǰ�ۼ�"
      Tab(1).Control(10)=   "lbl�������"
      Tab(1).Control(11)=   "lblָ������"
      Tab(1).Control(12)=   "lblָ������"
      Tab(1).Control(13)=   "lblָ���ۼ�"
      Tab(1).Control(14)=   "lbl�ӳ���"
      Tab(1).Control(15)=   "lbl�ɱ��۸�"
      Tab(1).Control(16)=   "lblPercent(1)"
      Tab(1).Control(17)=   "lbl����ѱ���"
      Tab(1).Control(18)=   "lbl�ɷ����"
      Tab(1).Control(19)=   "lbl��ֵ˰��"
      Tab(1).Control(20)=   "lblPercent(2)"
      Tab(1).Control(21)=   "fra��������"
      Tab(1).Control(22)=   "cboҩ������"
      Tab(1).Control(23)=   "chk���ηѱ�"
      Tab(1).Control(24)=   "txt�����"
      Tab(1).Control(25)=   "txt����"
      Tab(1).Control(26)=   "cbo�������"
      Tab(1).Control(27)=   "cbo��������"
      Tab(1).Control(28)=   "txt��ǰ�ۼ�"
      Tab(1).Control(29)=   "cbo�������"
      Tab(1).Control(30)=   "cboҩ�ۼ���"
      Tab(1).Control(31)=   "txtָ������"
      Tab(1).Control(32)=   "txtָ������"
      Tab(1).Control(33)=   "txtָ���ۼ�"
      Tab(1).Control(34)=   "txt�ӳ���"
      Tab(1).Control(35)=   "txt�ɱ��۸�"
      Tab(1).Control(36)=   "txt����ѱ���"
      Tab(1).Control(37)=   "cbo�ɷ����"
      Tab(1).Control(38)=   "txt��ֵ˰��"
      Tab(1).ControlCount=   39
      Begin VB.ComboBox cbo�����Ա� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   3660
         Width           =   1455
      End
      Begin VB.CheckBox chk���ҩ 
         Caption         =   "���ҩ(&K)"
         Height          =   210
         Left            =   6105
         TabIndex        =   124
         Top             =   4370
         Width           =   1305
      End
      Begin VB.TextBox txt��ֵ˰�� 
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
      Begin VB.TextBox txt��ѡ�� 
         Height          =   300
         Left            =   5550
         MaxLength       =   20
         TabIndex        =   117
         Top             =   5340
         Width           =   2400
      End
      Begin VB.ComboBox cbo��ҩ���� 
         Height          =   300
         Left            =   1230
         TabIndex        =   115
         Text            =   "cbo��ҩ����"
         Top             =   5340
         Width           =   3120
      End
      Begin VB.ComboBox cbo�ɷ���� 
         Height          =   300
         Left            =   -67560
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   495
         Width           =   1500
      End
      Begin VB.TextBox txtҩ����װ 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "30"
         Top             =   3720
         Width           =   510
      End
      Begin VB.ComboBox cbo�ݴ� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   1668
         Width           =   1455
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "&P"
         Height          =   285
         Left            =   5550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   435
         Width           =   285
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   108
         Top             =   450
         Width           =   4275
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   1230
         MaxLength       =   40
         TabIndex        =   15
         Top             =   2480
         Width           =   2250
      End
      Begin VB.TextBox txt����ϵ�� 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "1"
         Top             =   3300
         Width           =   525
      End
      Begin VB.TextBox txtҩ����λ 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "��"
         Top             =   3698
         Width           =   540
      End
      Begin VB.TextBox txt�ۼ۵�λ 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   18
         Text            =   "��"
         Top             =   3292
         Width           =   540
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   28
         Top             =   4980
         Width           =   3120
      End
      Begin VB.TextBox txt��ͬ��λ 
         Height          =   300
         Left            =   1230
         MaxLength       =   30
         TabIndex        =   27
         Top             =   4605
         Width           =   2820
      End
      Begin VB.CommandButton cmd��ͬ��λ 
         Caption         =   "��"
         Height          =   285
         Left            =   4080
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   4605
         Width           =   285
      End
      Begin VB.TextBox txt�ο� 
         Height          =   300
         Left            =   1230
         TabIndex        =   14
         Top             =   2074
         Width           =   4275
      End
      Begin VB.CommandButton cmd�ο� 
         Caption         =   "��"
         Height          =   285
         Left            =   5550
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   2074
         Width           =   285
      End
      Begin VB.ComboBox cbo���쵥λ 
         Height          =   300
         ItemData        =   "frmMediHerbal.frx":05C2
         Left            =   4620
         List            =   "frmMediHerbal.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2480
         Width           =   1215
      End
      Begin VB.TextBox txt���췧ֵ 
         Height          =   300
         Left            =   4620
         MaxLength       =   8
         TabIndex        =   25
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox txt����ѱ��� 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   75
         Top             =   1710
         Width           =   1350
      End
      Begin VB.TextBox txt�ɱ��۸� 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   69
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox txt�ӳ��� 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   67
         Text            =   "35.00"
         Top             =   2910
         Width           =   1665
      End
      Begin VB.CheckBox chk����Ӧ�� 
         Caption         =   "��ζʹ��(&Q)"
         Height          =   210
         Left            =   6105
         TabIndex        =   49
         Top             =   4080
         Width           =   1305
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4050
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1262
         Width           =   1470
      End
      Begin VB.TextBox txt��ʶ�� 
         Height          =   300
         Left            =   4050
         MaxLength       =   29
         TabIndex        =   5
         Top             =   856
         Width           =   1455
      End
      Begin VB.TextBox txtҩ�ⵥλ 
         Height          =   300
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   22
         Text            =   "ǧ��"
         Top             =   4110
         Width           =   540
      End
      Begin VB.TextBox txtҩ���װ 
         Height          =   300
         Left            =   2595
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "1000"
         Top             =   4110
         Width           =   510
      End
      Begin VB.TextBox txtָ���ۼ� 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   63
         Top             =   2100
         Width           =   1665
      End
      Begin VB.TextBox txtָ������ 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   54
         Top             =   900
         Width           =   1665
      End
      Begin VB.TextBox txtָ������ 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   65
         Text            =   "25.92593"
         Top             =   2505
         Width           =   1665
      End
      Begin VB.ComboBox cboҩ�ۼ��� 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2100
         Width           =   1635
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1305
         Width           =   1635
      End
      Begin VB.TextBox txt��ǰ�ۼ� 
         Height          =   300
         Left            =   -70350
         MaxLength       =   16
         TabIndex        =   71
         Top             =   900
         Width           =   1635
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         ItemData        =   "frmMediHerbal.frx":05C6
         Left            =   -70350
         List            =   "frmMediHerbal.frx":05C8
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2505
         Width           =   1635
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -70350
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   2910
         Width           =   1635
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   57
         Text            =   "100"
         Top             =   1305
         Width           =   1665
      End
      Begin VB.TextBox txt����� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73905
         MaxLength       =   16
         TabIndex        =   60
         Top             =   1695
         Width           =   1665
      End
      Begin VB.CheckBox chk���ηѱ� 
         Alignment       =   1  'Right Justify
         Caption         =   "���ηѱ�(&M)"
         Height          =   285
         Left            =   -68625
         TabIndex        =   84
         Top             =   960
         Width           =   1305
      End
      Begin VB.ComboBox cboҩ������ 
         Height          =   300
         Left            =   -73905
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   495
         Width           =   1665
      End
      Begin VB.Frame fra�������� 
         Caption         =   "��������"
         Height          =   1875
         Left            =   -68610
         TabIndex        =   85
         Top             =   1335
         Width           =   1845
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "ҩ����������"
            Enabled         =   0   'False
            Height          =   210
            Left            =   180
            TabIndex        =   87
            Top             =   720
            Width           =   1500
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "ҩ���������"
            Height          =   210
            Left            =   180
            TabIndex        =   86
            Top             =   345
            Width           =   1500
         End
      End
      Begin VB.CheckBox chkԭ��ҩ 
         Caption         =   "ԭ��ҩ(&M)"
         Height          =   210
         Left            =   7470
         TabIndex        =   50
         Top             =   4080
         Width           =   1155
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   4035
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1668
         Width           =   1110
      End
      Begin VB.ComboBox cboҩƷ���� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2074
         Width           =   1455
      End
      Begin VB.ComboBox cboҽ��ְ�� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2886
         Width           =   1455
      End
      Begin VB.TextBox txtƴ�� 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1668
         Width           =   1890
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   450
         Width           =   1455
      End
      Begin VB.ComboBox cbo��Դ 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1262
         Width           =   1455
      End
      Begin VB.ComboBox cbo��ֵ 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   856
         Width           =   1455
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   7170
         MaxLength       =   16
         TabIndex        =   48
         Text            =   "0"
         Top             =   3292
         Width           =   1455
      End
      Begin VB.ComboBox cbo��λ 
         Height          =   300
         Left            =   1230
         TabIndex        =   17
         Text            =   "g"
         Top             =   2886
         Width           =   780
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1230
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1262
         Width           =   1890
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1230
         MaxLength       =   13
         TabIndex        =   3
         Top             =   856
         Width           =   1890
      End
      Begin VB.ComboBox cbo����ְ�� 
         Height          =   300
         Left            =   7170
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2480
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit msf���� 
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
      Begin VB.Label lbl�����Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�(&S)"
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
      Begin VB.Label lbl��ֵ˰�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ֵ˰��(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   122
         Top             =   3420
         Width           =   990
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "վ����(&Z)"
         Height          =   180
         Left            =   135
         TabIndex        =   120
         Top             =   5760
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl��ѡ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��(&F)"
         Height          =   180
         Left            =   4680
         TabIndex        =   118
         Top             =   5400
         Width           =   810
      End
      Begin VB.Label lbl��ҩ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����(&H)"
         Height          =   180
         Left            =   135
         TabIndex        =   116
         Top             =   5400
         Width           =   990
      End
      Begin VB.Label lbl�ɷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʹ��(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -68610
         TabIndex        =   114
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lblҩ����װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2130
         TabIndex        =   113
         Top             =   3780
         Width           =   450
      End
      Begin VB.Label lblҩ����λChild 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   112
         Top             =   3780
         Width           =   180
      End
      Begin VB.Label Lbl�ݴ� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ�ݴ�(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   111
         Top             =   1725
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   109
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   107
         Top             =   2540
         Width           =   990
      End
      Begin VB.Label lbl����ϵ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2130
         TabIndex        =   106
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label lbl�ۼ۵�λChild 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   105
         Top             =   3345
         Width           =   180
      End
      Begin VB.Label lblҩ����λ 
         AutoSize        =   -1  'True
         Caption         =   "ҩ����λ(&I)"
         Height          =   180
         Left            =   165
         TabIndex        =   104
         Top             =   3758
         Width           =   990
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ˵��(&B)"
         Height          =   180
         Left            =   135
         TabIndex        =   103
         Top             =   5025
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(����д�ʵ���˵��������ʾ���á�����֢ҩƷ��)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4605
         TabIndex        =   102
         Top             =   5055
         Width           =   3960
      End
      Begin VB.Label lbl��ͬ��λ 
         AutoSize        =   -1  'True
         Caption         =   "��ͬ��λ(&C)"
         Height          =   180
         Left            =   135
         TabIndex        =   100
         Top             =   4650
         Width           =   990
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "(ָ���˺�ͬ��λ��ҩƷ��ֻ�ܰ���ͬ��λ��⡣)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4590
         TabIndex        =   99
         Top             =   4680
         Width           =   3945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�ο���Ŀ(&F)"
         Height          =   180
         Left            =   165
         TabIndex        =   97
         Top             =   2134
         Width           =   990
      End
      Begin VB.Label lbl���쵥λChild 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   5520
         TabIndex        =   95
         Top             =   2910
         Width           =   300
      End
      Begin VB.Label lbl���췧ֵ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���췧ֵ(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3585
         TabIndex        =   94
         Top             =   2910
         Width           =   990
      End
      Begin VB.Label ���쵥λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���쵥λ(&W)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3585
         TabIndex        =   33
         Top             =   2535
         Width           =   990
      End
      Begin VB.Label lbl����ѱ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ѱ���(&F)"
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
      Begin VB.Label lbl�ۼ۵�λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ۵�λ(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   3352
         Width           =   990
      End
      Begin VB.Label lbl�ɱ��۸� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɱ��۸�(&C)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   68
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lbl�ӳ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ӳ���(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74745
         TabIndex        =   66
         Top             =   2970
         Width           =   810
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(ƴ��)                (���)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   12
         Top             =   1710
         Width           =   2520
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3210
         TabIndex        =   8
         Top             =   1322
         Width           =   630
      End
      Begin VB.Label lbl��ʶ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3210
         TabIndex        =   4
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lblҩ�ⵥλ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ�ⵥλ(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   29
         Top             =   4170
         Width           =   990
      End
      Begin VB.Label lblҩ���װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1ǧ��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1950
         TabIndex        =   30
         Top             =   4170
         Width           =   630
      End
      Begin VB.Label lblҩ�ⵥλChild 
         AutoSize        =   -1  'True
         Caption         =   "g)"
         Height          =   180
         Left            =   3165
         TabIndex        =   31
         Top             =   4170
         Width           =   180
      End
      Begin VB.Label lblָ���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ָ���ۼ�(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   62
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɹ��޼�(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   53
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ָ������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   64
         Top             =   2565
         Width           =   990
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   72
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label lbl��ǰ�ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ�ۼ�(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   70
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   79
         Top             =   2565
         Width           =   990
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&S)"
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
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɹ�����(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   56
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label lbl����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74745
         TabIndex        =   59
         Top             =   1755
         Width           =   810
      End
      Begin VB.Label lblҩ�ۼ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ�ۼ���(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71385
         TabIndex        =   77
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lbl���۵�λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ/g"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   -72195
         TabIndex        =   55
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lbl���۵�λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ/g"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   -72195
         TabIndex        =   61
         Top             =   1755
         Width           =   360
      End
      Begin VB.Label lblҩ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ������(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74925
         TabIndex        =   51
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3615
         TabIndex        =   34
         Top             =   3210
         Width           =   990
      End
      Begin VB.Label LblҩƷ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   41
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label Lblҽ��ְ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��ְ��(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   45
         Top             =   2940
         Width           =   990
      End
      Begin VB.Label Lbl����ְ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ְ��(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   43
         Top             =   2535
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƽ���(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   1728
         Width           =   990
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   35
         Top             =   510
         Width           =   990
      End
      Begin VB.Label Lbl��Դ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ���(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   39
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Lbl��ֵ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ֵ����(&V)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   37
         Top             =   915
         Width           =   990
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6105
         TabIndex        =   47
         Top             =   3345
         Width           =   990
      End
      Begin VB.Label Lbl��λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   2946
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   6
         Top             =   1322
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����(&D)"
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
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   88
      Top             =   6300
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frmMediHerbal.frx":10FE
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7950
      TabIndex        =   89
      Top             =   6300
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf��ͬ��λ 
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
'˵����
'   1���༭״̬����Me.stbSpec.Tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�������
'---------------------------------------------------
Public lng����id As Long        '���༭��ҩƷ����ID���ϼ����򴫵ݽ���
Public lngҩ��ID As Long        '�޸ĺ͡���ѯʱ���ⲿ���򴫵ݽ���
Public strPrivs As String       '��ǰ�û��Ա������Ȩ�ޣ����ϼ�����򴫵ݽ���

Private lngҩƷID As Long       '�޸Ļ��ѯʱ���ݴ��ݽ���Ĳ���lngҩ��ID���ҵ�
Private mint������� As Integer     'ҩƷƷ�ֱ���Ĳ�������

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim ObjItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mstrMatch As String, strRefer As String '�ο�����
Dim mblnUsed As Boolean         '�Ƿ���ʹ��

Private mlng���볤�� As Long
Private mlng��񳤶� As Long
Private mlng���س��� As Long
Private mlng˵������ As Long
Private mlng���Ƴ��� As Long
Private mint���볤�� As Integer
Private mint��ѡ�볤�� As Integer

'�Ӳ�������ȡҩƷ�۸�С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��

Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer
Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    
    gstrSql = "Select A.����, A.����, A.���, A.˵��, A.����, B.����, A.��ѡ�� " & _
        " From �շ���ĿĿ¼ A, �շ���Ŀ���� B " & _
        " Where A.ID = B.�շ�ϸĿid And A.ID = 0 And B.���� = 1 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mlng��񳤶� = rsTmp.Fields("���").DefinedSize
    mlng���س��� = rsTmp.Fields("����").DefinedSize
    mlng˵������ = rsTmp.Fields("˵��").DefinedSize
    mlng���Ƴ��� = rsTmp.Fields("����").DefinedSize
    mint���볤�� = rsTmp.Fields("����").DefinedSize
    mint��ѡ�볤�� = rsTmp.Fields("��ѡ��").DefinedSize
    
    txt���.MaxLength = mlng��񳤶�
    txt����.MaxLength = mlng���س���
    txt˵��.MaxLength = mlng˵������
    txt��ѡ��.MaxLength = mint��ѡ�볤��
       
    gstrSql = "Select A.����, A.����, B.���� From ������ĿĿ¼ A, ������Ŀ���� B " & _
            " Where A.ID = B.������Ŀid And A.ID = 0 And B.���� = 1"
    Call zldatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
        
    If mlng���볤�� > rsTmp.Fields("����").DefinedSize Then
        mlng���볤�� = rsTmp.Fields("����").DefinedSize
    End If
    
    If mlng���Ƴ��� > rsTmp.Fields("����").DefinedSize Then
        mlng���Ƴ��� = rsTmp.Fields("����").DefinedSize
    End If
        
    If mint���볤�� > rsTmp.Fields("����").DefinedSize Then
        mint���볤�� = rsTmp.Fields("����").DefinedSize
    End If
        
    txt����.MaxLength = mlng���볤��
    txt����.MaxLength = mlng���Ƴ���
    txtƴ��.MaxLength = mint���볤��
    txt���.MaxLength = mint���볤��
    
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
    
    strSQL = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSQL = " Select ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=" & iAttr & " Order By ����"
    Else
        strSQLItem = " From ���Ʋο�Ŀ¼ A,���Ʋο����� B" & _
            " Where A.ID=B.�ο�Ŀ¼ID And A.����=" & iAttr & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%')"

        strSQL = " Select Distinct A.ID,A.����ID,A.����,A.����,A.˵�� " & strSQLItem & " Order By ����"
    End If
    Set SelectRefer = zldatabase.ShowSelect(Me, strSQL, 0, "�ο�", , , , , True)
End Function

Private Sub cbo����ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��λ_Change()
    Me.lbl�ۼ۵�λChild.Caption = Me.cbo��λ.Text & ")"
End Sub

Private Sub cbo��λ_Click()
    Me.lbl�ۼ۵�λChild.Caption = Me.cbo��λ.Text & ")"
End Sub

Private Sub cbo��λ_GotFocus()
    Me.cbo��λ.SelStart = 0: Me.cbo��λ.SelLength = 100
End Sub

Private Sub cbo��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��ҩ����_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���쵥λ_Click()
    Select Case cbo���쵥λ.ListIndex
    Case 0
        lbl���쵥λChild.Caption = txt�ۼ۵�λ.Text & ")"
    Case 1
        lbl���쵥λChild.Caption = txtҩ����λ.Text & ")"
    Case 2
        lbl���쵥λChild.Caption = txtҩ�ⵥλ.Text & ")"
    End Select
End Sub

Private Sub cbo���쵥λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ݴ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboҩ�ۼ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboҩ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboҩƷ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboҽ��ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk����Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chk���ηѱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chkҩ��_Click()
    Dim blnEnable As Boolean
    
    '��ҩ�������ǰ���£����ҩ��û�п�棬����������Ƿ����
    gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
             " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '%ҩ��' Or B.�������� Like '%�Ƽ���')"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
    
    With rsTemp
        blnEnable = True
        If .Fields(0).Value <> 0 Then
            blnEnable = False
        End If
    End With
    If Me.chkҩ��.Value = 0 Then
        Me.chkҩ��.Value = 0: Me.chkҩ��.Enabled = False
    Else
        Me.chkҩ��.Enabled = True
    End If
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chkԭ��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.stbSpec.Tab = 1
        If Me.cboҩ������.Enabled Then
            Me.cboҩ������.SetFocus
        Else
            Me.txtָ������.SetFocus
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
    Dim dblָ���ۼ� As Double, dbl��ǰ�ۼ� As Double, dbl�ɱ��۸� As Double
    Dim rsData As ADODB.Recordset
    Dim blnPackerReturn As Boolean
    
    '�༭���ݼ��
    If Trim(Me.txt����.Text) = "" Then MsgBox "��������룡", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt����.Text, vbFromUnicode)) > mlng���볤�� Then MsgBox "���볬��(���" & mlng���볤�� & "���ַ�)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mlng���Ƴ��� Then MsgBox "���Ƴ��������" & mlng���Ƴ��� & "���ַ���" & Int(mlng���Ƴ��� / 2) & "�����֣���", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.cbo��λ.Text) = "" Then MsgBox "�����������λ��", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.cbo��λ.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.cbo��λ.Text), vbFromUnicode)) > 6 Then MsgBox "������λ�ĳ��������6���ַ���3�����֣���", vbInformation, gstrSysName: stbSpec.Tab = 0: Me.cbo��λ.SetFocus: Exit Sub
    
    If LenB(StrConv(Me.txt��ѡ��.Text, vbFromUnicode)) > mint��ѡ�볤�� Then MsgBox "��ѡ�볬��(���" & mint��ѡ�볤�� & "���ַ�)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt��ѡ��.SetFocus: Exit Sub
    
    If Trim(Me.txt�ۼ۵�λ.Text) = "" Then MsgBox "�������ۼ۵�λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�ۼ۵�λ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt�ۼ۵�λ.Text, vbFromUnicode)) > 8 Then MsgBox "�ۼ۵�λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�ۼ۵�λ.SetFocus: Exit Sub
    If Val(Me.txt����ϵ��.Text) = 0 Then MsgBox "����ϵ������(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����ϵ��.SetFocus: Exit Sub
    If Val(Me.txt����ϵ��.Text) >= 100000 Then MsgBox "����ϵ���������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����ϵ��.SetFocus: Exit Sub
    
    If Trim(Me.txtҩ����λ.Text) = "" Then MsgBox "������ҩ����λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ����λ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txtҩ����λ.Text, vbFromUnicode)) > 8 Then MsgBox "ҩ����λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ����λ.SetFocus: Exit Sub
    If Val(Me.txtҩ����װ.Text) = 0 Then MsgBox "ҩ����װ����(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ����װ.SetFocus: Exit Sub
    If Val(Me.txtҩ����װ.Text) >= 100000 Then MsgBox "ҩ����װ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ����װ.SetFocus: Exit Sub
    
    strTemp = IIf(glngSys \ 100 <> 8, "ҩ��", "�ɹ�")
    If Trim(Me.txtҩ�ⵥλ.Text) = "" Then MsgBox "������" & strTemp & "��λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ�ⵥλ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txtҩ�ⵥλ.Text, vbFromUnicode)) > 8 Then MsgBox strTemp & "��λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ�ⵥλ.SetFocus: Exit Sub
    If Val(Me.txtҩ���װ.Text) = 0 Then MsgBox strTemp & "��װ����(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ���װ.SetFocus: Exit Sub
    If Val(Me.txtҩ���װ.Text) >= 100000 Then MsgBox strTemp & "��װ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ���װ.SetFocus: Exit Sub
    
    If Val(Me.txt���췧ֵ.Text) < 0 Then MsgBox strTemp & "���췧ֵ����С���㣡", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���췧ֵ.SetFocus: Exit Sub
    If Val(Me.txt���췧ֵ.Text) >= 100000 Then MsgBox strTemp & "���췧ֵ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���췧ֵ.SetFocus: Exit Sub
    
    If Val(Me.txtָ������.Text) = 0 And mblnUsed = True Then
        MsgBox "������ָ�����ۣ�", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus: Exit Sub
    End If
    If Val(Me.txtָ������.Text) > 1000000 Then
        MsgBox "ָ�����۳������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus: Exit Sub
    End If
    If Val(Me.txtָ���ۼ�.Text) = 0 And mblnUsed = True Then
        MsgBox "������ָ���ۼۣ�", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ���ۼ�.Enabled Then Me.txtָ���ۼ�.SetFocus: Exit Sub
    End If
    If Val(Me.txtָ���ۼ�.Text) > 1000000 Then
        MsgBox "ָ���ۼ۳������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ���ۼ�.Enabled Then Me.txtָ���ۼ�.SetFocus: Exit Sub
    End If
'    If Val(Me.txtָ������.Text) = 0 Then
'        MsgBox "������ָ�����ʣ�", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
'        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus: Exit Sub
'    End If
    If Val(Me.txtָ������.Text) > 100 Then
        MsgBox "ָ�����ʳ������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus: Exit Sub
    End If
    If Val(Me.txt����.Text) = 0 Then MsgBox "��������ʣ�", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����.SetFocus: Exit Sub
    If Val(Me.txt����.Text) > 100 Then MsgBox "���ʳ������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����.SetFocus: Exit Sub
    If Val(Me.txt����ѱ���.Text) < 0 Then MsgBox "����ѱ�������С���㣡", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����ѱ���.SetFocus: Exit Sub
    If Val(Me.txt����ѱ���.Text) > 100 Then MsgBox "����ѱ����������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����ѱ���.SetFocus: Exit Sub
    
    If Val(Me.txt��ֵ˰��.Text) < 0 Then MsgBox "��ֵ˰�ʱ�������С���㣡", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt��ֵ˰��.SetFocus: Exit Sub
    If Val(Me.txt��ֵ˰��.Text) > 100 Then MsgBox "��ֵ˰�ʱ����������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt��ֵ˰��.SetFocus: Exit Sub
    
    If Me.cboҩ������.ItemData(cboҩ������.ListIndex) = 0 Then
'        If Val(Me.txt��ǰ�ۼ�.Text) = 0 And Me.txt��ǰ�ۼ�.Enabled = True Then
'            MsgBox "�����뵱ǰ�ۼۣ�", vbInformation, gstrSysName
'            Me.stbSpec.Tab = 1
'            If Me.txt��ǰ�ۼ�.Enabled Then Me.txt��ǰ�ۼ�.SetFocus
'            Exit Sub
'        End If
        If Val(Me.txt��ǰ�ۼ�.Text) > Val(Me.txtָ���ۼ�.Text) Then
            If MsgBox("�ۼ۸���ָ�����ۼۡ�" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Me.stbSpec.Tab = 1
                If Me.txt��ǰ�ۼ�.Enabled Then Me.txt��ǰ�ۼ�.SetFocus
                Exit Sub
            End If
        End If
        If Val(Me.txt��ǰ�ۼ�.Text) > 1000000 Then
            MsgBox "��ǰ�ۼ۳������ֵ��", vbInformation, gstrSysName
            Me.stbSpec.Tab = 1
            If Me.txt��ǰ�ۼ�.Enabled Then Me.txt��ǰ�ۼ�.SetFocus
            Exit Sub
        End If
    End If
    
    If LenB(StrConv(Me.txt����.Text, vbFromUnicode)) > 60 Then MsgBox "���س���(���60���ַ���30������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    
    '�������
    strTemp = ";" & Trim(Me.txt����.Text)
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "���������ظ����������ƣ���", vbInformation, gstrSysName
                    stbSpec.Tab = 0: .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    '���ݱ���
    If Me.stbSpec.Tag = "����" Then
        lngҩ��ID = zldatabase.GetNextId("������ĿĿ¼")
        If zlClinicCodeRepeat(Trim(Me.txt����.Text)) = True Then Exit Sub
        If zlExseCodeRepeat(Trim(Me.txt����.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(Trim(Me.txt����.Text), lngҩ��ID) = True Then Exit Sub
    End If
    If Not CheckRequest Then Exit Sub
    
    gstrSql = Me.txt����.Tag & "," & lngҩ��ID & ",'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt��ʶ��.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txtƴ��.Text) & "','" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.cbo��λ.Text) & "','" & Trim(Me.txt���.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt�ۼ۵�λ.Text) & "'," & Val(Me.txt����ϵ��.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txtҩ����λ.Text) & "'," & Val(Me.txtҩ����װ.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txtҩ����λ.Text) & "'," & Val(Me.txtҩ����װ.Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txtҩ�ⵥλ.Text) & "'," & Val(Me.txtҩ���װ.Text)
    gstrSql = gstrSql & "," & IIf(cbo���쵥λ.ListIndex = 0, 1, IIf(cbo���쵥λ.ListIndex = 2, 4, 3)) '���쵥λ��1-���۵�λ;2-סԺ��λ;3-ҩ����λ;4-ҩ�ⵥλ�����в�ҩֻ��1,4
    gstrSql = gstrSql & "," & Val(txt���췧ֵ.Tag)           'ʼ�������۵�λ����
    gstrSql = gstrSql & ",'" & Mid(Me.cbo����.Text, InStr(1, Me.cbo����.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��ֵ.Text, InStr(1, Me.cbo��ֵ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��Դ.Text, InStr(1, Me.cbo��Դ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo�ݴ�.Text, InStr(1, Me.cbo�ݴ�.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Left(Me.cboҩƷ����.Text, 1) & ",'" & Left(Me.cbo����ְ��.Text, 1) & Left(Me.cboҽ��ְ��.Text, 1) & "'"
    gstrSql = gstrSql & "," & Val(Trim(Me.txt��������.Text)) & "," & Me.chk����Ӧ��.Value & "," & Me.chkԭ��ҩ.Value
    
    gstrSql = gstrSql & "," & Me.cboҩ������.ItemData(Me.cboҩ������.ListIndex)
    If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
        dblָ���ۼ� = Round(Val(txtָ���ۼ�.Text) / Val(txtҩ���װ.Text), mintSalePriceDigit)
        dbl��ǰ�ۼ� = Round(Val(txt��ǰ�ۼ�.Text) / Val(txtҩ���װ.Text), mintSalePriceDigit)
        dbl�ɱ��۸� = Round(Val(txt�ɱ��۸�.Text) / Val(txtҩ���װ.Text), mintSaleCostDigit)
        gstrSql = gstrSql & "," & Round(Val(Me.txtָ������.Text) / Val(Me.txtҩ���װ), mintSaleCostDigit)
    Else
        dbl��ǰ�ۼ� = Round(Val(txt��ǰ�ۼ�.Text), mintPriceDigit)
        dblָ���ۼ� = Round(Val(txtָ���ۼ�.Text), mintPriceDigit)
        dbl�ɱ��۸� = Round(Val(txt�ɱ��۸�.Text), mintCostDigit)
        gstrSql = gstrSql & "," & Val(Me.txtָ������.Text)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt����.Text) & "," & dblָ���ۼ� & "," & Val(Me.txtָ������.Text) & "," & Val(Me.txt����ѱ���.Text)
    gstrSql = gstrSql & ",'" & Mid(Me.cboҩ�ۼ���.Text, InStr(1, Me.cboҩ�ۼ���.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex) & "," & Me.chk���ηѱ�.Value
    gstrSql = gstrSql & "," & Me.chkҩ�� & "," & Me.chkҩ��
    gstrSql = gstrSql & "," & IIf(Val(Me.txt�ο�.Tag) = 0, "NULL", Val(Me.txt�ο�.Tag))
    strTemp = ""
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    '������������
    If LenB(strTemp) > 4000 Then
        msf����.SetFocus
        MsgBox "�����ַ���̫��������ٱ����������߱������ȡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    gstrSql = gstrSql & ",'" & strTemp & "'"
    gstrSql = gstrSql & "," & dbl�ɱ��۸�
    gstrSql = gstrSql & "," & dbl��ǰ�ۼ�
    gstrSql = gstrSql & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex)
    gstrSql = gstrSql & "," & IIf(Split(Me.txt��ͬ��λ.Tag, "|")(0) = "", "NULL", Split(Me.txt��ͬ��λ.Tag, "|")(0))
    gstrSql = gstrSql & ",'" & Me.txt˵��.Text & "'"
    gstrSql = gstrSql & "," & Me.cbo�ɷ����.ItemData(Me.cbo�ɷ����.ListIndex)
    gstrSql = gstrSql & ",'" & cbo��ҩ����.Text
    gstrSql = gstrSql & "','" & txt��ѡ��.Text & "'"
    gstrSql = gstrSql & "," & Val(Me.txt��ֵ˰��.Text)
    gstrSql = gstrSql & "," & Me.chk���ҩ.Value
    gstrSql = gstrSql & "," & Left(Me.cbo�����Ա�.Text, 1)
    gstrSql = gstrSql & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", cmbStationNo.Text)
    
    If Me.stbSpec.Tag = "����" Then
        gstrSql = "zl_��ҩҩƷ_INSERT(" & gstrSql & ")"
    Else
        gstrSql = "zl_��ҩҩƷ_Update(" & gstrSql & ")"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    '�������ݵ�ҩƷ�ְ��ӿ����ݿ�
    If gblnStartPacker = True And gblnPackerConnect = True Then
        gstrSql = "Select ҩƷid From ҩƷ��� Where ҩ��id = [1] "
        Set rsData = zldatabase.OpenSQLRecord(gstrSql, "ȡҩƷID", lngҩ��ID)
        If Not rsData.EOF Then
            blnPackerReturn = gobjPacker.TranDrugSingle(gcnOracle, Val(rsData!ҩƷID))
        End If
    End If
    
    If Me.stbSpec.Tag = "����" And Val(zldatabase.GetPara("Ʒ������ģʽ", glngSys, 1023, 0)) = 1 Then
        Call frmMediLists.zlRefRecords(lngҩ��ID)
        lngҩ��ID = 0
        Call Form_Activate
        Me.txt����.SetFocus
    Else
        Unload Me
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    ShowHelp App.ProductName, Me.hWnd, "frmMediItem", Int((glngSys) / 100)
End Sub

Private Sub cmd�ο�_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
    End If
End Sub

Private Sub cmd����_Click()
    With Me.tvwClass
        .Left = Me.txt����.Left + Me.stbSpec.Left
        .Top = Me.txt����.Top + Me.txt����.Height + Me.stbSpec.Top
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub cmd��ͬ��λ_Click()
    With rsTemp
        gstrSql = "Select ����,����,����,id" & _
        " From ��Ӧ��" & _
        " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By ���� "
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .EOF Then
            MsgBox "���ʼ����Ӧ�̣��ֵ������", vbInformation, gstrSysName
            Me.txt��ͬ��λ.Tag = "|": Me.txt��ͬ��λ.SetFocus: Exit Sub
        End If
        With Me.msf��ͬ��λ
            .Left = Me.stbSpec.Left + Me.txt��ͬ��λ.Left
            .Top = Me.stbSpec.Top + Me.txt��ͬ��λ.Top - Me.msf��ͬ��λ.Height
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
    '�������ݼ��
    If Me.cbo����.ListCount = 0 Then
        strMsg = "�޶���������ݣ�����ϵϵͳ����Ա"
        blnExit = True
    End If
    If Me.cbo��ֵ.ListCount = 0 And Not blnExit Then
        strMsg = "�޼�ֵ�������ݣ�����ϵϵͳ����Ա"
        blnExit = True
    End If
    If Me.cbo��Դ.ListCount = 0 And Not blnExit Then
        strMsg = "�޻�Դ�������ݣ�����ϵϵͳ����Ա"
        blnExit = True
    End If
    If Me.cbo�ݴ�.ListCount = 0 And Not blnExit Then
        strMsg = "����ҩ�ݴ����ݣ�����ϵϵͳ����Ա"
        blnExit = True
    End If
    If Me.cbo��������.ListCount = 0 And Not blnExit Then
        strMsg = "δ��������ҩƷ��ҽ�����ͣ��ֵ����"
        blnExit = True
    End If
    If Me.cbo�������.ListCount = 0 And Not blnExit Then
        strMsg = "δ������ϸ��������Ŀ��"
        blnExit = True
    End If
    If Me.stbSpec.Tag = "����" And Val(Me.lbl�������.Tag) = 0 Then
        strMsg = "û�����á��в�ҩ����Ӧ��������Ŀ���������ã���"
        blnExit = True
    End If
    If Me.cboҩ�ۼ���.ListCount = 0 And Not blnExit Then
        strMsg = "δ����ҩ�۹������ֵ������"
        blnExit = True
    End If
    If blnExit Then
        MsgBox strMsg, vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '----------����װ��-------------------------------------
    Me.tvwClass.Nodes("_" & lng����id).Selected = True
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    
    lngҩƷID = 0
    
    gstrSql = "select I.����ID,S.ҩƷID,I.����,I.����,I.���㵥λ ������λ,S.ҩ�ⵥλ,S.����ϵ��,S.ҩ���װ,C.����,S.��ʶ��," & _
            "        T.�������,T.��Դ���,T.��ֵ����,T.��ҩ�ݴ�,S.���쵥λ,S.���췧ֵ," & _
            "        nvl(T.ҩƷ����,0) as ҩƷ����,nvl(T.����ְ��,'00') as ����ְ��,nvl(T.��������,0) as ��������," & _
            "        nvl(T.�Ƿ�ԭ��,0) as �Ƿ�ԭ��,nvl(I.����Ӧ��,0) as ����Ӧ��," & _
            "        C.�Ƿ���,S.ָ��������,S.����,S.ָ�����ۼ�,S.ָ�������,S.����ѱ���,S.�ɱ���," & _
            "        S.ҩ�ۼ���,C.��������,C.�������,C.���ηѱ�,S.�ɷ����,S.��ҩ����," & _
            "        S.ҩ�����,S.ҩ������,S.���ﵥλ,S.�����װ,C.���,C.���㵥λ,C.��ѡ��," & _
            "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,B.���� as �ο�����,I.�ο�Ŀ¼id,S.��ͬ��λid,G.���� ��ͬ��λ,C.˵��,C.վ��,S.��ֵ˰��,S.���,Nvl(I.�����Ա�,0) As �����Ա� " & _
            " from ������ĿĿ¼ I,ҩƷ���� T,ҩƷ��� S,�շ���ĿĿ¼ C,���Ʋο�Ŀ¼ B,(Select Id,���� From ��Ӧ�� Where ĩ�� = 1 And substr(����,1,1) = '1' And " & _
            " ����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) G " & _
            " where I.ID=T.ҩ��ID and T.ҩ��ID=S.ҩ��ID and S.ҩƷID=C.ID and I.ID=[1] and I.�ο�Ŀ¼id=B.id(+) and G.id(+)=S.��ͬ��λid "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��ID)
    
    With rsTemp
        If .RecordCount > 0 Then
            lngҩƷID = !ҩƷID
            Me.lblFoot.Caption = "ע����ҩƷ������" & Format(!����ʱ��, "YYYY-MM-DD")
            If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                Me.lblFoot.Caption = Me.lblFoot.Caption & "����" & Format(!����ʱ��, "YYYY-MM-DD") & "ͣ�á�"
            End If
            Me.tvwClass.Nodes("_" & !����id).Selected = True
            Me.txt����.Text = Me.tvwClass.SelectedItem.Text
            Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            Me.txt����.Text = !����
            Me.txt����.Text = !����
            Me.txt����.Text = IIf(IsNull(!����), "", !����)
            Me.txt��ʶ��.Text = IIf(IsNull(!��ʶ��), "", !��ʶ��)
            Me.txt��ѡ��.Text = IIf(IsNull(!��ѡ��), "", !��ѡ��)
            
            Me.txt��ͬ��λ.Text = IIf(IsNull(!��ͬ��λ), "", !��ͬ��λ)
            Me.txt��ͬ��λ.Tag = IIf(IsNull(!��ͬ��λid), "|", !��ͬ��λid & "|" & !��ͬ��λ)
            
            Me.txt˵��.Text = IIf(IsNull(!˵��), "", !˵��)
            
            Me.txt���.Text = IIf(IsNull(!���), "", !���)
            Me.cbo���쵥λ.ListIndex = IIf(Nvl(!���쵥λ, 1) = 1, 0, IIf(Nvl(!���쵥λ, 4) = 4, 2, 1))
            Me.txt���췧ֵ.Text = Format(Nvl(!���췧ֵ, 0), "#0.00;-#0.00; ;")          'ȱʡ�����۵�λ��ʾ
            Me.cbo��λ.Text = IIf(IsNull(!������λ), "", !������λ)
            Me.txt�ۼ۵�λ.Text = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            Me.lblҩ�ⵥλChild.Caption = Me.cbo��λ.Text & ")"
            Me.txtҩ�ⵥλ.Text = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
            Me.lblҩ���װ.Caption = "(1" & Me.txtҩ�ⵥλ.Text & "="
            Me.txt�ο�.Text = Nvl(!�ο�����)
            Me.txt�ο�.Tag = Nvl(!�ο�Ŀ¼ID)
            strRefer = Me.txt�ο�.Text
            Me.txt����ϵ��.Text = IIf(IsNull(!����ϵ��), 1, !����ϵ��)
            Me.txtҩ���װ.Text = IIf(IsNull(!ҩ���װ), 1, !ҩ���װ)
            
            Me.txtҩ����λ.Text = IIf(IsNull(!���ﵥλ), "", !���ﵥλ)
            Me.txtҩ����װ.Text = IIf(IsNull(!�����װ), 1, !�����װ)
            
            Me.lblҩ����λChild.Caption = Me.txt�ۼ۵�λ & ")"
            Me.lblҩ�ⵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
            
            Me.cbo��ҩ����.Text = Nvl(!��ҩ����)
            Me.cbo�����Ա�.ListIndex = !�����Ա�
            
            SetStationNo IIf(IsNull(!վ��), "", !վ��)
            
            Select Case IIf(IsNull(!�ɷ����), 0, !�ɷ����)
            Case 0, 1
                Me.cbo�ɷ����.ListIndex = IIf(IsNull(!�ɷ����), 0, !�ɷ����)
            Case Else
                Me.cbo�ɷ����.ListIndex = 0
            End Select
            
            '����ǰ�ҩ�ⵥλ�������췧ֵ����ҩ�ⵥλ��ʾ
            If Me.cbo���쵥λ.ListIndex = 1 Then
                Me.txt���췧ֵ.Text = Format(Nvl(!���췧ֵ, 0) / Val(txtҩ����װ.Text), "#0.00;-#0.00; ;")
            ElseIf Me.cbo���쵥λ.ListIndex = 2 Then
                Me.txt���췧ֵ.Text = Format(Nvl(!���췧ֵ, 0) / Val(txtҩ���װ.Text), "#0.00;-#0.00; ;")
            End If
            
            For intCount = 0 To Me.cbo����.ListCount - 1
                If Mid(Me.cbo����.List(intCount), InStr(1, Me.cbo����.List(intCount), "-") + 1) = IIf(IsNull(!�������), "", !�������) Then
                    Me.cbo����.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo��ֵ.ListCount - 1
                If Mid(Me.cbo��ֵ.List(intCount), InStr(1, Me.cbo��ֵ.List(intCount), "-") + 1) = IIf(IsNull(!��ֵ����), "", !��ֵ����) Then
                    Me.cbo��ֵ.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo��Դ.ListCount - 1
                If Mid(Me.cbo��Դ.List(intCount), InStr(1, Me.cbo��Դ.List(intCount), "-") + 1) = IIf(IsNull(!��Դ���), "", !��Դ���) Then
                    Me.cbo��Դ.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo�ݴ�.ListCount - 1
                If Mid(Me.cbo�ݴ�.List(intCount), InStr(1, Me.cbo�ݴ�.List(intCount), "-") + 1) = IIf(IsNull(!��ҩ�ݴ�), "", !��ҩ�ݴ�) Then
                    Me.cbo�ݴ�.ListIndex = intCount: Exit For
                End If
            Next
            Me.cboҩƷ����.ListIndex = !ҩƷ����
            Me.cbo����ְ��.ListIndex = IIf(CInt(Left(Format(!����ְ��, "00"), 1)) <> 9, CInt(Left(Format(!����ְ��, "00"), 1)), Me.cbo����ְ��.ListCount - 1)
            Me.cboҽ��ְ��.ListIndex = IIf(CInt(Right(Format(!����ְ��, "00"), 1)) <> 9, CInt(Right(Format(!����ְ��, "00"), 1)), Me.cboҽ��ְ��.ListCount - 1)
            Me.chk����Ӧ��.Value = IIf(!����Ӧ�� = 0, 0, 1)
            Me.chkԭ��ҩ.Value = IIf(!�Ƿ�ԭ�� = 0, 0, 1)
            
            Me.cboҩ������.ListIndex = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
            Me.txt����.Text = IIf(IsNull(!����), 100, !����)
            If Val(Me.lbl���۵�λ(0).Tag) <> 0 = True Then
                Me.txtָ������.Text = FormatEx(IIf(IsNull(!ָ��������), 0, !ָ��������) * Me.txtҩ���װ.Text, mintCostDigit)
                Me.txtָ���ۼ�.Text = FormatEx(IIf(IsNull(!ָ�����ۼ�), 0, !ָ�����ۼ�) * Me.txtҩ���װ.Text, mintPriceDigit)
                Me.txt�ɱ��۸�.Text = FormatEx(IIf(IsNull(!�ɱ���), 0, !�ɱ���) * Me.txtҩ���װ.Text, mintCostDigit)
            Else
                Me.txtָ������.Text = FormatEx(IIf(IsNull(!ָ��������), 0, !ָ��������), mintCostDigit)
                Me.txtָ���ۼ�.Text = FormatEx(IIf(IsNull(!ָ�����ۼ�), 0, !ָ�����ۼ�), mintPriceDigit)
                Me.txt�ɱ��۸�.Text = FormatEx(IIf(IsNull(!�ɱ���), 0, !�ɱ���), mintCostDigit)
            End If
            Me.txt����� = FormatEx((Me.txtָ������.Text) * Me.txt����.Text / 100, mintPriceDigit)
            Me.txtָ������.Text = Format(IIf(IsNull(!ָ�������), 0, !ָ�������), "0.00000")
            Me.txt����ѱ���.Text = Format(Nvl(!����ѱ���, 0), "#0.00")
            Me.txt��ֵ˰��.Text = Format(Nvl(!��ֵ˰��, 0), "0.00")
            Me.chk���ҩ.Value = IIf(!��� = 0, 0, 1)
            '����ָ���ӳ���
            Dim cur�۸� As Double
            cur�۸� = Val(txtָ������.Text)
            If cur�۸� < 100 Then
                Call Calc(cur�۸�, True)
                Me.txt�ӳ���.Text = Format(cur�۸�, "0.00")
            End If
            
            For intCount = 0 To Me.cboҩ�ۼ���.ListCount - 1
                If Mid(Me.cboҩ�ۼ���.List(intCount), InStr(1, Me.cboҩ�ۼ���.List(intCount), "-") + 1) = IIf(IsNull(!ҩ�ۼ���), "", !ҩ�ۼ���) Then
                    Me.cboҩ�ۼ���.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(intCount), 4) = IIf(IsNull(!��������), "", !��������) Then
                    Me.cbo��������.ListIndex = intCount: Exit For
                End If
            Next
            Me.cbo�������.ListIndex = IIf(IsNull(!�������), 0, !�������)
            Me.chk���ηѱ�.Value = IIf(IsNull(!���ηѱ�), 0, !���ηѱ�)
            
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFoot.Caption = "ע����ҩƷ��" & Format(!����ʱ��, "YYYY��MM��DD��") & "������" & Format(!����ʱ��, "YYYY��MM��DD��") & "ͣ��"
            Else
                Me.lblFoot.Caption = ""
            End If
            
            Me.chkҩ��.Tag = IIf(IsNull(!ҩ������), 0, !ҩ������)
            Me.chkҩ��.Value = IIf(IsNull(!ҩ�����), 0, Abs(!ҩ�����))
            If Me.chkҩ��.Value = 0 Then
                Me.chkҩ��.Enabled = False: Me.chkҩ��.Value = 0
            Else
                Me.chkҩ��.Enabled = True: Me.chkҩ��.Value = Me.chkҩ��.Tag
            End If
        End If
        If Trim(Me.txt��ͬ��λ.Tag) = "" Then
            Me.txt��ͬ��λ.Tag = "|"
        End If
        If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
            Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
            Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
        Else
            Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
            Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
        End If
    End With
    
    If Me.stbSpec.Tag = "����" Then
        '����ʱ��������ȡ����
        Me.txt����.Text = "": Me.txt����.Text = "": Me.txt����.Text = "": Me.lblFoot.Caption = ""
        lngҩ��ID = 0
        Me.txt�ο� = "": Me.txt�ο�.Tag = "": strRefer = ""
        If mint������� = 0 Then
            gstrSql = "select nvl(max(����),'0000000') as ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ��� = '7'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): rsTemp.Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
            Me.txt����.Text = zlcommfun.IncStr(rsTemp!����)
        Else
            strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
            gstrSql = "select nvl(max(����),'') as ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ��� = '7' and ���� like [1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, "7" & strTemp & "%")
            
            Err = 0: On Error Resume Next
            strTemp = "7" & strTemp
            If Nvl(rsTemp!����) = "" Then
                Me.txt����.Text = strTemp & "01"
            Else
                Me.txt����.Text = zlcommfun.IncStr(rsTemp!����)
            End If
        End If
    Else
        '��������
        gstrSql = "select ����,����,����,���� from ������Ŀ���� where ���� in (1,2) and ������ĿID=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��ID)
        
        Do While Not rsTemp.EOF
            If rsTemp!���� = 1 And rsTemp!���� = 1 Then Me.txtƴ��.Text = rsTemp!����
            If rsTemp!���� = 1 And rsTemp!���� = 2 Then Me.txt���.Text = rsTemp!����
            rsTemp.MoveNext
        Loop
        '��������
        gstrSql = "select N.����,P.���� as ƴ��,W.���� as ���" & _
                " from (select distinct ���� from ������Ŀ���� where ������ĿID=[1] and ����=9) N," & _
                "      (select ����,���� from ������Ŀ���� where ������ĿID=" & lngҩ��ID & " and ����=9 and ����=1) P," & _
                "      (select ����,���� from ������Ŀ���� where ������ĿID=" & lngҩ��ID & " and ����=9 and ����=2) W" & _
                " where N.����=P.����(+) and N.����=W.����(+)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��ID)
        
        With rsTemp
            Do While Not .EOF
                If Me.msf����.Rows - 1 < .AbsolutePosition Then Me.msf����.Rows = Me.msf����.Rows + 1
                Me.msf����.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                Me.msf����.TextMatrix(.AbsolutePosition, 1) = !����
                Me.msf����.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!ƴ��), "", !ƴ��)
                Me.msf����.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!���), "", !���)
                .MoveNext
            Loop
        End With
        
        '��ȡ��ʾ��ǰ�ۼ�
        If Me.cboҩ������.ListIndex <> 0 Then
            Me.cboҩ������.Enabled = False
            gstrSql = "select Decode(K.�������,0,P.�ּ�,K.�����/Nvl(K.�������,1)) as �ּ�,P.������Ŀid" & _
                    " from �շѼ�Ŀ P," & _
                    "     (Select nvl(Sum(ʵ�ʽ��),0) as �����,nvl(Sum(ʵ������),0) as �������" & _
                    "      From ҩƷ��� Where ҩƷID=[1]) K" & _
                    " where P.�շ�ϸĿid=[1] and (P.��ֹ���� is null or Sysdate Between P.ִ������ And P.��ֹ����)"
        Else
            '��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
            gstrSql = "select P.�ּ�,P.������Ŀid" & _
                    " from �շѼ�Ŀ P" & _
                    " where P.�շ�ϸĿid=[1] and (P.��ֹ���� is null or Sysdate Between P.ִ������ And P.��ֹ����)"
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        With rsTemp
            If .RecordCount > 0 Then
                If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
                    Me.txt��ǰ�ۼ�.Text = FormatEx(!�ּ� * Val(Me.txtҩ���װ.Text), mintPriceDigit)
                Else
                    Me.txt��ǰ�ۼ�.Text = FormatEx(!�ּ�, mintPriceDigit)
                End If
        
                For intCount = 0 To Me.cbo�������.ListCount - 1
                    If Me.cbo�������.ItemData(intCount) = !������Ŀid Then
                        Me.cbo�������.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        End With

        
        '�����Ƿ��з�����ȷ�����ۼ۵�λ��ҩ�����ԡ��ɱ��ۡ����ۼ۸���޸ķ�
        gstrSql = " Select nvl(Count(*),0) " & _
            " From (Select 1 From ҩƷ�շ���¼ Where ҩƷID=[1] And rownum<2" & _
            "       Union Select 1 From ҩƷ��� Where ҩƷID=[1] And rownum<2)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        mblnUsed = False
        If rsTemp.Fields(0).Value > 0 Then
            mblnUsed = True
            If Me.cboҩ������.ListIndex <> 0 Then Me.cboҩ������.Enabled = False
            Me.txt�ɱ��۸�.Enabled = False
            Me.txt��ǰ�ۼ�.Enabled = False
            Me.cbo�������.Enabled = False
'            Me.txt����ϵ��.Enabled = False
            Me.txtҩ����װ.Enabled = False
            Me.txtҩ���װ.Enabled = False
        Else
            Me.cboҩ������.Enabled = True
            Me.txt�ɱ��۸�.Enabled = True
            Me.txt��ǰ�ۼ�.Enabled = True
            Me.cbo�������.Enabled = True
'            Me.txt����ϵ��.Enabled = True
            Me.txtҩ����װ.Enabled = True
            Me.txtҩ���װ.Enabled = True
        End If
        
        '�����Ƿ����ҽ����¼��ȷ������ϵ���Ƿ��ܹ��޸�
        gstrSql = "Select 1 From ����ҽ����¼ Where �շ�ϸĿID=[1] And Rownum=1"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        If rsTemp.RecordCount > 0 Then
            Me.txt����ϵ��.Enabled = False
        Else
            Me.txt����ϵ��.Enabled = True
        End If
        
        '�����Ƿ��п�棬ȷ���������Կ��޸ķ�
        gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                 " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And B.�������� Like '%ҩ��'"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        If rsTemp.Fields(0).Value > 0 Then
            Me.chkҩ��.Enabled = False
        Else
            Me.chkҩ��.Enabled = True
        End If
        If Me.chkҩ��.Value = 1 Then
            gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                     " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '%ҩ��' Or B.�������� Like '%�Ƽ���')"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
            
            If rsTemp.Fields(0).Value > 0 Then
                Me.chkҩ��.Enabled = False
                If Me.chkҩ��.Enabled Then Me.chkҩ��.Enabled = IIf(chkҩ��.Value = 1, False, True)
            Else
                Me.chkҩ��.Enabled = True
            End If
        End If
                
    End If
    
    '----------����Ȩ�޿���-------------------------------------
    If Me.stbSpec.Tag = "����" Or Me.stbSpec.Tag = "�޸�" Then
        If InStr(1, strPrivs, "Ŀ¼��ɾ��") = 0 Then
            Me.txt����.Enabled = False: Me.cmd����.Enabled = False
            Me.txt����.Enabled = False: Me.txt����.Enabled = False
            Me.txt����.Enabled = False: Me.txtƴ��.Enabled = False: Me.txt���.Enabled = False: Me.msf����.Active = False
            Me.cbo��λ.Enabled = False: Me.txtҩ�ⵥλ.Enabled = False: Me.txtҩ���װ.Enabled = False
            Me.cbo����.Enabled = False: Me.cbo��ֵ.Enabled = False: Me.cbo��Դ.Enabled = False: Me.cbo�ݴ�.Enabled = False
            Me.cboҩƷ����.Enabled = False: Me.cbo����ְ��.Enabled = False: Me.txt��������.Enabled = False ': Me.txt���췧ֵ.Enabled = False
            Me.chkԭ��ҩ.Enabled = False: Me.chk����Ӧ��.Enabled = False
            Me.cbo�������.Enabled = False: Me.chk���ηѱ�.Enabled = False
            Me.chkҩ��.Enabled = False: Me.chkҩ��.Enabled = False
            Me.txt�ο�.Enabled = False
            Me.cmd�ο�.Enabled = False
            Me.txt��ͬ��λ.Enabled = False: Me.cmd��ͬ��λ.Enabled = False
            Me.txt˵��.Enabled = False
            Me.txt���.Enabled = False
            Me.txt�ۼ۵�λ.Enabled = False
            Me.txt����ϵ��.Enabled = False
            Me.txtҩ����λ.Enabled = False
            Me.txtҩ����װ.Enabled = False
            Me.cbo�ɷ����.Enabled = False
            Me.cbo��ҩ����.Enabled = False
            Me.txt��ѡ��.Enabled = False
            Me.cmbStationNo.Enabled = False
            Me.txt��ֵ˰��.Enabled = False
            Me.chk���ҩ.Enabled = False
            Me.cbo�����Ա�.Enabled = False
        End If
        If InStr(1, strPrivs, "ҽ����ҩĿ¼") = 0 Then
            Me.cboҽ��ְ��.Enabled = False: Me.cbo��������.Enabled = False: Me.txt��ʶ��.Enabled = False
        End If
        If InStr(1, strPrivs, "�������") = 0 Then Me.txt����.Enabled = False
        If InStr(1, strPrivs, "ָ���۸����") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txtָ������.Text = "0"
                Me.txtָ���ۼ�.Text = "0"
            End If
            Me.txtָ������.Enabled = False: Me.txt�ӳ���.Enabled = False
            Me.txtָ������.Enabled = False: Me.txtָ���ۼ�.Enabled = False
        End If
        If InStr(1, strPrivs, "�ۼ۹���") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txt��ǰ�ۼ�.Text = "0"
                Me.cboҩ������.ListIndex = 0
            End If
            Me.cboҩ������.Enabled = False: Me.cbo�������.Enabled = False
            Me.txt��ǰ�ۼ�.Enabled = False
        End If
        If InStr(1, strPrivs, "ҩ�ۼ���") = 0 Then
             Me.cboҩ�ۼ���.Enabled = False
        End If
        If InStr(1, strPrivs, "�ɱ��۹���") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txt�ɱ��۸�.Text = "0"
            End If
            Me.txt�ɱ��۸�.Enabled = False
        End If
        If InStr(1, strPrivs, "�����������") = 0 Then
            Me.cbo�������.Enabled = False
        End If
    Else
        cmdOK.Visible = False: cmdCancel.Caption = "�ر�(&C)"
        
        Me.txt����.Enabled = False: Me.cmd����.Enabled = False
        Me.txt����.Enabled = False: Me.txt��ʶ��.Enabled = False: Me.txt����.Enabled = False
        Me.txt����.Enabled = False: Me.txtƴ��.Enabled = False: Me.txt���.Enabled = False: Me.msf����.Active = False
        Me.cbo��λ.Enabled = False: Me.txtҩ�ⵥλ.Enabled = False: Me.txtҩ���װ.Enabled = False
        Me.txt���췧ֵ.Enabled = False: Me.cbo���쵥λ.Enabled = False
        Me.cbo����.Enabled = False: Me.cbo��ֵ.Enabled = False: Me.cbo��Դ.Enabled = False: Me.cbo�ݴ�.Enabled = False
        Me.cboҩƷ����.Enabled = False: Me.cbo����ְ��.Enabled = False: Me.cboҽ��ְ��.Enabled = False: Me.txt��������.Enabled = False
        Me.chkԭ��ҩ.Enabled = False: Me.chk����Ӧ��.Enabled = False
        
        Me.cboҩ������.Enabled = False: Me.txtָ������.Enabled = False: Me.txt����.Enabled = False: Me.txt�����.Enabled = False
        Me.txtָ���ۼ�.Enabled = False: Me.txtָ������.Enabled = False: Me.txt�ӳ���.Enabled = False
        Me.cboҩ�ۼ���.Enabled = False: Me.cbo��������.Enabled = False: Me.cbo�������.Enabled = False: Me.chk���ηѱ�.Enabled = False
        Me.txt�ɱ��۸�.Enabled = False: Me.txt��ǰ�ۼ�.Enabled = False: Me.cbo�������.Enabled = False
        Me.chkҩ��.Enabled = False: Me.chkҩ��.Enabled = False: Me.txt����ѱ���.Enabled = False
        Me.txt�ο�.Enabled = False
        Me.cmd�ο�.Enabled = False
        Me.txt��ͬ��λ.Enabled = False: Me.cmd��ͬ��λ.Enabled = False
        Me.txt˵��.Enabled = False
        Me.txt���.Enabled = False
        Me.txt�ۼ۵�λ.Enabled = False
        Me.txt����ϵ��.Enabled = False
        Me.txtҩ����λ.Enabled = False
        Me.txtҩ����װ.Enabled = False
        Me.cbo�ɷ����.Enabled = False
        Me.cbo��ҩ����.Enabled = False
        Me.txt��ѡ��.Enabled = False
        Me.cmbStationNo.Enabled = False
        Me.txt��ֵ˰��.Enabled = False
        Me.chk���ҩ.Enabled = False
        Me.cbo�����Ա�.Enabled = False
    End If
    
    '������β������޸ģ������Ƿ���ڡ�ҩƷ��λ������Ȩ�ޣ�û���������޸�ҩƷ��λ��ϵ��
    If Me.stbSpec.Tag = "�޸�" Then
        If InStr(1, strPrivs, "ҩƷ��λ����") = 0 Then
            cbo��λ.Enabled = False
            txtҩ�ⵥλ.Enabled = False
            txtҩ���װ.Enabled = False
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
            tvwClass.Visible = False: txt����.SetFocus: Exit Sub
        End If
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    If glngSys \ 100 = 8 Then Me.lblҩ�ⵥλ.Caption = "�ɹ���λ(&W)"
    mint������� = Val(GetSysPara(87))
    
    Call GetDefineSize
    Call IniStationNo
    
    '-------------����ѡ������װ��-----------------------
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        '����ѡ����װ��
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ���Ʒ���Ŀ¼" & _
                " Where ���� = 3" & _
                " start with �ϼ�ID is null" & _
                " connect by prior ID=�ϼ�ID"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        
        gstrSql = "select distinct ���㵥λ from ������ĿĿ¼ where ���='7' and ���㵥λ is not null"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Me.cbo��λ.AddItem .Fields(0).Value
            .MoveNext
        Loop
        
        gstrSql = "select ����||'-'||���� from ҩƷ������� order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem .Fields(0).Value
            If InStr(1, .Fields(0).Value, "��ͨ") > 0 Then
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo����.ListIndex = -1 And Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ֵ���� order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo��ֵ.Clear
        Do While Not .EOF
            Me.cbo��ֵ.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo��ֵ.ListCount > 0 Then Me.cbo��ֵ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��Դ��� order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo��Դ.Clear
        Do While Not .EOF
            Me.cbo��Դ.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo��Դ.ListCount > 0 Then Me.cbo��Դ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ҩ�ݴ� order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo�ݴ�.Clear
        Do While Not .EOF
            Me.cbo�ݴ�.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo�ݴ�.ListCount > 0 Then Me.cbo�ݴ�.ListIndex = 0
    
        gstrSql = "Select ����||'-'||���� From �������� where ����=1 Order By ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo��������.Clear
        Me.cbo��������.AddItem ""
        Do While Not .EOF
            Me.cbo��������.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        
        gstrSql = "Select ID,'['||����||']'||���� as ����" & _
                " From ������Ŀ" & _
                " where ĩ��=1 and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo�������.Clear
        Do While Not .EOF
            Me.cbo�������.AddItem !����: Me.cbo�������.ItemData(Me.cbo�������.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo�������.ListCount > 0 Then Me.cbo�������.ListIndex = 0
    
        Me.lbl�������.Tag = zldatabase.GetPara("�в�ҩ������Ŀ", glngSys, 1023, False)
        For intCount = 0 To Me.cbo�������.ListCount - 1
            If Me.cbo�������.ItemData(intCount) = Val(Me.lbl�������.Tag) Then
                Me.cbo�������.ListIndex = intCount: Exit For
            End If
        Next
        
        gstrSql = "Select ���� From ��ҩ���� Order By ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cbo��ҩ����.Clear
        Do While Not .EOF
            Me.cbo��ҩ����.AddItem .Fields(0).Value
            .MoveNext
        Loop
        
        Me.lbl���۵�λ(0).Tag = Val(GetSysPara(29))
        
        mintCostDigit = GetDigit(1, 1, IIf(Me.lbl���۵�λ(0).Tag = 0, 1, 4))
        mintPriceDigit = GetDigit(1, 2, IIf(Me.lbl���۵�λ(0).Tag = 0, 1, 4))
        
        mintSaleCostDigit = GetDigit(1, 1, 1)
        mintSalePriceDigit = GetDigit(1, 2, 1)
    End With
    
    With Me.cbo���쵥λ
        .Clear
        .AddItem "�ۼ۵�λ"
        .AddItem "ҩ����λ"
        .AddItem "ҩ�ⵥλ"
        .ListIndex = 0
    End With
    
    With Me.cboҩ������
        .Clear
        aryTemp = Split("0-����;1-ʱ��", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(intCount): .ItemData(.NewIndex) = intCount
        Next
        .ListIndex = 0
    End With
    
    gstrSql = "Select ����||'-'||���� ���� From ҩ�۹����� where ����=1 Order By ����"
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.cboҩ�ۼ���.Clear
        Do While Not .EOF
            Me.cboҩ�ۼ���.AddItem !����
            .MoveNext
        Loop
    End With
        
    With Me.cbo�������
        If glngSys \ 100 <> 8 Then
            aryTemp = Split("0-��Ӧ���ڲ���;1-����;2-סԺ;3-�����סԺ", ";")
            For intCount = LBound(aryTemp) To UBound(aryTemp)
                .AddItem aryTemp(intCount): .ItemData(.NewIndex) = intCount
            Next
            .ListIndex = 3
        Else
            .AddItem "0-������": .ItemData(.NewIndex) = 0
            .AddItem "1-����": .ItemData(.NewIndex) = 3
            .ListIndex = 0
        End If
    End With
    
    With Me.cboҩƷ����
        .Clear
        aryTemp = Split("0-δ�趨;1-����ҩ;2-����Ǵ���ҩ;3-����Ǵ���ҩ;4-�Ǵ���ҩ;5-������ҩ", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(intCount)
        Next
        .ListIndex = 0
    End With
    
    Me.cbo����ְ��.Clear: Me.cboҽ��ְ��.Clear
    aryTemp = Split("0-����;1-����;2-����;3-�м�;4-����/ʦ��;5-Ա/ʿ;9-��Ƹ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����ְ��.AddItem aryTemp(intCount): Me.cboҽ��ְ��.AddItem aryTemp(intCount)
    Next
    Me.cbo����ְ��.ListIndex = 0: Me.cboҽ��ְ��.ListIndex = 0
    
    aryTemp = Split("0-���Ա�����;1-����;2-Ů��", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�����Ա�.AddItem aryTemp(intCount)
    Next
    Me.cbo�����Ա�.ListIndex = 0
    
    With Me.cbo�ɷ����
        .Clear
        .AddItem "0-���Է���": .ItemData(.NewIndex) = 0
        .AddItem "1-���ɷ���": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    '----------------�༭��������----------------------
    With Me.msf����
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "ҩƷ����": .TextMatrix(0, 2) = "ƴ����": .TextMatrix(0, 3) = "�����"
        .ColData(0) = 5: .ColData(1) = 4: .ColData(2) = 4: .ColData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 1000: .ColWidth(2) = 650: .ColWidth(3) = 650
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    strRefer = ""
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msf����_AfterAddRow(Row As Long)
    With Me.msf����
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_AfterDeleteRow()
    With Me.msf����
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_EditKeyPress(KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Exit Sub
            strTemp = Trim(.Text)
            If strTemp = "" Then Exit Sub
            .TextMatrix(.Row, 1) = strTemp
            .TextMatrix(.Row, 2) = zlGetSymbol(strTemp, 0, mint���볤��)
            .TextMatrix(.Row, 3) = zlGetSymbol(strTemp, 1, mint���볤��)
        End If
    End With
End Sub

Private Sub msf����_KeyPress(KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf��ͬ��λ_DblClick()
    With Me.msf��ͬ��λ
        Me.txt��ͬ��λ.Text = .TextMatrix(.Row, 1)
        Me.txt��ͬ��λ.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With
    Me.txt��ͬ��λ.SetFocus
    Call zlcommfun.PressKey(vbKeyTab)
End Sub


Private Sub msf��ͬ��λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call msf��ͬ��λ_DblClick
End Sub

Private Sub msf��ͬ��λ_LostFocus()
    Me.msf��ͬ��λ.Visible = False
End Sub





Private Sub stbSpec_Click(PreviousTab As Integer)
    Select Case stbSpec.Tab
    Case 0
        If Me.txt����.Enabled Then Me.txt����.SetFocus
    Case 1
        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus
        If Me.cboҩ������.Enabled Then Me.cboҩ������.SetFocus
    End Select
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    Me.txt����.SetFocus
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
    If Me.cmd���� Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt��ѡ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ѡ��_Validate(Cancel As Boolean)
    Dim i As Integer
    
    If Len(Trim(txt��ѡ��.Text)) > 0 Then
        For i = 1 To Len(Trim(txt��ѡ��.Text))
            If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Trim(txt��ѡ��.Text), i, 1)) < 1 Then
                MsgBox "��ѡ�����������ĸ��������ɡ�", vbExclamation, gstrSysName
                Me.stbSpec.Tab = 0
                If txt��ѡ��.Enabled And txt��ѡ��.Visible Then
                    txt��ѡ��.SetFocus
                End If
            End If
        Next
    End If
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
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

Private Sub txt��ʶ��_GotFocus()
    Me.txt��ʶ��.SelStart = 0: Me.txt��ʶ��.SelLength = 100
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*_+|=-`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii > 255 Or KeyAscii < 0 Then KeyAscii = 0
End Sub

Private Sub txt�ο�_GotFocus()
    Me.txt�ο�.SelStart = 0: Me.txt�ο�.SelLength = 100
End Sub


Private Sub txt�ο�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt�ο� <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt�ο�))
            If rsTmp Is Nothing Then
                Me.txt�ο� = strRefer
                Me.SetFocus
                Exit Sub
            Else
                Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
            End If
        End If
        Call zlcommfun.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt�ο�_LostFocus()
    If Me.txt�ο� <> strRefer Then
        Me.txt�ο� = strRefer
    End If
End Sub


Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt�ɱ��۸�_GotFocus()
    Me.txt�ɱ��۸�.SelStart = 0: Me.txt�ɱ��۸�.SelLength = 100
End Sub

Private Sub txt�ɱ��۸�_KeyPress(KeyAscii As Integer)
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

Private Sub txt�ɱ��۸�_LostFocus()
    Dim dblSalePrice As Double
    Me.txt�ɱ��۸�.Text = FormatEx(Val(Me.txt�ɱ��۸�.Text), mintCostDigit)
    If Val(Me.txt��ǰ�ۼ�.Text) = 0 And Val(Me.txt�ɱ��۸�.Text) <> 0 Then
        dblSalePrice = Val(Me.txt�ɱ��۸�.Text) * (1 + Val(Me.txt�ӳ���.Text) / 100)
        If dblSalePrice > Val(Me.txtָ���ۼ�.Text) Then dblSalePrice = Val(Me.txtָ���ۼ�.Text)
        Me.txt��ǰ�ۼ�.Text = FormatEx(dblSalePrice, mintPriceDigit)
    End If
End Sub

Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 100
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��ǰ�ۼ�_GotFocus()
    Me.txt��ǰ�ۼ�.SelStart = 0: Me.txt��ǰ�ۼ�.SelLength = 100
End Sub

Private Sub txt��ǰ�ۼ�_KeyPress(KeyAscii As Integer)
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

Private Sub txt��ǰ�ۼ�_LostFocus()
Dim dbl�ɱ��� As Double
    Dim dblָ���ۼ� As Double
    Dim dbl�ӳ��� As Double
    Dim dbl����� As Double
    Dim dbl���ۼ� As Double
    
    Me.txt��ǰ�ۼ�.Text = FormatEx(Val(txt��ǰ�ۼ�), mintPriceDigit)
    
    dbl���ۼ� = Val(Me.txt��ǰ�ۼ�.Text)
    dbl�ɱ��� = Val(Me.txt�ɱ��۸�.Text)
    dblָ���ۼ� = Val(Me.txtָ���ۼ�.Text)
    
    '������Щ�����ż���ӳ���
    If dbl�ɱ��� > 0 And dblָ���ۼ� > 0 And dbl���ۼ� > 0 And dbl���ۼ� <= dblָ���ۼ� Then
        dbl�ӳ��� = dbl���ۼ� / dbl�ɱ��� - 1
        
        If dbl�ӳ��� < 0 Then Exit Sub
        
        dbl�ӳ��� = dbl�ӳ��� * 100
        
        Me.txt�ӳ���.Text = Format(dbl�ӳ���, "0.00")
        
        'ͨ���ӳ��ʼ���ָ�������
        dbl����� = dbl�ӳ���
        Call Calc(dbl�����, False)
        Me.txtָ������.Text = Format(dbl�����, "0.00000")
    End If
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����ѱ���_GotFocus()
    txt����ѱ���.SelStart = 0: txt����ѱ���.SelLength = 100
End Sub

Private Sub txt����ѱ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txt����ѱ���_Validate(Cancel As Boolean)
    txt����ѱ���.Text = Format(Val(txt����ѱ���.Text), "#0.00")
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt��ͬ��λ_GotFocus()
    Me.txt��ͬ��λ.SelStart = 0: Me.txt��ͬ��λ.SelLength = Len(Me.txt��ͬ��λ.Text)
End Sub

Private Sub txt��ͬ��λ_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
        
    strTmp = UCase(Trim(Me.txt��ͬ��λ.Text))
    
    If strTmp = "" Then
        Me.txt��ͬ��λ.Tag = "|"
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    ElseIf strTmp = Split(Me.txt��ͬ��λ.Tag, "|")(1) Then
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    End If
       
    gstrSql = "Select ����,����,����,id" & _
            " From ��Ӧ��" & _
            " where (���� Like [1] " & _
            "       Or ���� Like [2] " & _
            "       Or ���� Like [2])" & _
            " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    
    With rsTemp
        If .EOF Then
            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
            Me.txt��ͬ��λ.Text = Split(Me.txt��ͬ��λ.Tag, "|")(1)
            Me.txt��ͬ��λ.SelStart = 0: Me.txt��ͬ��λ.SelLength = Len(Me.txt��ͬ��λ.Text)
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            Me.txt��ͬ��λ.Text = Trim(rsTemp!����): Me.txt��ͬ��λ.Tag = rsTemp!ID & "|" & rsTemp!����
            Call zlcommfun.PressKey(vbKeyTab): Exit Sub
        Else
            With Me.msf��ͬ��λ
                .Left = Me.stbSpec.Left + Me.txt��ͬ��λ.Left
                .Top = Me.stbSpec.Top + Me.txt��ͬ��λ.Top - Me.msf��ͬ��λ.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 1500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
End Sub


Private Sub txt��ͬ��λ_Validate(Cancel As Boolean)
    If Me.txt��ͬ��λ.Text = "" Then
        Me.txt��ͬ��λ.Tag = "|"
    ElseIf Me.txt��ͬ��λ.Text <> Split(Me.txt��ͬ��λ.Tag, "|")(1) Then
        txt��ͬ��λ_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub txt����ϵ��_Change()
    If glngSys \ 100 = 8 Then
        Me.txtҩ����װ = 1
    End If
End Sub

Private Sub txt����ϵ��_GotFocus()
    Me.txt����ϵ��.SelStart = 0: Me.txt����ϵ��.SelLength = 100
End Sub


Private Sub txt����ϵ��_KeyPress(KeyAscii As Integer)
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


Private Sub txt�ӳ���_GotFocus()
    Call zlControl.TxtSelAll(txt�ӳ���)
End Sub

Private Sub txt�ӳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub txt�ӳ���_LostFocus()
    Dim cur�۸� As Double
    '���¼���ָ������ʺͼӳ���
    cur�۸� = Val(txt�ӳ���.Text)
    Call Calc(cur�۸�, False)
    Me.txt�ӳ���.Text = Format(txt�ӳ���.Text, "0.00")
    Me.txtָ������.Text = Format(cur�۸�, "0.00000")
End Sub

Private Sub txt�����_GotFocus()
    Me.txt�����.SelStart = 0: Me.txt�����.SelLength = 100
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
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

Private Sub txt�����_LostFocus()
    Me.txt�����.Text = FormatEx(Val(txt�����), mintPriceDigit)
End Sub

Private Sub txt����_Change()
    Me.txt�����.Text = FormatEx(Val(Me.txtָ������.Text) * Val(Me.txt����.Text) / 100, mintCostDigit)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
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

Private Sub txt����_LostFocus()
    Me.txt����.Text = FormatEx(Val(txt����), mintPriceDigit)
End Sub

Private Sub txt����_Change()
    Dim strTmp As String
    '���¼�����ƣ���ȥ �������ַ�
    strTmp = MoveSpecialChar(txt����.Text)
    If txt����.Text <> strTmp Then
        txt����.Text = strTmp
    End If
    Me.txtƴ��.Text = zlGetSymbol(strTmp, 0, mint���볤��)
    Me.txt���.Text = zlGetSymbol(strTmp, 1, mint���볤��)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
    End Select
    If KeyAscii = vbKeyReturn Then
        Call zlcommfun.PressKey(vbKeyTab)
    Else
        If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Me.txtƴ��.Text = zlGetSymbol(Me.txt����.Text, 0, mint���볤��)
        Me.txt���.Text = zlGetSymbol(Me.txt����.Text, 1, mint���볤��)
    End If
End Sub

Private Sub txt����_LostFocus()
    Me.txtƴ��.Text = zlGetSymbol(Me.txt����.Text, 0, mint���볤��)
    Me.txt���.Text = zlGetSymbol(Me.txt����.Text, 1, mint���볤��)
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txtƴ��_GotFocus()
    Me.txtƴ��.SelStart = 0: Me.txtƴ��.SelLength = 100
End Sub

Private Sub txtƴ��_KeyPress(KeyAscii As Integer)
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

Private Sub txt���췧ֵ_GotFocus()
    txt���췧ֵ.SelStart = 0: txt���췧ֵ.SelLength = Len(txt���췧ֵ)
End Sub

Private Sub txt���췧ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt�ۼ۵�λ_Change()
    Me.lbl����ϵ��.Caption = "(1" & Me.txt�ۼ۵�λ.Text & "="
    If glngSys \ 100 = 8 Then
        Me.txtҩ����λ = Me.txt�ۼ۵�λ
    End If
    Me.lblҩ����λChild.Caption = Me.txt�ۼ۵�λ & ")"
    Me.lblҩ�ⵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
    Me.lbl���쵥λChild.Caption = Me.txt�ۼ۵�λ & ")"
    If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
        Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
        Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
    Else
        Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
        Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
    End If
    Call cbo���쵥λ_Click
End Sub

Private Sub txt�ۼ۵�λ_GotFocus()
    Me.txt�ۼ۵�λ.SelStart = 0: Me.txt�ۼ۵�λ.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub


Private Sub txt�ۼ۵�λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt�ۼ۵�λ_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub


Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
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

Private Sub txtҩ����װ_GotFocus()
    Me.txtҩ����װ.SelStart = 0: Me.txtҩ����װ.SelLength = 100
End Sub


Private Sub txtҩ����װ_KeyPress(KeyAscii As Integer)
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


Private Sub txtҩ����λ_Change()
    Me.lblҩ����װ.Caption = "(1" & Me.txtҩ����λ.Text & "="
    Call cbo���쵥λ_Click
End Sub

Private Sub txtҩ����λ_GotFocus()
    Me.txtҩ����λ.SelStart = 0: Me.txtҩ����λ.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub


Private Sub txtҩ����λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txtҩ����λ_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub


Private Sub txtҩ���װ_GotFocus()
    Me.txtҩ���װ.SelStart = 0: Me.txtҩ���װ.SelLength = 100
End Sub

Private Sub txtҩ���װ_KeyPress(KeyAscii As Integer)
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

Private Sub txtҩ�ⵥλ_Change()
    Me.lblҩ���װ.Caption = "(1" & Me.txtҩ�ⵥλ.Text & "="
    If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
        Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
        Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txtҩ�ⵥλ.Text
    Else
        Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
        Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txt�ۼ۵�λ.Text
    End If
    Call cbo���쵥λ_Click
End Sub

Private Sub txtҩ�ⵥλ_GotFocus()
    Me.txtҩ�ⵥλ.SelStart = 0: Me.txtҩ�ⵥλ.SelLength = 100
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txtҩ�ⵥλ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtҩ�ⵥλ_LostFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt��ֵ˰��_GotFocus()
    Call zlControl.TxtSelAll(txt��ֵ˰��)
End Sub

Private Sub txt��ֵ˰��_KeyPress(KeyAscii As Integer)
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


Private Sub txt��ֵ˰��_LostFocus()
    txt��ֵ˰��.Text = Format(txt��ֵ˰��.Text, "0.00")
End Sub


Private Sub txtָ������_GotFocus()
    Me.txtָ������.SelStart = 0: Me.txtָ������.SelLength = 100
End Sub

Private Sub txtָ������_KeyPress(KeyAscii As Integer)
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

Private Sub txtָ������_LostFocus()
    Dim cur�۸� As Double
    '���¼���ָ������ʺͼӳ���
    cur�۸� = Val(txtָ������.Text)
    If cur�۸� < 100 Then
        Call Calc(cur�۸�, True)
        Me.txtָ������.Text = Format(txtָ������.Text, "0.00000")
        Me.txt�ӳ���.Text = Format(cur�۸�, "0.00")
    Else
        '���������ָ������ʴ��ڵ���100������������Ҫ�Ӽӳ��ʷ������
        cur�۸� = Val(txt�ӳ���.Text)
        Call Calc(cur�۸�, False)
        Me.txtָ������.Text = Format(cur�۸�, "0.00000")
    End If
End Sub

Private Sub txtָ������_Change()
    Me.txt�����.Text = FormatEx(Val(Me.txtָ������.Text) * Val(Me.txt����.Text) / 100, mintCostDigit)
End Sub

Private Sub txtָ������_GotFocus()
    Me.txtָ������.SelStart = 0: Me.txtָ������.SelLength = 100
End Sub

Private Sub txtָ������_KeyPress(KeyAscii As Integer)
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

Private Sub txtָ������_LostFocus()
    Me.txtָ������.Text = FormatEx(Val(txtָ������), mintCostDigit)
End Sub

Private Sub txtָ���ۼ�_GotFocus()
    Me.txtָ���ۼ�.SelStart = 0: Me.txtָ���ۼ�.SelLength = 100
End Sub

Private Sub txtָ���ۼ�_KeyPress(KeyAscii As Integer)
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

Private Sub txtָ���ۼ�_LostFocus()
    Me.txtָ���ۼ�.Text = FormatEx(Val(txtָ���ۼ�), mintPriceDigit)
End Sub

Private Sub Calc(dbl�۸� As Double, Optional ByVal bln����� As Boolean = True)
    '���������ǲ���ʣ�����ӳ��ʲ����أ�����������ʲ�����
    '�ӳ��������ʼ䣬�������ж�Ӧ��ϵ
    '�ӳ���=1/(1-�����)-1
    '�����=1-1/(1+�ӳ���)
    dbl�۸� = dbl�۸� / 100
    If bln����� Then
        dbl�۸� = 1 / (1 - dbl�۸�) - 1
    Else
        dbl�۸� = 1 - 1 / (1 + dbl�۸�)
    End If
    dbl�۸� = dbl�۸� * 100
End Sub

Private Function CheckRequest() As Boolean
    Dim dbl�������� As Double
    Dim str�������� As String
    '������췧ֵת��Ϊ���۵�λ���Ƿ�Ϊ����������5λС������ʾ����Ա����ǿ�Ʊ���
    dbl�������� = Val(txt���췧ֵ.Text)
    
    Select Case cbo���쵥λ.ListIndex
    Case 1 'ҩ����λ
        dbl�������� = dbl�������� * Val(txtҩ����װ.Text)
    Case 2 'ҩ�ⵥλ
        dbl�������� = dbl�������� * Val(txtҩ���װ.Text)
    End Select
    txt���췧ֵ.Tag = dbl��������
    
    CheckRequest = True
End Function
