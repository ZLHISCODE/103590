VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ���ά��"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmMediSpec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "���������Ʒ��(&A)"
      Height          =   350
      Left            =   3120
      TabIndex        =   68
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "������������(&B)"
      Height          =   350
      Left            =   5040
      TabIndex        =   69
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�����˳�(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   66
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox picFound 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5040
      ScaleHeight     =   210
      ScaleWidth      =   5145
      TabIndex        =   113
      Top             =   400
      Width           =   5145
      Begin VB.Label lblFound 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ù������2002��12��20�գ���2003��8��10��ͣ��"
         Height          =   180
         Left            =   180
         TabIndex        =   122
         Top             =   0
         Width           =   4230
      End
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   112
      Top             =   285
      Width           =   9525
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   67
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frmMediSpec.frx":058A
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1100
   End
   Begin TabDlg.SSTab stbSpec 
      Height          =   5835
      Left            =   120
      TabIndex        =   106
      Top             =   360
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   10292
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "�����Ϣ(&1)"
      TabPicture(0)   =   "frmMediSpec.frx":06D4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl��Ʒ��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl��ʶ��"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl����"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl���"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl����"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl�ۼ۵�λ"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl����ϵ��"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl���ﵥλ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl�����װ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblסԺ��λ"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblסԺ��װ"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblҩ�ⵥλ"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblҩ���װ"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl������"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl�ۼ۵�λChild"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblסԺ��λChild"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl���ﵥλChild"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblҩƷ��Դ"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl��׼�ĺ�"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl����"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblע���̱�"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "���쵥λ"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl��ͬ��λ"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblComment"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl˵��"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl��ҩ����"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl��ѡ��"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblStationNo"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl����child"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl����"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl���쵥λ"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblҩ�ⵥλChild"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblddd"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbldddֵ"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl��ΣҩƷ"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lbl�ͻ���λ"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl�ͻ���װ"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lbl�ͻ���λchild"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl��λ��"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt��ͬ��λ"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtƴ��"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt��Ʒ��"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt��ʶ��"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt����"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt���"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt����ϵ��"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt���ﵥλ"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt�����װ"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtסԺ��λ"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtסԺ��װ"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtҩ�ⵥλ"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtҩ���װ"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt�ۼ۵�λ"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt������"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt����"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cboҩƷ��Դ"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt��׼�ĺ�"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt���"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtע���̱�"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmd��ͬ��λ"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt˵��"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt��ѡ��"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmbStationNo"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt����"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "cbo���쵥λ"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt���췧ֵ"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "cbo��ҩ����"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "cmd����"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtDDDֵ"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cbo��ΣҩƷ"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txt�ͻ���λ"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txt�ͻ���װ"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txt��λ��"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).ControlCount=   75
      TabCaption(1)   =   "ҩ����Ϣ(&2)"
      TabPicture(1)   =   "frmMediSpec.frx":06F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblָ���ۼ�"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblָ������"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPercent(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl����"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl�����"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbl���۵�λ(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblҩ������"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblҩ�ۼ���"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl��ǰ�ۼ�"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl�������"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl�ӳ���"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl�ɱ��۸�"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl��������"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lbl�ɷ����"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lbl�������"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblBasicDrug"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbl���۵�λ(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lbl������Ŀ"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label3"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txt������Ŀ"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtָ���ۼ�"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtָ������"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cboҩ�ۼ���"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cbo�������"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txt��ǰ�ۼ�"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cbo��������"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cbo�������"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "cboסԺ����"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txt����"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txt�����"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cboҩ������"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "fra��������"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "chkGMP��֤"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txt�ӳ���"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txt�ɱ��۸�"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "chk���ηѱ�"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkסԺ��̬����"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "chk�ǳ���ҩ"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cboBasicDrug"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "cmd����"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "cbo�������"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "chk��ҩ"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "chk����"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "chk�׵���"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "chk�����ɹ�"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "��ҩ����(&3)"
      TabPicture(2)   =   "frmMediSpec.frx":070C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtNotice"
      Tab(2).Control(1)=   "chkDosage"
      Tab(2).Control(2)=   "chkCondition"
      Tab(2).Control(3)=   "cboPrepareType"
      Tab(2).Control(4)=   "cboTemperature"
      Tab(2).Control(5)=   "lblNotice"
      Tab(2).Control(6)=   "lblPrepareType"
      Tab(2).Control(7)=   "lblCondition"
      Tab(2).Control(8)=   "lblTemperature"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "��չ����(&4)"
      TabPicture(3)   =   "frmMediSpec.frx":0728
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "vsfItem"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chk�����ɹ� 
         Caption         =   "�����ɹ�"
         Height          =   180
         Left            =   -67080
         TabIndex        =   56
         Top             =   2050
         Width           =   1080
      End
      Begin VB.CheckBox chk�׵��� 
         Caption         =   "�׵���"
         Height          =   180
         Left            =   -68715
         TabIndex        =   55
         Top             =   2050
         Width           =   1080
      End
      Begin VB.TextBox txt��λ�� 
         Height          =   300
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   2
         Top             =   750
         Width           =   1995
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�������۹���ģʽ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   878
         Width           =   2895
      End
      Begin VB.CheckBox chk��ҩ 
         Caption         =   "��ҩ"
         Height          =   180
         Left            =   -67080
         TabIndex        =   54
         Top             =   1695
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.TextBox txtNotice 
         Height          =   1335
         Left            =   -74700
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txt�ͻ���װ 
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "90"
         Top             =   2700
         Width           =   945
      End
      Begin VB.TextBox txt�ͻ���λ 
         Height          =   300
         Left            =   5760
         MaxLength       =   8
         TabIndex        =   28
         Text            =   "��"
         Top             =   2700
         Width           =   585
      End
      Begin VB.ComboBox cbo��ΣҩƷ 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4440
         Width           =   3285
      End
      Begin VB.TextBox txtDDDֵ 
         Height          =   300
         Left            =   5790
         TabIndex        =   32
         Top             =   4800
         Width           =   1215
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -67320
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   855
         Width           =   1725
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   285
         Left            =   4150
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1515
         Width           =   285
      End
      Begin VB.ComboBox cbo��ҩ���� 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4080
         Width           =   3285
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Left            =   -69240
         TabIndex        =   131
         TabStop         =   0   'False
         Tag             =   "����"
         ToolTipText     =   "��*��ѡ����"
         Top             =   885
         Width           =   255
      End
      Begin VB.TextBox txt���췧ֵ 
         Height          =   300
         Left            =   7365
         MaxLength       =   8
         TabIndex        =   27
         Top             =   2295
         Width           =   585
      End
      Begin VB.ComboBox cbo���쵥λ 
         Height          =   300
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2295
         Width           =   1320
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   5790
         TabIndex        =   31
         Top             =   4380
         Width           =   1215
      End
      Begin VB.ComboBox cboBasicDrug 
         Height          =   300
         Left            =   -70680
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   2490
         Width           =   1725
      End
      Begin VB.CheckBox chkDosage 
         Caption         =   "������䣨��������ҩ��ֱ�Ӵ�����ͣ�"
         Height          =   255
         Left            =   -74700
         TabIndex        =   64
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CheckBox chkCondition 
         Caption         =   "�ܹ��ܱ�"
         Height          =   255
         Left            =   -73860
         TabIndex        =   62
         Top             =   923
         Width           =   1455
      End
      Begin VB.ComboBox cboPrepareType 
         Height          =   300
         Left            =   -73860
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1320
         Width           =   2445
      End
      Begin VB.ComboBox cboTemperature 
         Height          =   300
         Left            =   -73860
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   480
         Width           =   2445
      End
      Begin VB.CheckBox chk�ǳ���ҩ 
         Caption         =   "�ǳ���ҩ"
         Height          =   180
         Left            =   -68715
         TabIndex        =   53
         Top             =   1700
         Width           =   1080
      End
      Begin VB.ComboBox cmbStationNo 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   5220
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txt��ѡ�� 
         Height          =   300
         Left            =   5790
         MaxLength       =   20
         TabIndex        =   30
         Top             =   4005
         Width           =   2400
      End
      Begin VB.CheckBox chkסԺ��̬���� 
         Caption         =   "סԺ��̬����"
         Height          =   180
         Left            =   -68715
         TabIndex        =   51
         Top             =   1350
         Width           =   1440
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   12
         Top             =   3765
         Width           =   3285
      End
      Begin VB.CommandButton cmd��ͬ��λ 
         Caption         =   "��"
         Height          =   285
         Left            =   4140
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3405
         Width           =   285
      End
      Begin VB.CheckBox chk���ηѱ� 
         Alignment       =   1  'Right Justify
         Caption         =   "���ηѱ�(&M)"
         Height          =   285
         Left            =   -71820
         TabIndex        =   48
         Top             =   2880
         Width           =   1290
      End
      Begin VB.TextBox txtע���̱� 
         Height          =   300
         Left            =   5790
         MaxLength       =   50
         TabIndex        =   17
         Top             =   405
         Width           =   2400
      End
      Begin VB.TextBox txt�ɱ��۸� 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   35
         Top             =   1215
         Width           =   1485
      End
      Begin VB.TextBox txt�ӳ��� 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   41
         Text            =   "15.00"
         Top             =   3645
         Width           =   1470
      End
      Begin VB.CheckBox chkGMP��֤ 
         Caption         =   "GMP��֤(&Z)"
         Height          =   180
         Left            =   -67080
         TabIndex        =   52
         Top             =   1320
         Width           =   1290
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   2865
         MaxLength       =   12
         TabIndex        =   7
         Top             =   2250
         Width           =   1020
      End
      Begin VB.TextBox txt��׼�ĺ� 
         Height          =   300
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   15
         Top             =   4860
         Width           =   3285
      End
      Begin VB.ComboBox cboҩƷ��Դ 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3015
         Width           =   3300
      End
      Begin VB.Frame fra�������� 
         Caption         =   "��������(&K)"
         Height          =   1065
         Left            =   -68715
         TabIndex        =   105
         Top             =   2520
         Width           =   2520
         Begin VB.TextBox txtЧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   60
            Text            =   "0"
            Top             =   600
            Width           =   465
         End
         Begin VB.CheckBox chkЧ�� 
            Caption         =   "������(��)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   330
            TabIndex        =   59
            Top             =   660
            Width           =   1215
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "ҩ��"
            Height          =   210
            Left            =   330
            TabIndex        =   57
            Top             =   300
            Width           =   675
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "ҩ��"
            Enabled         =   0   'False
            Height          =   210
            Left            =   1470
            TabIndex        =   58
            Top             =   300
            Width           =   675
         End
      End
      Begin VB.ComboBox cboҩ������ 
         Height          =   300
         Left            =   -73860
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   450
         Width           =   1470
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1140
         MaxLength       =   13
         TabIndex        =   1
         Top             =   375
         Width           =   1995
      End
      Begin VB.TextBox txt����� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   39
         Top             =   2820
         Width           =   1470
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   38
         Text            =   "100"
         Top             =   2445
         Width           =   1470
      End
      Begin VB.ComboBox cboסԺ���� 
         Height          =   300
         Left            =   -67320
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   450
         Width           =   1725
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   1140
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2625
         Width           =   1020
      End
      Begin VB.TextBox txt�ۼ۵�λ 
         Height          =   300
         Left            =   5790
         MaxLength       =   8
         TabIndex        =   18
         Text            =   "Ƭ"
         Top             =   780
         Width           =   585
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2085
         Width           =   1725
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1695
         Width           =   1725
      End
      Begin VB.TextBox txt��ǰ�ۼ� 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   36
         Top             =   1605
         Width           =   1485
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   450
         Width           =   1725
      End
      Begin VB.ComboBox cboҩ�ۼ��� 
         Height          =   300
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1245
         Width           =   1725
      End
      Begin VB.TextBox txtָ������ 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   37
         Top             =   2055
         Width           =   1470
      End
      Begin VB.TextBox txtָ���ۼ� 
         Height          =   300
         Left            =   -73860
         MaxLength       =   16
         TabIndex        =   40
         Top             =   3240
         Width           =   1470
      End
      Begin VB.TextBox txtҩ���װ 
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "30"
         Top             =   1905
         Width           =   945
      End
      Begin VB.TextBox txtҩ�ⵥλ 
         Height          =   300
         Left            =   5790
         MaxLength       =   8
         TabIndex        =   24
         Text            =   "��"
         Top             =   1920
         Width           =   585
      End
      Begin VB.TextBox txtסԺ��װ 
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "1"
         Top             =   1155
         Width           =   945
      End
      Begin VB.TextBox txtסԺ��λ 
         Height          =   300
         Left            =   5790
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "֧"
         Top             =   1155
         Width           =   585
      End
      Begin VB.TextBox txt�����װ 
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "10"
         Top             =   1530
         Width           =   945
      End
      Begin VB.TextBox txt���ﵥλ 
         Height          =   300
         Left            =   5790
         MaxLength       =   8
         TabIndex        =   22
         Text            =   "��"
         Top             =   1530
         Width           =   585
      End
      Begin VB.TextBox txt����ϵ�� 
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "5"
         Top             =   780
         Width           =   945
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1125
         Width           =   3285
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1500
         Width           =   2985
      End
      Begin VB.TextBox txt��ʶ�� 
         Height          =   300
         Left            =   3165
         MaxLength       =   29
         TabIndex        =   9
         Top             =   2625
         Width           =   1275
      End
      Begin VB.TextBox txt��Ʒ�� 
         Height          =   300
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1875
         Width           =   3285
      End
      Begin VB.TextBox txtƴ�� 
         Height          =   300
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   6
         Top             =   2250
         Width           =   1020
      End
      Begin VB.TextBox txt������Ŀ 
         Height          =   300
         Left            =   -70680
         MaxLength       =   40
         TabIndex        =   43
         ToolTipText     =   "��*��ѡ����"
         Top             =   855
         Width           =   1725
      End
      Begin VB.TextBox txt��ͬ��λ 
         Height          =   300
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   11
         Top             =   3405
         Width           =   2985
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfItem 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   142
         Top             =   360
         Width           =   9195
         _cx             =   16219
         _cy             =   8705
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         FormatString    =   $"frmMediSpec.frx":0744
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
      Begin VB.Label lbl��λ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   143
         Top             =   802
         Width           =   540
      End
      Begin VB.Label lblNotice 
         Caption         =   "��Һע������"
         Height          =   255
         Left            =   -74700
         TabIndex        =   141
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lbl�ͻ���λchild 
         AutoSize        =   -1  'True
         Caption         =   "��)"
         Height          =   180
         Left            =   7980
         TabIndex        =   140
         Top             =   2760
         Width           =   270
      End
      Begin VB.Label lbl�ͻ���װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6555
         TabIndex        =   139
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label lbl�ͻ���λ 
         AutoSize        =   -1  'True
         Caption         =   "�ͻ���λ(&V)"
         Height          =   180
         Left            =   4770
         TabIndex        =   138
         Top             =   2760
         Width           =   990
      End
      Begin VB.Label lbl��ΣҩƷ 
         AutoSize        =   -1  'True
         Caption         =   "��ΣҩƷ(&0)"
         Height          =   180
         Left            =   105
         TabIndex        =   137
         Top             =   4545
         Width           =   990
      End
      Begin VB.Label lbldddֵ 
         Caption         =   "ml"
         Height          =   255
         Left            =   7080
         TabIndex        =   136
         Top             =   4830
         Width           =   1455
      End
      Begin VB.Label lblddd 
         Caption         =   "DDDֵ(&1)"
         Height          =   255
         Left            =   4770
         TabIndex        =   135
         Top             =   4830
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������ʹ��(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -68715
         TabIndex        =   134
         Top             =   915
         Width           =   1350
      End
      Begin VB.Label lbl������Ŀ 
         Caption         =   "������Ŀ(&F)"
         Height          =   255
         Left            =   -71820
         TabIndex        =   132
         Top             =   878
         Width           =   990
      End
      Begin VB.Label lbl���۵�λ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ/Ƭ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   -72360
         TabIndex        =   130
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label lblҩ�ⵥλChild 
         AutoSize        =   -1  'True
         Caption         =   "Ƭ)"
         Height          =   180
         Left            =   7980
         TabIndex        =   129
         Top             =   1965
         Width           =   300
      End
      Begin VB.Label lbl���쵥λ 
         AutoSize        =   -1  'True
         Caption         =   "Ƭ)"
         Height          =   180
         Left            =   7980
         TabIndex        =   128
         Top             =   2355
         Width           =   300
      End
      Begin VB.Label lbl���� 
         Caption         =   "����(&R)"
         Height          =   255
         Left            =   4770
         TabIndex        =   127
         Top             =   4440
         Width           =   630
      End
      Begin VB.Label lbl����child 
         Caption         =   "ml"
         Height          =   255
         Left            =   7080
         TabIndex        =   126
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblPrepareType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74700
         TabIndex        =   125
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblCondition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�洢����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74700
         TabIndex        =   124
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblTemperature 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�洢�¶�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74700
         TabIndex        =   123
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblBasicDrug 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҩ��(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   104
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "Ժ�����(&Z)"
         Height          =   180
         Left            =   105
         TabIndex        =   121
         Top             =   5280
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl��ѡ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��(&F)"
         Height          =   180
         Left            =   4770
         TabIndex        =   120
         Top             =   4065
         Width           =   810
      End
      Begin VB.Label lbl��ҩ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����(&H)"
         Height          =   180
         Left            =   105
         TabIndex        =   119
         Top             =   4185
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(����д�ʵ���˵��������ʾ���á�����֢ҩƷ��)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4755
         TabIndex        =   118
         Top             =   3495
         Width           =   3960
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ˵��(&X)"
         Height          =   180
         Left            =   105
         TabIndex        =   117
         Top             =   3810
         Width           =   990
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "(ָ���˺�ͬ��λ��ҩƷ��ֻ�ܰ���ͬ��λ��⡣)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4755
         TabIndex        =   116
         Top             =   3120
         Width           =   3960
      End
      Begin VB.Label lbl��ͬ��λ 
         AutoSize        =   -1  'True
         Caption         =   "��ͬ��λ(&C)"
         Height          =   180
         Left            =   105
         TabIndex        =   114
         Top             =   3450
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
         Left            =   4770
         TabIndex        =   88
         Top             =   2355
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
         Left            =   -71820
         TabIndex        =   102
         Top             =   2145
         Width           =   990
      End
      Begin VB.Label lbl�ɷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����ʹ��(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -68715
         TabIndex        =   103
         Top             =   510
         Width           =   1350
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
         Left            =   -71820
         TabIndex        =   101
         Top             =   1755
         Width           =   990
      End
      Begin VB.Label lblע���̱� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ע���̱�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4770
         TabIndex        =   79
         Top             =   465
         Width           =   720
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
         Left            =   -74880
         TabIndex        =   90
         Top             =   1275
         Width           =   990
      End
      Begin VB.Label lbl�ӳ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ӳ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   96
         Top             =   3705
         Width           =   540
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71820
         TabIndex        =   99
         Top             =   510
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
         Left            =   -74880
         TabIndex        =   91
         Top             =   1665
         Width           =   990
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
         Left            =   -71820
         TabIndex        =   100
         Top             =   1305
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(ƴ��)             (���)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2190
         TabIndex        =   111
         Top             =   2310
         Width           =   2250
      End
      Begin VB.Label lbl��׼�ĺ� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��׼�ĺ�(&W)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   78
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label lblҩƷ��Դ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ����(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   77
         Top             =   3075
         Width           =   990
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
         Left            =   -74880
         TabIndex        =   89
         Top             =   525
         Width           =   990
      End
      Begin VB.Label lbl���ﵥλChild 
         AutoSize        =   -1  'True
         Caption         =   "Ƭ)"
         Height          =   180
         Left            =   7980
         TabIndex        =   109
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label lblסԺ��λChild 
         AutoSize        =   -1  'True
         Caption         =   "Ƭ)"
         Height          =   180
         Left            =   7980
         TabIndex        =   108
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label lbl�ۼ۵�λChild 
         AutoSize        =   -1  'True
         Caption         =   "mg)"
         Height          =   180
         Left            =   7980
         TabIndex        =   107
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lbl���۵�λ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ/Ƭ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   -72375
         TabIndex        =   97
         Top             =   2115
         Width           =   645
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
         Left            =   -74880
         TabIndex        =   94
         Top             =   2880
         Width           =   810
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
         Left            =   -74880
         TabIndex        =   93
         Top             =   2505
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
         Left            =   -72360
         TabIndex        =   98
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   75
         Top             =   2685
         Width           =   540
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɹ��޼�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74880
         TabIndex        =   92
         Top             =   2115
         Width           =   720
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
         Left            =   -74880
         TabIndex        =   95
         Top             =   3270
         Width           =   990
      End
      Begin VB.Label lblҩ���װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6555
         TabIndex        =   87
         Top             =   1965
         Width           =   450
      End
      Begin VB.Label lblҩ�ⵥλ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ�ⵥλ(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4770
         TabIndex        =   86
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lblסԺ��װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1֧="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6555
         TabIndex        =   83
         Top             =   1215
         Width           =   450
      End
      Begin VB.Label lblסԺ��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��λ(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4770
         TabIndex        =   82
         Top             =   1215
         Width           =   990
      End
      Begin VB.Label lbl�����װ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1��="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6555
         TabIndex        =   85
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label lbl���ﵥλ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ﵥλ(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4770
         TabIndex        =   84
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label lbl����ϵ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(1Ƭ="
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6555
         TabIndex        =   81
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lbl�ۼ۵�λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ۵�λ(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4770
         TabIndex        =   80
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʒ������(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   74
         Top             =   2310
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   0
         Top             =   435
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
         Left            =   105
         TabIndex        =   71
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   72
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lbl��ʶ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2595
         TabIndex        =   76
         Top             =   2685
         Width           =   540
      End
      Begin VB.Label lbl��Ʒ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   73
         Top             =   1935
         Width           =   720
      End
   End
   Begin VB.Label lblƷ�� 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ���룺2010303   ͨ�����ƣ�ͷ��߻����   ���ͣ�Ƭ��   ������λ��mg"
      Height          =   180
      Left            =   165
      TabIndex        =   110
      Top             =   75
      Width           =   6120
   End
End
Attribute VB_Name = "frmMediSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ�����Me.tag��ţ��ֱ�Ϊ"5"-����ҩ��"6"-�г�ҩ������lngҩ��ID��ѯȷ��
'   2���༭״̬����Me.stbSpec.Tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ����򴫵ݽ���
'---------------------------------------------------
Public lngҩ��id As Long        '��ǰ�������ҩƷƷ�֣����ⲿ���򴫵ݽ��룻����Ʒ��ȷ������
Public lngҩƷID As Long        '�޸ĺ͡���ѯʱ���ⲿ���򴫵ݽ��룻����ʱ����Ϊ0����ʾ���ݸù���������µĹ��
Public strPrivs As String       '��ǰ�û����еı�����Ȩ��
Public mlng����id As Long      '��¼�������ķ���id

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mblnUsed As Boolean         '�Ƿ���ʹ��
Private mstr���м�¼ As String  '��¼���������м�¼��ֵ
Private mblnOK As Boolean       '��¼ȷ����ť�Ƿ񱻵����
Private mblnCancel As Boolean   '��¼ȡ����ť�Ƿ񱻵����
Private mrs������Ŀ As ADODB.Recordset '��¼ͨ��������������ѯ������������Ŀ�ļ���
Private mstr������Ŀ As String  '��¼�ϴβ�ѯʱ�����ֵ
Private mint�ֶμӳ� As Integer '������ȡϵͳ�����У��Ƿ�ѡ��ʱ��ҩƷ���ֶμӳ���� 0-δ��ѡ��1-��ѡ
Private mrs�ֶμӳ� As ADODB.Recordset '������¼�ֶμӳ������
Private mblnOtherSave As Boolean    '�������水ť�������
Private mintSet���� As Integer  '�ⷿ�������� 0-�ֹ����÷������ԣ�Ĭ��ֵ����1-��ҩ�������2-ҩ���ҩ��������3-ҩ���ҩ����������
Private mbln������Ŀ As Boolean    '��¼������Ŀ�Ƿ񱻵��
Private mdbl�ӳ��� As Double
Private mdbl��۶� As Double

'--Э��ҩƷ������ҩƷ�г���--
Private mint�б�ҩƷ As Integer
Private Const colҩƷ���� As Integer = 1
Private Const col�ۼ۵�λ As Integer = 2
Private Const col��� As Integer = 3
Private Const col���� As Integer = 4
Private Const col������ As Integer = 5
Private Const col������λ As Integer = 6

'--�����޶��г���--
Private Const col�ⷿ As Integer = 1
Private Const col���� As Integer = 2
Private Const col���� As Integer = 3
Private Const col���� As Integer = 4
Private Const col���� As Integer = 5
Private Const col���� As Integer = 6
Private Const col���� As Integer = 7
Private Const col��λ As Integer = 8

Private mlng���볤�� As Long
Private mlng��񳤶� As Long
Private mlng���س��� As Long
Private mlng˵������ As Long
Private mlng���볤�� As Long
Private mint��ѡ�볤�� As Integer
'Private mblnLoad As Boolean      'ֻ��activeһ��

'�Ӳ�������ȡҩƷ�۸�С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��

Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer

Private mlngExpItemMaxLength As Long    '��չ��Ŀ���ݵ���󳤶�
Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSql = "Select A.����, A.���, A.˵��, A.����, B.����, A.��ѡ�� " & _
        " From �շ���ĿĿ¼ A, �շ���Ŀ���� B " & _
        " Where A.ID = B.�շ�ϸĿid And A.ID = 0 And B.���� = 1 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mlng��񳤶� = rsTmp.Fields("���").DefinedSize
    mlng���س��� = rsTmp.Fields("����").DefinedSize
    mlng˵������ = rsTmp.Fields("˵��").DefinedSize
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mint��ѡ�볤�� = rsTmp.Fields("��ѡ��").DefinedSize
    
    txt����.MaxLength = mlng���볤��
    txt���.MaxLength = mlng��񳤶�
    txt����.MaxLength = mlng���س���
    txt˵��.MaxLength = mlng˵������
    txtƴ��.MaxLength = mlng���볤��
    txt���.MaxLength = mlng���볤��
    txt��ѡ��.MaxLength = mint��ѡ�볤��
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboPrepareType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboTemperature_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo��ҩ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
        Exit Sub
    End If
    KeyAscii = 0
End Sub

Private Sub cbo��ΣҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
        Exit Sub
    End If
    KeyAscii = 0
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cboסԺ����_Click()
    If cboסԺ����.ListIndex = 0 Then
        chkסԺ��̬����.Enabled = False
    Else
        chkסԺ��̬����.Enabled = True
    End If
End Sub

Private Sub cboסԺ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo���쵥λ_Click()
    Select Case cbo���쵥λ.ListIndex
    Case 0
        lbl���쵥λ.Caption = txt�ۼ۵�λ.Text & ")"
    Case 1
        lbl���쵥λ.Caption = txtסԺ��λ.Text & ")"
    Case 2
        lbl���쵥λ.Caption = txt���ﵥλ.Text & ")"
    Case 3
        lbl���쵥λ.Caption = txtҩ�ⵥλ.Text & ")"
    End Select
End Sub

Private Sub cbo���쵥λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim i As Integer
    
    On Error GoTo errHandle
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    Else
        strkey = UCase(Chr(KeyAscii))
        If strkey = "" Then Exit Sub
        If mstr������Ŀ <> strkey Then    '�Ѿ��������
            mstr������Ŀ = strkey
            gstrSql = "select id from ������Ŀ where ĩ�� = 1 And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) and (���� like [1] or ���� like [1])"
            Set mrs������Ŀ = zlDatabase.OpenSQLRecord(gstrSql, "������Ŀ", strkey & "%")
            
            If mrs������Ŀ.RecordCount > 0 Then
                For i = 0 To cbo�������.ListCount - 1
                    If Me.cbo�������.ItemData(i) = mrs������Ŀ!ID Then
                        Me.cbo�������.ListIndex = i
                        Exit For
                    End If
                Next
                mrs������Ŀ.MoveNext
            End If
        Else
            If Not mrs������Ŀ.EOF Then
                mrs������Ŀ.MoveNext
                If Not mrs������Ŀ.EOF Then
                    For i = 0 To cbo�������.ListCount - 1
                        If Me.cbo�������.ItemData(i) = mrs������Ŀ!ID Then
                            Me.cbo�������.ListIndex = i
                            Exit For
                        End If
                    Next
                End If
            ElseIf mrs������Ŀ.EOF Then
                mrs������Ŀ.MoveFirst
                If Not mrs������Ŀ.EOF Then
                    For i = 0 To cbo�������.ListCount - 1
                        If Me.cbo�������.ItemData(i) = mrs������Ŀ!ID Then
                            Me.cbo�������.ListIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboҩ�ۼ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cboҩ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cboҩƷ��Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub
Private Sub chk�׵���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub
Private Sub chk�����ɹ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkCondition_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkDosage_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkGMP��֤_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk�ǳ���ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk���ηѱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkЧ��_Click()
    On Error Resume Next
    Me.txtЧ��.Enabled = (chkЧ��.Value = 1)
    If Me.txtЧ��.Enabled = False Then
        Me.txtЧ��.Text = 0
    Else
        If Val(Me.txtЧ��.Text) = 0 Then Me.txtЧ��.Text = 24
    End If
    If Me.chkЧ��.Value = 1 Then Me.txtЧ��.SetFocus
End Sub

Private Sub chkЧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Me.txtЧ��.Enabled = True Then
            Me.txtЧ��.SetFocus
        Else
            If txtЧ��.Enabled = True Then
                Call OS.PressKey(vbKeyTab)
            Else
                If stbSpec.TabVisible(2) = True Then
                    stbSpec.Tab = 2
                    cboTemperature.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkҩ��_Click()
    Dim blnEnable As Boolean
    Dim rsTem As ADODB.Recordset
    
    On Error GoTo errHandle
    '��ҩ�������ǰ���£����ҩ��û�п�棬����������Ƿ����
    gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
             " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '%ҩ��' Or B.�������� Like '%�Ƽ���')"
    Set rsTem = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
    
    With rsTem
        blnEnable = True
        If .Fields(0).Value <> 0 Then
            blnEnable = False
        End If
    End With
    If Me.chkҩ��.Value = 0 Then
        Me.chkҩ��.Value = 0: Me.chkҩ��.Enabled = False
        Me.chkЧ��.Value = 0: Me.chkЧ��.Enabled = False
        Me.txtЧ��.Text = 0: Me.txtЧ��.Enabled = False
    Else
        Me.chkҩ��.Enabled = True
        Me.chkЧ��.Enabled = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If stbSpec.TabVisible(2) = True And chkҩ��.Enabled = False Then
        stbSpec.Tab = 2
        cboTemperature.SetFocus
    Else
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkסԺ��̬����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cmbStationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim strTemp As String
    
    If mblnOtherSave = False Then
        strTemp = txt����.Text & "|" & txt��λ�� & "|" & txt���.Text & "|" & txt����.Text & "|" & txt��Ʒ��.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & _
                        txt������.Text & "|" & txt��ʶ��.Text & "|" & cboҩƷ��Դ.Text & "|" & txt��ͬ��λ.Text & "|" & txt˵��.Text & "|" & cbo��ҩ����.Text & "|" & _
                        cmbStationNo.Text & "|" & txt��׼�ĺ�.Text & "|" & txtע���̱�.Text & "|" & txt�ۼ۵�λ.Text & "|" & txt����ϵ��.Text & "|" & txtסԺ��λ.Text & "|" & _
                        txtסԺ��װ.Text & "|" & txt���ﵥλ.Text & "|" & txt�����װ.Text & "|" & txtҩ�ⵥλ.Text & "|" & txtҩ���װ.Text & "|" & cbo���쵥λ.Text & "|" & txt���췧ֵ.Text & "|" & _
                        txt��ѡ��.Text & "|" & txt����.Text & "|" & cboҩ������.Text & "|" & txt�ɱ��۸�.Text & "|" & txt��ǰ�ۼ�.Text & "|" & txtָ������.Text & "|" & txt����.Text & "|" & txt�����.Text & "|" & _
                        txtָ���ۼ�.Text & "|" & txt�ӳ���.Text & "|" & cbo�������.Text & "|" & txt������Ŀ.Text & "|" & cboҩ�ۼ���.Text & "|" & _
                        chk���ηѱ�.Value & "|" & cbo��������.Text & "|" & cbo�������.Text & "|" & cboסԺ����.Text & "|" & cboBasicDrug.Text & "|" & chkסԺ��̬����.Value & "|" & _
                        chkGMP��֤.Value & "|" & chk�ǳ���ҩ.Value & "|" & chkҩ��.Value & "|" & chkҩ��.Value & "|" & chkЧ��.Value & "|" & txtЧ��.Text & "|" & cboTemperature.Text & "|" & chkCondition.Value & "|" & _
                        cboPrepareType.Text & "|" & chkDosage.Value & "|" & cbo�������.Text & "|" & txtDDDֵ.Text & "|" & cbo��ΣҩƷ.Text & "|" & chk�׵���.Value & "|" & chk�����ɹ�.Value
        If strTemp <> mstr���м�¼ Then
            mblnCancel = True
            If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
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
End Sub

Private Sub cmdOk_Click()
    Dim dbl��ǰ�ۼ� As Double, dblָ���ۼ� As Double, dbl�ɱ��۸� As Double
    Dim blnPackerReturn As Boolean
    Dim strվ�� As String
    Dim blnTran As Boolean
    Dim strItems As String
    Dim n As Integer
    Dim rsPrice As New ADODB.Recordset
    
    mblnOK = True
    '�����ҳ����������Ƿ���ȷ
    strTemp = IIf(glngSys \ 100 <> 8, "ҩ��", "�ɹ�")
    If Trim(Me.txt����.Text) = "" Then MsgBox "��������룡", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt����.Text, vbFromUnicode)) > mlng���볤�� Then MsgBox "���볬��(���" & mlng���볤�� & "���ַ�)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.txt���.Text) = "" Then MsgBox "��������", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt���.Text, vbFromUnicode)) > mlng��񳤶� Then MsgBox "��񳬳�(���" & mlng��񳤶� & "���ַ���" & Int(mlng��񳤶� / 2) & "������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt��Ʒ��.Text, vbFromUnicode)) > 40 Then MsgBox "��Ʒ������(���40���ַ���20������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt��Ʒ��.SetFocus: Exit Sub
    
    
    
    If Trim(Me.txt�ۼ۵�λ.Text) = "" Then MsgBox "�������ۼ۵�λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�ۼ۵�λ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt�ۼ۵�λ.Text, vbFromUnicode)) > 8 Then MsgBox "�ۼ۵�λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�ۼ۵�λ.SetFocus: Exit Sub
    If Val(Me.txt����ϵ��.Text) = 0 Then MsgBox "����ϵ������(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����ϵ��.SetFocus: Exit Sub
    If Val(Me.txt����ϵ��.Text) >= 100000 Then MsgBox "����ϵ���������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����ϵ��.SetFocus: Exit Sub
    
    If Trim(Me.txt���ﵥλ.Text) = "" Then MsgBox "���������ﵥλ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���ﵥλ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txt���ﵥλ.Text, vbFromUnicode)) > 8 Then MsgBox "���ﵥλ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���ﵥλ.SetFocus: Exit Sub
    If Val(Me.txt�����װ.Text) = 0 Then MsgBox "�����װ����(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�����װ.SetFocus: Exit Sub
    If Val(Me.txt�����װ.Text) >= 100000 Then MsgBox "�����װ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt�����װ.SetFocus: Exit Sub
    
    If Trim(Me.txtסԺ��λ.Text) = "" Then MsgBox "������סԺ��λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtסԺ��λ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txtסԺ��λ.Text, vbFromUnicode)) > 8 Then MsgBox "סԺ��λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtסԺ��λ.SetFocus: Exit Sub
    If Val(Me.txtסԺ��װ.Text) = 0 Then MsgBox "סԺ��װ����(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtסԺ��װ.SetFocus: Exit Sub
    If Val(Me.txtסԺ��װ.Text) >= 100000 Then MsgBox "סԺ��װ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtסԺ��װ.SetFocus: Exit Sub
    
    If Trim(Me.txtҩ�ⵥλ.Text) = "" Then MsgBox "������" & strTemp & "��λ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ�ⵥλ.SetFocus: Exit Sub
    If LenB(StrConv(Me.txtҩ�ⵥλ.Text, vbFromUnicode)) > 8 Then MsgBox strTemp & "��λ����(���8���ַ���4������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ�ⵥλ.SetFocus: Exit Sub
    If Val(Me.txtҩ���װ.Text) = 0 Then MsgBox strTemp & "��װ����(����Ϊ0)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ���װ.SetFocus: Exit Sub
    If Val(Me.txtҩ���װ.Text) >= 100000 Then MsgBox strTemp & "��װ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txtҩ���װ.SetFocus: Exit Sub
    If Trim(txt�ͻ���λ.Text) <> "" And Trim(txt�ͻ���װ.Text) = "" Then
        MsgBox "���ͻ���λ����£��ͻ���װ����Ϊ�գ�", vbInformation, gstrSysName
        txt�ͻ���װ.SetFocus
        txt�ͻ���װ.SelStart = 0
        txt�ͻ���װ.SelLength = 100
        Exit Sub
    End If
    If Trim(txt�ͻ���װ.Text) <> "" And IsNumeric(txt�ͻ���װ.Text) = False Then
        MsgBox "�ͻ���װֻ�������֣����������룡", vbInformation, gstrSysName
        txt�ͻ���װ.SetFocus
        txt�ͻ���װ.SelStart = 0
        txt�ͻ���װ.SelLength = 100
        Exit Sub
    End If
    
    If LenB(StrConv(Me.txtע���̱�.Text, vbFromUnicode)) > 50 Then
        MsgBox "ע���̱곬�������50���ַ���25�����֣�", vbInformation, gstrSysName
        Me.stbSpec.Tab = 0
        txtע���̱�.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(Me.txt��ѡ��.Text, vbFromUnicode)) > mint��ѡ�볤�� Then
        MsgBox "��ѡ�볬��(���" & mint��ѡ�볤�� & "���ַ�)��", vbInformation, gstrSysName
        Me.stbSpec.Tab = 0
        txt��ѡ��.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.txt����.Text) <> "" And Not IsNumeric(Me.txt����.Text) Then MsgBox "����ֻ��Ϊ���֣�", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    
    If LenB(StrConv(Me.txt����.Text, vbFromUnicode)) > mlng���س��� Then MsgBox "�����̳���(���" & mlng���س��� & "���ַ���" & Int(mlng���س��� / 2) & "������)��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt����.SetFocus: Exit Sub
    
    If Val(Me.txt���췧ֵ.Text) < 0 Then MsgBox strTemp & "���췧ֵ����С���㣡", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���췧ֵ.SetFocus: Exit Sub
    If Val(Me.txt���췧ֵ.Text) >= 100000 Then MsgBox strTemp & "���췧ֵ�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 0: Me.txt���췧ֵ.SetFocus: Exit Sub
    
    If Val(Me.txtָ������.Text) = 0 And mblnUsed = True Then
        MsgBox "������" & IIf(mint�б�ҩƷ = 1, "�б�۸�", "ָ������") & "��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus: Exit Sub
    End If
    If Val(Me.txtָ������.Text) > 1000000 Then
        MsgBox IIf(mint�б�ҩƷ = 1, "�б�۸�", "ָ������") & "�������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1
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
    If Val(Me.txt����.Text) = 0 Then MsgBox "��������ʣ�", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����.SetFocus: Exit Sub
    If Val(Me.txt����.Text) > 100 Then MsgBox "���ʳ������ֵ��", vbInformation, gstrSysName: Me.stbSpec.Tab = 1: Me.txt����.SetFocus: Exit Sub
    
    If LenB(StrConv(Me.cboBasicDrug.Text, vbFromUnicode)) > 30 Then
        MsgBox "����ҩ�ﳬ�������30���ַ���15�����֣�", vbInformation, gstrSysName
        Me.stbSpec.Tab = 1
        cboBasicDrug.SetFocus
        Exit Sub
    End If
    
    If Val(Me.txt�ӳ���.Text) > 1000000 Then
        MsgBox "��ǰ�ӳ��ʳ������ֵ��", vbInformation, gstrSysName
        Me.stbSpec.Tab = 1
        If Me.txt�ӳ���.Enabled Then Me.txt�ӳ���.SetFocus
        Exit Sub
    End If
    
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
    
    '�����չ��Ŀ���ݳ���
    If stbSpec.TabVisible(3) = True Then
        With vsfItem
            For n = 1 To .Rows - 1
                If LenB(StrConv(Trim(.TextMatrix(n, .ColIndex("����"))), vbFromUnicode)) > mlngExpItemMaxLength Then
                    MsgBox "��չ��Ŀ���ݳ���(���" & mlngExpItemMaxLength & "���ַ���" & Int(mlngExpItemMaxLength) / 2 & "������)��", vbInformation, gstrSysName
                    Me.stbSpec.Tab = 3
                    
                    .Row = n
                    .Col = .ColIndex("����")
                    Exit Sub
                End If
            Next
        End With
    End If
    
    '���۹���ģʽ���۸�
    If chk����.Enabled = True And chk����.Value = 1 Then
        If Me.stbSpec.Tag = "����" Then
            If Val(Me.txt��ǰ�ۼ�.Text) <> Val(Me.txt�ɱ��۸�.Text) Then
                MsgBox "�������۹���ģʽʱ���ۼۺͳɱ���Ҫһ�£�", vbInformation, gstrSysName
                Me.stbSpec.Tab = 0
                If Me.txt��ǰ�ۼ�.Enabled Then Me.txt��ǰ�ۼ�.SetFocus
                Exit Sub
            End If
        ElseIf txt��ǰ�ۼ�.Enabled = True And txt�ɱ��۸�.Enabled = True Then
            If Val(Me.txt��ǰ�ۼ�.Text) <> Val(Me.txt�ɱ��۸�.Text) Then
                MsgBox "�������۹���ģʽʱ���ۼۺͳɱ���Ҫһ�£�", vbInformation, gstrSysName
                Me.stbSpec.Tab = 0
                If Me.txt��ǰ�ۼ�.Enabled Then Me.txt��ǰ�ۼ�.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Not CheckUnit Then Exit Sub
    If Not CheckRequest Then Exit Sub
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    If Me.stbSpec.Tag = "�޸�" Then
        If cboҩ������.Tag = 0 And Me.cboҩ������.ItemData(Me.cboҩ������.ListIndex) = 1 Then
            If MsgBox("ҩƷ�۸������ɡ����ۡ���Ϊ�ˡ�ʱ�ۡ����Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        ElseIf cboҩ������.Tag = 1 And Me.cboҩ������.ItemData(Me.cboҩ������.ListIndex) = 0 Then
            If txt��ǰ�ۼ�.Enabled = False Then
                gstrSql = "Select b.�ϴ��ۼ� as �۸�, b.ҩ���װ" & vbNewLine & _
                                "From ҩƷ��� B" & vbNewLine & _
                                "Where b.ҩƷid=[1]"
        
                Set rsPrice = zlDatabase.OpenSQLRecord(gstrSql, "ʱ��ת����", lngҩƷID)
                If IsNull(rsPrice!�۸�) Then
                    gstrSql = "Select a.�ּ� as �۸�, b.ҩ���װ" & vbNewLine & _
                                "From �շѼ�Ŀ A, ҩƷ��� B" & vbNewLine & _
                                "Where a.�շ�ϸĿid = b.ҩƷid And a.�շ�ϸĿid =[1] And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And" & vbNewLine & _
                                "      �䶯ԭ�� = 1"
                    
                    Set rsPrice = zlDatabase.OpenSQLRecord(gstrSql, "ʱ��ת����", lngҩƷID)
                End If
                If MsgBox("ҩƷ�۸������ɡ�ʱ�ۡ���Ϊ�ˡ����ۡ����ۼ�Ϊ" & zlStr.FormatEx(rsPrice!�۸� * rsPrice!ҩ���װ, mintPriceDigit, , True) & "���Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("ҩƷ�۸������ɡ�ʱ�ۡ���Ϊ�ˡ����ۡ����ۼ�Ϊ" & txt��ǰ�ۼ�.Text & "���Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    '------------------------------------------
    '���ݱ���
    gstrSql = "'" & Me.txt����.Text & "','" & MoveSpecialChar(Me.txt���.Text) & "','" & MoveSpecialChar(Me.txt����.Text, False) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Me.txt��Ʒ��.Text) & "','" & MoveSpecialChar(Me.txtƴ��.Text) & "','" & MoveSpecialChar(Me.txt���.Text) & "','" & MoveSpecialChar(Me.txt������.Text) & "'"
    gstrSql = gstrSql & ",'" & Me.txt��ʶ��.Text & "','" & Mid(Me.cboҩƷ��Դ.Text, InStr(1, Me.cboҩƷ��Դ.Text, "-") + 1) & "','" & MoveSpecialChar(Me.txt��׼�ĺ�.Text) & "','" & MoveSpecialChar(Me.txtע���̱�.Text) & "'"
    gstrSql = gstrSql & ",'" & Me.txt�ۼ۵�λ.Text & "'," & Val(Me.txt����ϵ��.Text)
    gstrSql = gstrSql & ",'" & Me.txt���ﵥλ.Text & "'," & Val(Me.txt�����װ.Text)
    gstrSql = gstrSql & ",'" & Me.txtסԺ��λ.Text & "'," & Val(Me.txtסԺ��װ.Text)
    gstrSql = gstrSql & ",'" & Me.txtҩ�ⵥλ.Text & "'," & Val(Me.txtҩ���װ.Text)
    gstrSql = gstrSql & "," & cbo���쵥λ.ListIndex + 1  '���쵥λ��1-���۵�λ;2-סԺ��λ;3-���ﵥλ;4-ҩ�ⵥλ��
    gstrSql = gstrSql & "," & Val(txt���췧ֵ.Tag)       'ʼ�������۵�λ����
    gstrSql = gstrSql & "," & Me.cboҩ������.ItemData(Me.cboҩ������.ListIndex)
    If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
        dblָ���ۼ� = FormatEx(Val(txtָ���ۼ�.Text) / Val(txtҩ���װ.Text), gtype_MaxDigits.dig_���ۼ�)
        dbl��ǰ�ۼ� = FormatEx(Val(txt��ǰ�ۼ�.Text) / Val(txtҩ���װ.Text), gtype_MaxDigits.dig_���ۼ�)
        dbl�ɱ��۸� = FormatEx(Val(txt�ɱ��۸�.Text) / Val(txtҩ���װ.Text), gtype_MaxDigits.dig_�ɱ���)
        gstrSql = gstrSql & "," & FormatEx(Val(Me.txtָ������.Text) / Val(Me.txtҩ���װ), gtype_MaxDigits.dig_�ɱ���)
    Else
        dbl��ǰ�ۼ� = FormatEx(Val(txt��ǰ�ۼ�.Text), gtype_MaxDigits.dig_���ۼ�)
        dblָ���ۼ� = FormatEx(Val(txtָ���ۼ�.Text), gtype_MaxDigits.dig_���ۼ�)
        dbl�ɱ��۸� = FormatEx(Val(txt�ɱ��۸�.Text), gtype_MaxDigits.dig_�ɱ���)
        gstrSql = gstrSql & "," & FormatEx(Val(Me.txtָ������.Text), gtype_MaxDigits.dig_�ɱ���)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt����.Text) & "," & dblָ���ۼ� & "," & Val(Trim(txt�ӳ���.Text)) & "," & 0
    gstrSql = gstrSql & ",'" & Mid(Me.cboҩ�ۼ���.Text, InStr(1, Me.cboҩ�ۼ���.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex) & "," & Me.chkGMP��֤.Value & "," & mint�б�ҩƷ & "," & Me.chk���ηѱ�.Value
    gstrSql = gstrSql & "," & Me.cboסԺ����.ItemData(Me.cboסԺ����.ListIndex)
    gstrSql = gstrSql & "," & Me.chkҩ�� & "," & Me.chkҩ�� & "," & IIf(Me.chkЧ��.Value = 0, 0, Val(Me.txtЧ��.Text))
    gstrSql = gstrSql & "," & 100
    
    If Me.stbSpec.Tag = "����" Then
        lngҩƷID = Sys.NextId("�շ���ĿĿ¼")
        gstrSql = "zl_��ҩ���_Insert(" & lngҩ��id & "," & lngҩƷID & "," & gstrSql
        gstrSql = gstrSql & "," & dbl�ɱ��۸� & "," & dbl��ǰ�ۼ� & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex) & ""
    Else
        gstrSql = "zl_��ҩ���_Update(" & lngҩƷID & "," & gstrSql
        gstrSql = gstrSql & "," & dbl�ɱ��۸� & "," & dbl��ǰ�ۼ� & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex) & ""
    End If
    
    gstrSql = gstrSql & "," & ZVal(Split(Me.txt��ͬ��λ.Tag, "|")(0)) & ",'"
    gstrSql = gstrSql & MoveSpecialChar(Me.txt˵��.Text) & "'" & ","
    gstrSql = gstrSql & IIf(Me.chkסԺ��̬����.Enabled = False, 0, chkסԺ��̬����.Value) & ",'"
    gstrSql = gstrSql & cbo��ҩ����.Text & "','"
    gstrSql = gstrSql & MoveSpecialChar(txt��ѡ��.Text) & "',"
    gstrSql = gstrSql & 0
    If Trim(Me.cboBasicDrug.Text) = "" Then
        gstrSql = gstrSql & ",null,"
    Else
        gstrSql = gstrSql & ",'" & Trim(Me.cboBasicDrug.Text) & "',"
    End If
    gstrSql = gstrSql & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ��) & ","
    gstrSql = gstrSql & chk�ǳ���ҩ.Value & ","
    
    '��ҺҩƷ����
    gstrSql = gstrSql & IIf(cboTemperature.ListIndex = 0 Or cboTemperature.ListIndex = -1, "Null", "'" & cboTemperature.Text & "'") & ","
    gstrSql = gstrSql & chkCondition.Value & ","
    gstrSql = gstrSql & IIf(cboPrepareType.ListIndex = 0 Or cboPrepareType.ListIndex = -1, "Null", "'" & cboPrepareType.Text & "'") & ","
    gstrSql = gstrSql & chkDosage.Value & ","
    gstrSql = gstrSql & Val(Me.txt����.Text) & ","
    gstrSql = gstrSql & "'" & txt������Ŀ.Text & "'"
    gstrSql = gstrSql & "," & Me.cbo�������.ItemData(Me.cbo�������.ListIndex) & ","
    gstrSql = gstrSql & Val(txtDDDֵ.Text) & ","
    gstrSql = gstrSql & Val(Mid(cbo��ΣҩƷ.Text, 1, 1))
    gstrSql = gstrSql & ",'" & Trim(txt�ͻ���λ.Text) & "'"
    gstrSql = gstrSql & "," & IIf(Trim(txt�ͻ���װ.Text) = "", "Null", Val(Trim(txt�ͻ���װ.Text)) * Val(txtҩ���װ.Text))
    gstrSql = gstrSql & ",'" & Trim(txtNotice.Text) & "',"
    gstrSql = gstrSql & chk��ҩ.Value & ","
    gstrSql = gstrSql & chk����.Value & ","
    gstrSql = gstrSql & "'" & MoveSpecialChar(Me.txt��λ��.Text) & "',"
    gstrSql = gstrSql & chk�׵���.Value
    gstrSql = gstrSql & "," & chk�����ɹ�.Value
    gstrSql = gstrSql & " )"
  
    err = 0: On Error GoTo ErrHand
    
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    '������չ��Ŀ
    If stbSpec.TabVisible(3) = True Then
        With vsfItem
            For n = 1 To .Rows - 1
                If Trim(.TextMatrix(n, .ColIndex("����"))) <> "" Then
                    strItems = IIf(strItems = "", "", strItems & "|") & .TextMatrix(n, .ColIndex("��Ŀ")) & "," & Trim(.TextMatrix(n, .ColIndex("����")))
                End If
            Next
        End With
        
        If strItems <> "" Then
            gstrSql = "Zl_ҩƷ�����չ��Ϣ_Update("
            'ҩƷID
            gstrSql = gstrSql & lngҩƷID
            '��Ŀ���ݴ�
            gstrSql = gstrSql & "," & "'" & strItems & "'"
            gstrSql = gstrSql & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSql, "����ҩƷ�����չ��Ϣ")
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    '���۹���
    If chk����.Enabled = True And chk����.Value = 1 Then
        If CheckPriceAdjust(lngҩƷID, 0, -1) = False Then
            MsgBox "��ҩƷ���������۹������ۼۺͳɱ��۲�һ�£���ע����ۣ�", vbInformation, gstrSysName
        End If
    End If
    
    '�������޸ĵ�ҩƷ��Ϣͬ���ϴ�����ƽ̨
    UploadDrugInfo lngҩƷID
    
    If Me.stbSpec.Tag = "����" Then
        'Val(zldatabase.GetPara("�������ģʽ", glngSys, 1023, 0)) = 0
        If ActiveControl Is cmdOK Then  '��ͨģʽ
            Unload Me
        ElseIf ActiveControl Is cmdSaveAddSpec Then   '�������ӹ��ģʽ
            Call frmMediLists.zlRefRecords(lngҩ��id)
            Call Form_Activate
            Me.stbSpec.Tab = 0: Me.txt���.SetFocus
        ElseIf ActiveControl Is cmdSaveAddItem Then
            With frmMediItem
                .Tag = IIf(Me.Tag = "5", 1, 2)
                .cmdCancel.Tag = "����"
                .lng����id = mlng����id
                .lngҩ��id = 0
                .strPrivs = gstrPrivs
                .lng������ = 0
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
    
    On Error GoTo errHandle
    lblStationNo.Visible = False
    cmbStationNo.Visible = False
    
    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select ���,���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "վ���ѯ")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Call cmdOk_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    mblnOtherSave = True
    Call cmdOk_Click
End Sub

Private Sub cmd����_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmd����_Click()
    On Error GoTo errHandle
    Dim strSql As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    
    mbln������Ŀ = True
    strSql = "Select ���� as id,�ϼ� as �ϼ�id, ����, ����, ĩ�� From ������Ŀ Start With �ϼ� Is Null Connect By Prior ���� = �ϼ�"
    blnRe = frmTreeLeafSel.ShowTree(strSql, strID, str����, "������Ŀ")
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        lbl������Ŀ.Tag = strID
        txt������Ŀ.Text = str����
        stbSpec.Tab = 1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim vRect As RECT, blnCancel As Boolean

    vRect = zlControl.GetControlRect(txt����.hwnd)

    On Error GoTo errHandle
    
    gstrSql = "Select ���� as id,����,���� From ҩƷ������ Order By ���� "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "cmd����_Click", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
        
    If blnCancel = True Then txt����.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsTemp.State = 0 Then
        MsgBox "���ʼ��ҩƷ�����̣��ֵ������", vbInformation, gstrSysName
        Me.txt����.Tag = "": Me.txt����.SetFocus: Exit Sub
        Exit Sub
    End If
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    End If

    txt����.SetFocus
    txt����.Text = rsTemp!����
    txt����.Tag = txt����.Text
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd��ͬ��λ_Click()
    Dim vRect As RECT, blnCancel As Boolean

    vRect = zlControl.GetControlRect(txt��ͬ��λ.hwnd)
    
    On Error GoTo errHandle
    
    gstrSql = "Select ����,����,����,ID" & _
    " From ��Ӧ��" & _
    " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
    " Order By ���� "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "cmd��ͬ��λ_Click", False, "", "", False, False, _
    True, vRect.Left, vRect.Top, 300, blnCancel, False, True)

    If blnCancel = True Then txt��ͬ��λ.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsTemp.State = 0 Then
        MsgBox "���ʼ��ҩƷ��Ӧ�̣��ֵ������", vbInformation, gstrSysName
        Me.txt��ͬ��λ.Tag = "": Me.txt��ͬ��λ.SetFocus: Exit Sub
        Exit Sub
    End If
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    End If

    txt��ͬ��λ.SetFocus
    Me.txt��ͬ��λ.Text = rsTemp!����
    Me.txt��ͬ��λ.Tag = rsTemp!ID & "|" & rsTemp!����
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    Dim blnExit As Boolean
    Dim strMsg As String
    Dim i As Integer
    Dim rs����� As ADODB.Recordset
    Dim str�ͻ���λ As String
    Dim dbl�ͻ���װ As Double
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbln������Ŀ = True Then Exit Sub
    If Me.stbSpec.Tag <> "����" Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    
    mintSet���� = Val(zlDatabase.GetPara("ҩƷ���������Զ�����", glngSys, 1023, 0))
    '----------������ϵ�ж�-------------------------------------
    If Me.cboҩƷ��Դ.ListCount = 0 Then
        strMsg = "δ����ҩƷ��Դ���ࣨ�ֵ������"
        blnExit = True
    End If
    If Me.cbo��������.ListCount = 0 And Not blnExit Then
        strMsg = "δ��������ҩƷ��ҽ�����ͣ��ֵ������"
        blnExit = True
    End If
    If Me.cbo�������.ListCount = 0 And Not blnExit Then
        strMsg = "δ������ϸ��������Ŀ��"
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
    
    txt��λ��.MaxLength = Val(zlDatabase.GetPara("��λ��", glngSys, 1023, 20))
    txt������.MaxLength = Val(zlDatabase.GetPara("������", glngSys, 1023, 7))
'    If mblnLoad = True Then Exit Sub
    '----------ҩƷƷ��ʶ��-------------------------------------
    gstrSql = "select I.���,I.����,I.����,I.���㵥λ,T.ҩƷ����" & _
            " from ������ĿĿ¼ I,ҩƷ���� T" & _
            " where I.ID=T.ҩ��ID and I.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
    
    With rsTemp
        If !��� = "5" Then
            Me.Tag = "5": Me.Caption = "����ҩ���༭"
            Me.lbl�������.Tag = zlDatabase.GetPara("����ҩ������Ŀ", glngSys, 1023, False)
        Else
            Me.Tag = "6": Me.Caption = "�г�ҩ���༭"
            Me.lbl�������.Tag = zlDatabase.GetPara("�г�ҩ������Ŀ", glngSys, 1023, False)
        End If
        If Me.stbSpec.Tag = "����" And Val(Me.lbl�������.Tag) = 0 Then
            MsgBox "û�����á�" & IIf(Me.Tag = "5", "����ҩ", "�г�ҩ") & "����Ӧ��������Ŀ�����ز������ã���", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        For intCount = 0 To Me.cbo�������.ListCount - 1
            If Me.cbo�������.ItemData(intCount) = Val(Me.lbl�������.Tag) Then
                Me.cbo�������.ListIndex = intCount: Exit For
            End If
        Next
        
        Me.lblƷ��.Caption = "ҩƷ���룺" & !���� & _
                "   ͨ�����ƣ�" & !���� & _
                "   ���ͣ�" & IIf(IsNull(!ҩƷ����), "", !ҩƷ����) & _
                "   ������λ��" & IIf(IsNull(!���㵥λ), "", !���㵥λ)
        Me.lblƷ��.Tag = !����
        Me.lbldddֵ.Caption = IIf(IsNull(!���㵥λ), "", !���㵥λ)
        Me.lbl�ۼ۵�λChild.Caption = IIf(IsNull(!���㵥λ), "", !���㵥λ)
        
        Me.lbl���۵�λ(0).Tag = Val(zlDatabase.GetPara(29, glngSys))
        
        mintCostDigit = GetDigit(1, 1, IIf(Me.lbl���۵�λ(0).Tag = 0, 1, 4))
        mintPriceDigit = GetDigit(1, 2, IIf(Me.lbl���۵�λ(0).Tag = 0, 1, 4))
        
        mintSaleCostDigit = GetDigit(1, 1, 1)
        mintSalePriceDigit = GetDigit(1, 2, 1)

    End With
    
    '----------����װ��-------------------------------------
    'ֻҪ����lngҩƷID��������ʲô״̬�����ù����Ϣ
    gstrSql = "select I.����,S.��λ��,I.���,I.����,S.��ʶ��,S.ҩƷ��Դ,S.��׼�ĺ�,S.ע���̱�,S.����," & _
            "        I.���㵥λ,S.����ϵ��,S.���ﵥλ,S.�����װ,S.סԺ��λ,S.סԺ��װ,S.ҩ�ⵥλ,S.ҩ���װ,s.�ͻ���λ,s.�ͻ���װ," & _
            "        I.�Ƿ���,S.ָ��������,S.����,S.ָ�����ۼ�,S.�ӳ���,S.�ɱ���,S.�б�ҩƷ,s.dddֵ,S.GMP��֤,S.����ҩ��, " & _
            "        S.ҩ�ۼ���,i.������Ŀ,I.��������,I.�������,I.���ηѱ�,S.���쵥λ,S.���췧ֵ," & _
            "        S.סԺ�ɷ����,S.��̬���� as סԺ��̬����,S.����ɷ����,S.ҩ�����,S.ҩ������,S.���Ч��,S.��ҩ����,I.��ѡ��," & _
            "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,S.��ͬ��λid,G.���� ��ͬ��λ,I.˵��,I.վ��,S.��ֵ˰��,S.�Ƿ񳣱�, " & _
            "        Nvl(a.�洢�¶�, 0) As �洢�¶�, Nvl(a.�洢����, 0) As �洢����, Nvl(a.��ҩ����, 0) As ��ҩ����,Nvl(a.�Ƿ�������,0) As �Ƿ�������,s.��ΣҩƷ, " & _
            "        A.��Һע������,s.�Ƿ��ҩ,s.�Ƿ����۹���,s.�Ƿ���������,s.�Ƿ�����ɹ� " & _
            " from �շ���ĿĿ¼ I,ҩƷ��� S,��ҺҩƷ���� A,(Select Id,���� From ��Ӧ�� Where ĩ�� = 1 And substr(����,1,1) = '1' And " & _
            " ����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) G " & _
            " where I.ID=S.ҩƷID and G.id(+)=S.��ͬ��λid And i.Id = a.ҩƷid(+) and I.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt����.Text = !����
            Me.txt��λ��.Text = NVL(!��λ��)
            Me.txt���.Text = IIf(IsNull(!���), "", !���)
            Me.txt����.Text = IIf(IsNull(!����), "", !����)
            Me.txt��ͬ��λ.Text = IIf(IsNull(!��ͬ��λ), "", !��ͬ��λ)
            Me.txt��ͬ��λ.Tag = IIf(IsNull(!��ͬ��λid), "|", !��ͬ��λid & "|" & !��ͬ��λ)
            Me.txt��ʶ��.Text = IIf(IsNull(!��ʶ��), "", !��ʶ��)
            Me.txt˵��.Text = IIf(IsNull(!˵��), "", !˵��)
            Me.txt��ѡ��.Text = IIf(IsNull(!��ѡ��), "", !��ѡ��)

            For intCount = 0 To Me.cboҩƷ��Դ.ListCount - 1
                If Mid(Me.cboҩƷ��Դ.List(intCount), InStr(1, Me.cboҩƷ��Դ.List(intCount), "-") + 1) = IIf(IsNull(!ҩƷ��Դ), "", !ҩƷ��Դ) Then
                    Me.cboҩƷ��Դ.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt��׼�ĺ�.Text = IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)
            Me.txtע���̱�.Text = IIf(IsNull(!ע���̱�), "", !ע���̱�)
            Me.txt����.Text = IIf(IsNull(!����), "", Format(!����, "0.00000"))
            Me.txt�ۼ۵�λ.Text = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            txtDDDֵ.Text = IIf(IsNull(!DDDֵ), "", !DDDֵ)
            Me.lbl���ﵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
            Me.lblסԺ��λChild.Caption = Me.txt�ۼ۵�λ & ")"
            Me.lblҩ�ⵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
            Me.lbl����ϵ��.Caption = "(1" & Me.txt�ۼ۵�λ & "="
            Me.txt����ϵ��.Text = FormatEx(IIf(IsNull(!����ϵ��), 1, !����ϵ��), 5, , False)
            Me.txt���ﵥλ.Text = IIf(IsNull(!���ﵥλ), "", !���ﵥλ)
            Me.lbl�����װ.Caption = "(1" & Me.txt���ﵥλ.Text & "="
            Me.txt�����װ.Text = FormatEx(IIf(IsNull(!�����װ), 1, !�����װ), 5, , False)
            Me.txtסԺ��λ.Text = IIf(IsNull(!סԺ��λ), "", !סԺ��λ)
            Me.lblסԺ��װ.Caption = "(1" & Me.txtסԺ��λ.Text & "="
            Me.txtסԺ��װ.Text = FormatEx(IIf(IsNull(!סԺ��װ), 1, !סԺ��װ), 5, , False)
            Me.txtҩ�ⵥλ.Text = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
            Me.lblҩ���װ.Caption = "(1" & Me.txtҩ�ⵥλ.Text & "="
            Me.txtҩ���װ.Text = FormatEx(IIf(IsNull(!ҩ���װ), 1, !ҩ���װ), 5, , False)
            str�ͻ���λ = IIf(IsNull(!�ͻ���λ), "", !�ͻ���λ)
            dbl�ͻ���װ = IIf(IsNull(!�ͻ���λ), 0, !�ͻ���װ)
            Me.txt�ͻ���λ.Text = str�ͻ���λ
            Me.txt�ͻ���װ.Text = IIf(dbl�ͻ���װ = 0, "", FormatEx(dbl�ͻ���װ / !ҩ���װ, 1, , True))
            lbl�ͻ���λchild.Caption = txtҩ�ⵥλ.Text
            Me.txtNotice.Text = NVL(!��Һע������)
            
            Me.cbo���쵥λ.ListIndex = (NVL(!���쵥λ, 1) - 1)
            For i = 0 To cbo��ҩ����.ListCount
                If cbo��ҩ����.List(i) = !��ҩ���� Then
                    Me.cbo��ҩ����.ListIndex = i
                    Exit For
                ElseIf IsNull(!��ҩ����) Then
                    Me.cbo��ҩ����.ListIndex = 0
                End If
            Next
            
            For i = 0 To cbo��ΣҩƷ.ListCount
                If Val(Mid(cbo��ΣҩƷ.List(i), 1, 1)) = IIf(IsNull(!��ΣҩƷ), 0, !��ΣҩƷ) Then
                    Me.cbo��ΣҩƷ.ListIndex = i
                    Exit For
                ElseIf IsNull(!��ΣҩƷ) Then
                    Me.cbo��ΣҩƷ.ListIndex = 0
                End If
            Next
            
            SetStationNo IIf(IsNull(!վ��), "", !վ��)
            
            Select Case NVL(!���쵥λ, 1)
            Case 1 '����
                Me.txt���췧ֵ.Text = Format(NVL(!���췧ֵ, 0), "#0.00;-#0.00; ;")
            Case 2 'סԺ
                Me.txt���췧ֵ.Text = Format(NVL(!���췧ֵ, 0) / NVL(!סԺ��װ, 1), "#0.00;-#0.00; ;")
            Case 3 '����
                Me.txt���췧ֵ.Text = Format(NVL(!���췧ֵ, 0) / NVL(!�����װ, 1), "#0.00;-#0.00; ;")
            Case 4 'ҩ��
                Me.txt���췧ֵ.Text = Format(NVL(!���췧ֵ, 0) / NVL(!ҩ���װ, 1), "#0.00;-#0.00; ;")
            End Select
            
            Me.cboҩ������.ListIndex = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
            cboҩ������.Tag = Me.cboҩ������.ListIndex
            Me.txt����.Text = IIf(IsNull(!����), 100, !����)
            
            If Me.stbSpec.Tag = "����" Then
                Me.txtָ������.Text = ""
                Me.txtָ���ۼ�.Text = ""
                Me.txt�ɱ��۸�.Text = ""
                txt��ǰ�ۼ�.Text = ""
            Else
                If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
                    Me.txtָ������.Text = FormatEx(IIf(IsNull(!ָ��������), 0, !ָ��������) * Me.txtҩ���װ.Text, mintCostDigit, , True)
                    Me.txtָ���ۼ�.Text = FormatEx(IIf(IsNull(!ָ�����ۼ�), 0, !ָ�����ۼ�) * Me.txtҩ���װ.Text, mintPriceDigit, , True)
                    Me.txt�ɱ��۸�.Text = FormatEx(IIf(IsNull(!�ɱ���), 0, !�ɱ���) * Me.txtҩ���װ.Text, mintCostDigit, , True)
                Else
                    Me.txtָ������.Text = FormatEx(IIf(IsNull(!ָ��������), 0, !ָ��������), mintCostDigit, , True)
                    Me.txtָ���ۼ�.Text = FormatEx(IIf(IsNull(!ָ�����ۼ�), 0, !ָ�����ۼ�), mintPriceDigit, , True)
                    Me.txt�ɱ��۸�.Text = FormatEx(IIf(IsNull(!�ɱ���), 0, !�ɱ���), mintCostDigit, , True)
                End If
            End If
            Me.txt����� = FormatEx(Val(Me.txtָ������.Text) * Me.txt����.Text / 100, mintPriceDigit, , True)
                        
            Me.txt�ӳ���.Text = Format(IIf(IsNull(!�ӳ���), 0, !�ӳ���), "0.00")
            txt������Ŀ.Text = IIf(IsNull(!������Ŀ), "", !������Ŀ)
            
            For intCount = 0 To Me.cboҩ�ۼ���.ListCount - 1
                If Mid(Me.cboҩ�ۼ���.List(intCount), InStr(1, Me.cboҩ�ۼ���.List(intCount), "-") + 1) = IIf(IsNull(!ҩ�ۼ���), "", !ҩ�ۼ���) Then
                    Me.cboҩ�ۼ���.ListIndex = intCount: Exit For
                End If
            Next
            For intCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(intCount), InStr(1, Me.cbo��������.List(intCount), "-") + 1) = IIf(IsNull(!��������), "", !��������) Then
                    Me.cbo��������.ListIndex = intCount: Exit For
                End If
            Next
            Me.cbo�������.ListIndex = IIf(IsNull(!�������), 0, !�������)
            Me.chk���ηѱ�.Value = IIf(IsNull(!���ηѱ�), 0, !���ηѱ�)
            Me.chkסԺ��̬����.Value = IIf(IsNull(!סԺ��̬����), 0, !סԺ��̬����)
            Me.chk�ǳ���ҩ.Value = IIf(IsNull(!�Ƿ񳣱�), 0, !�Ƿ񳣱�)
            Me.chk��ҩ.Value = IIf(IsNull(!�Ƿ��ҩ), 0, !�Ƿ��ҩ)
            Me.chk����.Value = IIf(IsNull(!�Ƿ����۹���), 0, !�Ƿ����۹���)
            Me.chk�׵���.Value = IIf(IsNull(!�Ƿ���������), 0, !�Ƿ���������)
            Me.chk�����ɹ�.Value = IIf(IsNull(!�Ƿ�����ɹ�), 0, !�Ƿ�����ɹ�)
            
            If IsNull(!סԺ�ɷ����) Then
                Me.cboסԺ����.ListIndex = 0
            Else
                Select Case !סԺ�ɷ����
                Case Is >= 0
                    Me.cboסԺ����.ListIndex = !סԺ�ɷ����
                Case Else
                    Me.cboסԺ����.ListIndex = 2 + Abs(!סԺ�ɷ����)
                End Select
            End If
            
            If IsNull(!����ɷ����) Then
                Me.cbo�������.ListIndex = 0
            Else
                Select Case !����ɷ����
                Case Is >= 0
                    Me.cbo�������.ListIndex = !����ɷ����
                Case Else
                    Me.cbo�������.ListIndex = 2 + Abs(!����ɷ����)
                End Select
            End If
            
            Me.chkGMP��֤.Value = IIf(IsNull(!GMP��֤), 0, !GMP��֤)
'            Me.cboBasicDrug.MaxLength = .Fields("����ҩ��").DefinedSize
            Me.cboBasicDrug.Text = IIf(IsNull(!����ҩ��), "", !����ҩ��)
            
            If Me.stbSpec.Tag <> "����" Then mint�б�ҩƷ = IIf(IsNull(!�б�ҩƷ), 0, !�б�ҩƷ)
            If mint�б�ҩƷ = 1 Then Me.lblָ������.Caption = "�б�۸�"
            
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "ע���ù����" & Format(!����ʱ��, "YYYY��MM��DD��") & "����"
            Else
                Me.lblFound.Caption = "ע���ù����" & Format(!����ʱ��, "YYYY��MM��DD��") & "������" & Format(!����ʱ��, "YYYY��MM��DD��") & "ͣ��"
            End If
            Me.chkҩ��.Tag = IIf(IsNull(!ҩ������), 0, !ҩ������)
            Me.txtЧ��.Tag = IIf(IsNull(!���Ч��), 0, !���Ч��)
            
            Me.chkҩ��.Value = IIf(IsNull(!ҩ�����), 0, Abs(!ҩ�����))
            If Me.chkҩ��.Value = 0 Then
                Me.chkҩ��.Enabled = False: Me.chkҩ��.Value = 0
                Me.chkЧ��.Enabled = False: Me.chkЧ��.Value = 0
                Me.txtЧ��.Enabled = False: Me.chkЧ��.Value = 0
            Else
                Me.chkҩ��.Enabled = True
                Me.chkЧ��.Enabled = True
                Me.chkҩ��.Value = Me.chkҩ��.Tag
                Me.txtЧ��.Text = Me.txtЧ��.Tag
                If Val(Me.txtЧ��.Text) = 0 Then
                    Me.txtЧ��.Enabled = False: Me.chkЧ��.Value = 0
                Else
                    Me.txtЧ��.Enabled = True: Me.chkЧ��.Value = 1
                End If
            End If
            
            If NVL(!�洢�¶�) <> "" Then
                For i = 0 To Me.cboTemperature.ListCount - 1
                    If Me.cboTemperature.List(i) = NVL(!�洢�¶�) Then
                        Me.cboTemperature.Text = NVL(!�洢�¶�)
                        Exit For
                    End If
                Next
            Else
                Me.cboPrepareType.ListIndex = 0
            End If
            
            Me.chkCondition.Value = IIf(!�洢���� = 1, 1, 0)
            
            If Val(NVL(!��ҩ����)) <> 0 Then
                For i = 0 To Me.cboPrepareType.ListCount - 1
                    If Me.cboPrepareType.List(i) = NVL(!��ҩ����) Then
                        Me.cboPrepareType.Text = NVL(!��ҩ����)
                        Exit For
                    End If
                Next
            Else
                Me.cboPrepareType.ListIndex = 0
            End If
            
            Me.chkDosage.Value = IIf(!�Ƿ������� = 1, 1, 0)
        End If
        If Trim(Me.txt��ͬ��λ.Tag) = "" Then
            Me.txt��ͬ��λ.Tag = "|"
        End If
        If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
            Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txtҩ�ⵥλ
            Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txtҩ�ⵥλ
        Else
            Me.lbl���۵�λ(0).Caption = "Ԫ/" & Me.txt�ۼ۵�λ
            Me.lbl���۵�λ(1).Caption = "Ԫ/" & Me.txt�ۼ۵�λ
        End If
    End With
    
    If Me.stbSpec.Tag = "����" Then
        gstrSql = "Select �ӳ���" & vbNewLine & _
                "From ҩƷ���" & vbNewLine & _
                "Where ҩƷid = (Select Max(ҩƷid) From ҩƷ��� A, �շ���ĿĿ¼ B Where a.ҩƷid = b.Id And b.��� = [1])"

        Set rs����� = zlDatabase.OpenSQLRecord(gstrSql, "�ӳ��ʲ�ѯ", Me.Tag)
        If rs�����.RecordCount > 0 Then
            Me.txt�ӳ���.Text = Format(IIf(IsNull(rs�����!�ӳ���), 0, rs�����!�ӳ���), "0.00000")
        End If
        
        '����ʱ��������ȡ����ţ���չ��ͳ���
        Me.txt����.Text = "": Me.txt���.Text = "": Me.txt����.Text = "": Me.lblFound.Caption = "": chk�����ɹ�.Value = 0
        gstrSql = "select max(I.����) as ������ from �շ���ĿĿ¼ I,ҩƷ��� S where I.ID=S.ҩƷID and  S.ҩ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
        With rsTemp
            If .BOF Or .EOF Then
                Me.txt����.Text = Me.lblƷ��.Tag & "01"
            ElseIf IsNull(!������) Then
                Me.txt����.Text = Me.lblƷ��.Tag & "01"
            Else
                Me.txt����.Text = zlStr.Increase(!������)
            End If
        End With
        If txtDDDֵ.Visible = True Then
            gstrSql = "Select nvl(a.Dddֵ,0) dddֵ" & _
                      "  From ҩƷ��� A, �շ���ĿĿ¼ B, (Select Max(����ʱ��) ����ʱ�� From �շ���ĿĿ¼) C" & _
                       " Where a.ҩƷid = b.ID And b.����ʱ�� = c.����ʱ�� And a.ҩ��id = [1]" & _
                       " Union All" & _
                       " Select nvl(Dddֵ,0) From �����÷����� Where ��Ŀid = [1] and ����<>0"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "DDDֵ", lngҩ��id)
            Do While Not rsTemp.EOF
                If rsTemp!DDDֵ <> 0 Then
                    txtDDDֵ.Text = rsTemp!DDDֵ
                    Exit Do
                End If
                rsTemp.MoveNext
            Loop
        End If
        
        If mintSet���� = 0 Then
            gstrSql = "Select b.ҩ�����, b.ҩ������" & _
                       " From ҩƷ��� B, (Select Max(a.Id) As ID From �շ���ĿĿ¼ A, ҩƷ��� B Where a.Id = b.ҩƷid And b.ҩ��id = [1]) C" & _
                       " Where b.ҩƷid = c.Id"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            If rsTemp.RecordCount <> 0 Then
                chkҩ��.Value = IIf(IsNull(rsTemp!ҩ�����), "0", rsTemp!ҩ�����)
                chkҩ��.Value = IIf(IsNull(rsTemp!ҩ������), "0", rsTemp!ҩ������)
            End If
        ElseIf mintSet���� = 1 Then
            chkҩ��.Value = 1
            chkҩ��.Value = 0
            chkҩ��.Enabled = False
            chkҩ��.Enabled = False
        ElseIf mintSet���� = 2 Then
            chkҩ��.Value = 1
            chkҩ��.Value = 1
            chkҩ��.Enabled = False
            chkҩ��.Enabled = False
        ElseIf mintSet���� = 3 Then
            chkҩ��.Value = 0
            chkҩ��.Value = 0
            chkҩ��.Enabled = False
            chkҩ��.Enabled = False
        End If
    Else
        '��ȡ��Ʒ���ͼ��롢������
        gstrSql = "select ����,����,����,���� from �շ���Ŀ���� where �շ�ϸĿid=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        With rsTemp
            Do While Not .EOF
                If !���� = 1 And !���� = 3 Then
                    Me.txt������.Text = IIf(IsNull(!����), "", !����)
                End If
                If !���� = 3 And !���� = 1 Then
                    Me.txt��Ʒ��.Text = IIf(IsNull(!����), "", !����)
                    Me.txtƴ��.Text = IIf(IsNull(!����), "", !����)
                End If
                If !���� = 3 And !���� = 2 Then
                    Me.txt��Ʒ��.Text = IIf(IsNull(!����), "", !����)
                    Me.txt���.Text = IIf(IsNull(!����), "", !����)
                End If
                .MoveNext
            Loop
        End With
        
        '��ȡ��ʾ��ǰ�ۼ�
        If Me.cboҩ������.ListIndex <> 0 Then
            'ʱ��ҩƷ��ȡ�����/���������Ϊ��۸��޿��ʱȡ�۱���
            gstrSql = "select Decode(K.�������,0,P.�ּ�,K.�����/Nvl(K.�������,1)) as �ּ�,P.������Ŀid" & _
                    " from �շѼ�Ŀ P," & _
                    "     (Select nvl(Sum(ʵ�ʽ��),0) as �����,nvl(Sum(ʵ������),0) as �������" & _
                    "      From ҩƷ��� Where ҩƷID=[1]) K" & _
                    " where P.�շ�ϸĿid=[1] and (P.��ֹ���� is null or Sysdate Between P.ִ������ And P.��ֹ����)" & _
                    GetPriceClassString("P")
        Else
            '��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
            gstrSql = "select P.�ּ�,P.������Ŀid" & _
                    " from �շѼ�Ŀ P" & _
                    " where P.�շ�ϸĿid=[1] and (P.��ֹ���� is null or Sysdate Between P.ִ������ And P.��ֹ����)" & _
                    GetPriceClassString("P")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        With rsTemp
            If .RecordCount > 0 Then
                If Val(Me.lbl���۵�λ(0).Tag) <> 0 Then
                    Me.txt��ǰ�ۼ�.Text = FormatEx(!�ּ� * Val(txtҩ���װ.Text), mintPriceDigit, , True)
                Else
                    Me.txt��ǰ�ۼ�.Text = FormatEx(!�ּ�, mintPriceDigit, , True)
                End If
                For intCount = 0 To Me.cbo�������.ListCount - 1
                    If Me.cbo�������.ItemData(intCount) = !������Ŀid Then
                        Me.cbo�������.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        End With
        
        '�����Ƿ��з�����ȷ����ҩ�����ԡ��ɱ��۸����ۼ۸���޸ķ�
        gstrSql = " Select nvl(Count(*),0) From ҩƷ�շ���¼ Where ҩƷID=[1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        mblnUsed = False
        With rsTemp
            If .Fields(0).Value > 0 Then
                mblnUsed = True
                Me.txt�ɱ��۸�.Enabled = False
                Me.txt��ǰ�ۼ�.Enabled = False
'                Me.cbo�������.Enabled = False
'                Me.txt����ϵ��.Enabled = False
                Me.txtסԺ��װ.Enabled = False
                Me.txt�����װ.Enabled = False
                Me.txtҩ���װ.Enabled = False
            Else
                Me.txt��ǰ�ۼ�.Enabled = True
                Me.txt�ɱ��۸�.Enabled = True
'                Me.cbo�������.Enabled = True
'                Me.txt����ϵ��.Enabled = True
                Me.txtסԺ��װ.Enabled = True
                Me.txt�����װ.Enabled = True
                Me.txtҩ���װ.Enabled = True
            End If
        End With
        
        '�����Ƿ����ҽ����¼��ȷ������ϵ���Ƿ��ܹ��޸�
        gstrSql = "Select 1 From ����ҽ����¼ Where �շ�ϸĿID=[1] And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        If rsTemp.RecordCount > 0 Then
            Me.txt����ϵ��.Enabled = False
        Else
            Me.txt����ϵ��.Enabled = True
        End If
        
        '�����Ƿ��п�棬ȷ�����������Կ��޸ķ�
        gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                 " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And B.�������� Like '%ҩ��'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        If rsTemp.Fields(0).Value > 0 Then
            Me.chkҩ��.Enabled = False
            Me.chkЧ��.Enabled = False
        Else
            Me.chkҩ��.Enabled = True
        End If
        If Me.chkҩ��.Value = 1 Then
            gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                     " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '%ҩ��' Or B.�������� Like '%�Ƽ���')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
            
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
        If InStr(1, strPrivs, "ҽ����ҩĿ¼") = 0 Then
            Me.cbo��������.Enabled = False: Me.txt��ʶ��.Enabled = False:
        End If
        If InStr(1, strPrivs, "�������") = 0 Then Me.txt����.Enabled = False
        If InStr(1, strPrivs, "ָ���۸����") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txtָ������.Text = ""
                Me.txtָ���ۼ�.Text = ""
            End If
            Me.txt�ӳ���.Enabled = False
            Me.txtָ������.Enabled = False: Me.txtָ���ۼ�.Enabled = False
        End If
        If InStr(1, strPrivs, "�ۼ۹���") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txt��ǰ�ۼ�.Text = ""
                Me.cboҩ������.ListIndex = 0
            End If
            Me.cboҩ������.Enabled = False
        End If
        If InStr(1, strPrivs, "����������Ŀ") = 0 Then
            cbo�������.Enabled = False
        End If
        If InStr(1, strPrivs, "ҩ�ۼ���") = 0 Then
             Me.cboҩ�ۼ���.Enabled = False
        End If
        If InStr(1, strPrivs, "�ɱ��۹���") = 0 Then
            If Me.stbSpec.Tag = "����" Then
                Me.txt�ɱ��۸�.Text = ""
            End If
            Me.txt�ɱ��۸�.Enabled = False
        End If
        If InStr(1, strPrivs, "�����������") = 0 Then
            Me.cbo�������.Enabled = False
        End If
    Else
        Me.txt����.Enabled = False: Me.txt��λ��.Enabled = False: Me.txt���.Enabled = False: Me.txt����.Enabled = False: cmd����.Enabled = False
        Me.txt��Ʒ��.Enabled = False: Me.txtƴ��.Enabled = False: Me.txt���.Enabled = False: Me.txt������.Enabled = False
        Me.txt��ʶ��.Enabled = False: Me.cboҩƷ��Դ.Enabled = False: Me.txt��׼�ĺ�.Enabled = False: Me.txtע���̱�.Enabled = False
        Me.txt�ۼ۵�λ.Enabled = False: Me.txt����ϵ��.Enabled = False: Me.txt���ﵥλ.Enabled = False: Me.txt�����װ.Enabled = False
        Me.txtסԺ��λ.Enabled = False: Me.txtסԺ��װ.Enabled = False: Me.txtҩ�ⵥλ.Enabled = False: Me.txtҩ���װ.Enabled = False
        Me.cbo���쵥λ.Enabled = False: Me.txt���췧ֵ.Enabled = False: Me.cbo��ҩ����.Enabled = False: Me.txt����.Enabled = False: Me.cbo��ΣҩƷ.Enabled = False
        
        Me.cboҩ������.Enabled = False: Me.txtָ������.Enabled = False: Me.txt����.Enabled = False: Me.txt�����.Enabled = False
        Me.txtָ���ۼ�.Enabled = False: Me.txt�ӳ���.Enabled = False
        Me.cboҩ�ۼ���.Enabled = False: Me.cbo��������.Enabled = False: Me.cbo�������.Enabled = False: Me.chk���ηѱ�.Enabled = False
        Me.txt�ɱ��۸�.Enabled = False: Me.txt��ǰ�ۼ�.Enabled = False: Me.cbo�������.Enabled = False
        Me.cboסԺ����.Enabled = False: Me.chkҩ��.Enabled = False: Me.chkҩ��.Enabled = False: Me.chkЧ��.Enabled = False: Me.txtЧ��.Enabled = False
        Me.cbo�������.Enabled = False
        Me.chkסԺ��̬����.Enabled = False
        Me.txt��ͬ��λ.Enabled = False: Me.cmd��ͬ��λ.Enabled = False
        Me.txt˵��.Enabled = False
        Me.cboBasicDrug.Enabled = False
        Me.txt��ѡ��.Enabled = False
        Me.cmbStationNo.Enabled = False
        Me.chk�ǳ���ҩ.Enabled = False
        Me.cboTemperature.Enabled = False
        Me.chkCondition.Enabled = False
        Me.cboPrepareType.Enabled = False
        Me.chkDosage.Enabled = False
        txt������Ŀ.Enabled = False
        cmd����.Enabled = False
        Me.txt����.Enabled = False
        txtDDDֵ.Visible = False
        lblddd.Visible = False
        lbldddֵ.Visible = False
        cmdOK.Visible = False: cmdCancel.Caption = "�ر�(&C)"
        chk��ҩ.Enabled = False
        chk�׵���.Enabled = False
        chk����.Enabled = False
        chkGMP��֤.Enabled = False
        vsfItem.Enabled = False
        chk�����ɹ�.Enabled = False
    End If
    
    '������β������޸ģ������Ƿ���ڡ�ҩƷ��λ������Ȩ�ޣ�û���������޸�ҩƷ��λ��ϵ��
    If Me.stbSpec.Tag = "�޸�" Then
        If InStr(1, strPrivs, "ҩƷ��λ����") = 0 Then
            txt�ۼ۵�λ.Enabled = False
            txtסԺ��λ.Enabled = False
            txt���ﵥλ.Enabled = False
            txtҩ�ⵥλ.Enabled = False
            txt����ϵ��.Enabled = False
            txtסԺ��װ.Enabled = False
            txt�����װ.Enabled = False
            txtҩ���װ.Enabled = False
        End If
    End If
'    mblnLoad = True
    Me.stbSpec.Tab = 0
    mstr���м�¼ = ""
    mstr���м�¼ = txt����.Text & "|" & txt��λ�� & "|" & txt���.Text & "|" & txt����.Text & "|" & txt��Ʒ��.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & _
                    txt������.Text & "|" & txt��ʶ��.Text & "|" & cboҩƷ��Դ.Text & "|" & txt��ͬ��λ.Text & "|" & txt˵��.Text & "|" & cbo��ҩ����.Text & "|" & _
                    cmbStationNo.Text & "|" & txt��׼�ĺ�.Text & "|" & txtע���̱�.Text & "|" & txt�ۼ۵�λ.Text & "|" & txt����ϵ��.Text & "|" & txtסԺ��λ.Text & "|" & _
                    txtסԺ��װ.Text & "|" & txt���ﵥλ.Text & "|" & txt�����װ.Text & "|" & txtҩ�ⵥλ.Text & "|" & txtҩ���װ.Text & "|" & cbo���쵥λ.Text & "|" & txt���췧ֵ.Text & "|" & _
                    txt��ѡ��.Text & "|" & txt����.Text & "|" & cboҩ������.Text & "|" & txt�ɱ��۸�.Text & "|" & txt��ǰ�ۼ�.Text & "|" & txtָ������.Text & "|" & txt����.Text & "|" & txt�����.Text & "|" & _
                    txtָ���ۼ�.Text & "|" & txt�ӳ���.Text & "|" & cbo�������.Text & "|" & txt������Ŀ.Text & "|" & cboҩ�ۼ���.Text & "|" & _
                    chk���ηѱ�.Value & "|" & cbo��������.Text & "|" & cbo�������.Text & "|" & cboסԺ����.Text & "|" & cboBasicDrug.Text & "|" & chkסԺ��̬����.Value & "|" & _
                    chkGMP��֤.Value & "|" & chk�ǳ���ҩ.Value & "|" & chkҩ��.Value & "|" & chkҩ��.Value & "|" & chkЧ��.Value & "|" & txtЧ��.Text & "|" & cboTemperature.Text & "|" & chkCondition.Value & "|" & _
                    cboPrepareType.Text & "|" & chkDosage.Value & "|" & cbo�������.Text & "|" & txtDDDֵ.Text & "|" & cbo��ΣҩƷ.Text & "|" & chk�׵���.Value & "|" & chk�����ɹ�.Value
    If txt���.Enabled = True Then
        txt���.SetFocus
    End If
    
    '��չ����
    gstrSql = "Select 1 From ҩƷ�����չ��Ŀ Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷ�����չ��Ŀ")
    If rsTmp.RecordCount = 0 Then
        '���û����չ��Ŀ�Ͳ���ʾ��չҳ��
        stbSpec.TabVisible(3) = False
    Else
        gstrSql = "Select b.����, a.���� From ҩƷ�����չ��Ϣ A, ҩƷ�����չ��Ŀ B " & _
            " Where a.��Ŀ(+) = b.���� And a.ҩƷid(+) = [1] Order By b.���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷ�����չ��Ϣ", lngҩƷID)
        
        mlngExpItemMaxLength = rsTmp.Fields("����").DefinedSize
        
        With vsfItem
            .Rows = 1
            
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = rsTmp!����
                .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTmp!����)
                .TextMatrix(.Rows - 1, .ColIndex("ԭ����")) = NVL(rsTmp!����)
                
                rsTmp.MoveNext
            Loop
        End With
    End If
    
    '����ģʽ����
    If Val(zlDatabase.GetPara(275, glngSys, , 0)) = 0 Then
        chk����.Enabled = False
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    
    mint�б�ҩƷ = 0
    On Error GoTo errHandle
    
    Call GetMaxDigit
        
    '�����ҩ�����ϵͳ������ʾ���ＰסԺ��ص�λ��ϵ�������������ۼ۵�λ��ϵ��һ��
    If glngSys \ 100 = 8 Then
        Me.lbl���ﵥλ.Visible = False: Me.txt���ﵥλ.Visible = False: Me.lbl�����װ.Visible = False: Me.txt�����װ.Visible = False: Me.lbl���ﵥλChild.Visible = False
        Me.lblסԺ��λ.Visible = False: Me.txtסԺ��λ.Visible = False: Me.lblסԺ��װ.Visible = False: Me.txtסԺ��װ.Visible = False: Me.lblҩ�ⵥλChild.Visible = False
        Me.lblҩ���װ.Top = Me.lblסԺ��װ.Top: Me.txtҩ�ⵥλ.Top = Me.txtסԺ��λ.Top: Me.lblҩ�ⵥλ.Top = Me.lblסԺ��λ.Top: Me.txtҩ���װ.Top = Me.txtסԺ��װ.Top
        Me.lblҩ�ⵥλ.Caption = "�ɹ���λ(&W)"
    End If
    
    Call GetDefineSize
    Call IniStationNo
    
    mint�ֶμӳ� = Val(zlDatabase.GetPara("�ۼ۰��ӳɼ���", glngSys, 1023, 0))
    
    Set mrs�ֶμӳ� = Nothing
    If mint�ֶμӳ� = 1 Then
        gstrSql = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵�� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷ�ӳɷ���")
    End If
    '----------------װ���ѡ�Ļ�������----------------------
    With Me.cboҩ������
        .Clear
        aryTemp = Split("0-����;1-ʱ��", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(intCount): .ItemData(.NewIndex) = intCount
        Next
        .ListIndex = 0
    End With
    
    gstrSql = "Select ����||'-'||���� ���� From ҩ�۹�����  Order By ����"
    With rsTemp
'        If .State = adStateOpen Then .Close
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        Me.cboҩ�ۼ���.Clear
        Do While Not rsTemp.EOF
            Me.cboҩ�ۼ���.AddItem rsTemp!����
            rsTemp.MoveNext
        Loop
    End With
    
    With Me.cboסԺ����
        .Clear
        .AddItem "0-���Է���": .ItemData(.NewIndex) = 0
        .AddItem "1-���ɷ���": .ItemData(.NewIndex) = 1
        .AddItem "2-һ����ʹ��": .ItemData(.NewIndex) = 2
        .AddItem "3-�����һ������Ч": .ItemData(.NewIndex) = -1
        .AddItem "4-�������������Ч": .ItemData(.NewIndex) = -2
        .AddItem "5-�������������Ч": .ItemData(.NewIndex) = -3
        .ListIndex = 0
    End With
    
    With Me.cbo�������
        .Clear
        .AddItem "0-���Է���": .ItemData(.NewIndex) = 0
        .AddItem "1-���ɷ���": .ItemData(.NewIndex) = 1
        .AddItem "2-һ����ʹ��": .ItemData(.NewIndex) = 2
        .AddItem "3-�����һ������Ч": .ItemData(.NewIndex) = -1
        .AddItem "4-�������������Ч": .ItemData(.NewIndex) = -2
        .AddItem "5-�������������Ч": .ItemData(.NewIndex) = -3
        .ListIndex = 0
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
    
    gstrSql = "Select ����  From ����ҩ��˵��  Order By ����"
    With cboBasicDrug
        Dim rsRecord As ADODB.Recordset
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "����ҩ��˵��")
            .AddItem ""
        Do While Not rsRecord.EOF
            .AddItem rsRecord!����
            rsRecord.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    With cbo���쵥λ
        .Clear
        .AddItem "�ۼ۵�λ"
        .AddItem "סԺ��λ"
        .AddItem "���ﵥλ"
        .AddItem "ҩ�ⵥλ"
        .ListIndex = 0
    End With
    
    With rsTemp
        gstrSql = "Select ����||'-'||���� From ҩƷ��Դ���� Order By ����"
'        If .State = adStateOpen Then .Close
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        Me.cboҩƷ��Դ.Clear
        Do While Not rsTemp.EOF
            Me.cboҩƷ��Դ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cboҩƷ��Դ.ListCount > 0 Then Me.cboҩƷ��Դ.ListIndex = 0
        
        gstrSql = "Select ���� From ��ҩ���� Order By ����"
'        If .State = adStateOpen Then .Close
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        Me.cbo��ҩ����.Clear
        Me.cbo��ҩ����.AddItem ""
        Do While Not rsTemp.EOF
            Me.cbo��ҩ����.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
    
        gstrSql = "Select ����||'-'||���� From �������� where ����=1 Order By ����"
'        If .State = adStateOpen Then .Close
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        Me.cbo��������.Clear
        Me.cbo��������.AddItem ""
        Do While Not rsTemp.EOF
            Me.cbo��������.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        
        gstrSql = "Select ID,���� as ����" & _
                " From ������Ŀ" & _
                " where ĩ��=1 and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By ����"
'        If .State = adStateOpen Then .Close
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        Me.cbo�������.Clear
        Do While Not rsTemp.EOF
            Me.cbo�������.AddItem rsTemp!����: Me.cbo�������.ItemData(Me.cbo�������.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo�������.ListCount > 0 Then Me.cbo�������.ListIndex = 0
    End With
    
    With cbo��ΣҩƷ
        .AddItem ""
        .AddItem "1-A��"
        .AddItem "2-B��"
        .AddItem "3-C��"
        .ListIndex = 0
    End With
    
'    '��Һ����������Ҫ��ҩƷ��ҩ��������
'    stbSpec.TabVisible(2) = False
'    gstrSql = "Select Nvl(����ֵ, 0) From zlParameters Where ϵͳ = 100 And Nvl(˽��, 0) = 0 And ģ�� Is Null And ������ = 153"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "������������")
'    If Not rsTmp.EOF Then
'        If rsTmp.Fields(0).Value > 1 Then
'            stbSpec.TabVisible(2) = True
'        End If
'    End If

    gstrSql = "select ����,���� from ҩƷ�洢�¶� order by ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷ�洢�¶�")
    With cboTemperature
        .Clear
        .AddItem ""
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    gstrSql = "select ����,���� from ��Һ��ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "��ҩ��������")
    With cboPrepareType
        .Clear
        .AddItem ""
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
   
    zlControl.CboSetWidth cbo�������.hwnd, 1500
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    
    If mblnOtherSave = False Then
        If mblnOK = False And mblnCancel = False Then
            strTemp = txt����.Text & "|" & txt��λ�� & "|" & txt���.Text & "|" & txt����.Text & "|" & txt��Ʒ��.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & _
                            txt������.Text & "|" & txt��ʶ��.Text & "|" & cboҩƷ��Դ.Text & "|" & txt��ͬ��λ.Text & "|" & txt˵��.Text & "|" & cbo��ҩ����.Text & "|" & _
                            cmbStationNo.Text & "|" & txt��׼�ĺ�.Text & "|" & txtע���̱�.Text & "|" & txt�ۼ۵�λ.Text & "|" & txt����ϵ��.Text & "|" & txtסԺ��λ.Text & "|" & _
                            txtסԺ��װ.Text & "|" & txt���ﵥλ.Text & "|" & txt�����װ.Text & "|" & txtҩ�ⵥλ.Text & "|" & txtҩ���װ.Text & "|" & cbo���쵥λ.Text & "|" & txt���췧ֵ.Text & "|" & _
                            txt��ѡ��.Text & "|" & txt����.Text & "|" & cboҩ������.Text & "|" & txt�ɱ��۸�.Text & "|" & txt��ǰ�ۼ�.Text & "|" & txtָ������.Text & "|" & txt����.Text & "|" & txt�����.Text & "|" & _
                            txtָ���ۼ�.Text & "|" & txt�ӳ���.Text & "|" & cbo�������.Text & "|" & txt������Ŀ.Text & "|" & cboҩ�ۼ���.Text & "|" & _
                            chk���ηѱ�.Value & "|" & cbo��������.Text & "|" & cbo�������.Text & "|" & cboסԺ����.Text & "|" & cboBasicDrug.Text & "|" & chkסԺ��̬����.Value & "|" & _
                            chkGMP��֤.Value & "|" & chk�ǳ���ҩ.Value & "|" & chkҩ��.Value & "|" & chkҩ��.Value & "|" & chkЧ��.Value & "|" & txtЧ��.Text & "|" & cboTemperature.Text & "|" & chkCondition.Value & "|" & _
                            cboPrepareType.Text & "|" & chkDosage.Value & "|" & cbo�������.Text & "|" & txtDDDֵ.Text & "|" & cbo��ΣҩƷ.Text & "|" & chk�׵���.Value & "|" & chk�����ɹ�.Value
            If strTemp <> mstr���м�¼ Then
                If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
    '                mblnLoad = False
                    mblnOK = False
                    mblnCancel = False
                    mbln������Ŀ = False
                End If
            Else
    '            mblnLoad = False
                mblnOK = False
                mblnCancel = False
                mbln������Ŀ = False
            End If
        End If
    End If
'    mblnLoad = False
    mblnOK = False
    mblnCancel = False
    mblnOtherSave = False
    mbln������Ŀ = False
End Sub

Private Sub txtDDDֵ_GotFocus()
    zlControl.TxtSelAll txtDDDֵ
End Sub

Private Sub txtDDDֵ_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim Count As Integer
    
    If KeyAscii = vbKeyReturn Then
        stbSpec.Tab = 1
        If cboҩ������.Enabled = True Then
            cboҩ������.SetFocus
        End If
        Exit Sub
    End If
    strText = Me.txtDDDֵ.Text
    If Val(strText) > 100000000 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13 Then
        If strText <> "" Then
            If KeyAscii = 46 Then
                Count = (Len(strText) - Len(Replace(strText, ".", ""))) / Len(".")
                
                If Count > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Else
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
    strText = ""
    
    If KeyAscii = vbKeyReturn Then
        Me.stbSpec.Tab = 1
        If Me.cboҩ������.Enabled Then
            Me.cboҩ������.SetFocus
        Else
            Me.txtָ������.SetFocus
        End If
    End If
End Sub

Private Sub txt������Ŀ_GotFocus()
    txt������Ŀ.SelStart = 0
    txt������Ŀ.SelLength = Len(txt������Ŀ)
    If Me.stbSpec.Tag = "����" Or Me.stbSpec.Tag = "�޸�" Then
        txt������Ŀ.SetFocus
    End If
End Sub

Private Sub txt������Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
        Exit Sub
    End If
    If KeyAscii = vbKeyDelete Then
        txt������Ŀ.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cboBasicDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��ѡ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
        Exit Sub
    End If
    
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
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Asc("-")
        If InStr(1, txt����.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
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
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*_+|=-`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii > 255 Or KeyAscii < 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim reTmp As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt����.hwnd)

    If InStr("~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(txt����))
    If strTemp = "" Then Me.txt����.Tag = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    On Error GoTo errHandle
    gstrSql = "Select ���� as id,����,����" & _
            " From ҩƷ������" & _
            " where ���� Like [1] " & _
            "       Or ���� Like [1] " & _
            "       Or ���� Like [1] Order By ���� "
'    Set reTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    Set reTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, Me.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrMatch & strTemp & "%")

    If blnCancel = True Then txt����.SetFocus: Exit Sub  '��ѡ����ʱ����Esc�������´���
    
    With reTmp
        If reTmp Is Nothing Then
            If Me.txt����.Tag <> strTemp Then
                If Asc(strTemp) > 0 Then
                    MsgBox "û���ҵ�ƥ��������̣����������룡", vbInformation, gstrSysName
                    Me.txt����.SelStart = 0: Me.txt����.SelLength = LenB(StrConv(txt����, vbFromUnicode)): Me.txt����.Tag = "":
                    Exit Sub
                End If
                If MsgBox("û���ҵ���ص������̣����Ӹ���������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Me.txt����.SelStart = 0: Me.txt����.SelLength = LenB(StrConv(txt����, vbFromUnicode)): Me.txt����.Tag = "": Me.txt����.Text = "": Exit Sub
                Else
                    If zlSureManufacturer = False Then
                        MsgBox "�����̵ı��볬�����޷��Զ����ӡ�" & vbCrLf & "�������ѡ�����е�ҩƷ�����̣�", vbInformation, gstrSysName
                        Me.txt����.Text = "": Me.txt����.Tag = "": Exit Sub
                    Else
                        Me.txt����.Tag = Me.txt����: Call OS.PressKey(vbKeyTab): Exit Sub
                    End If
                End If
            End If
            Exit Sub
        End If
        
        txt����.SetFocus
        txt���� = !����
        txt����.Tag = txt����
            
    End With
    
    Call OS.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt�ɱ��۸�_GotFocus()
    Me.txt�ɱ��۸�.SelStart = 0: Me.txt�ɱ��۸�.SelLength = 100
End Sub

Private Sub txt�ɱ��۸�_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt�ɱ��۸�.SelLength = Len(txt�ɱ��۸�.Text) Then Exit Sub
            If Len(Mid(txt�ɱ��۸�, InStr(1, txt�ɱ��۸�.Text, ".") + 1)) >= mintCostDigit And txt�ɱ��۸�.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ɱ��۸�_LostFocus()
    Dim dblSalePrice As Double
    Dim dbl�۸� As Double
    
    Me.txt�ɱ��۸�.Text = FormatEx(Val(Me.txt�ɱ��۸�.Text), mintCostDigit, , True)
    txtָ������.Text = txt�ɱ��۸�.Text
    If Val(Me.txt��ǰ�ۼ�.Text) = 0 And Val(Me.txt�ɱ��۸�.Text) <> 0 Then
        If mint�ֶμӳ� = 0 Then    '����ͨ�ӳɷ�ʽ
            dblSalePrice = Val(Me.txt�ɱ��۸�.Text) * (1 + Val(Me.txt�ӳ���.Text) / 100)
        Else    '���ֶμӳɷ�ʽ
            dblSalePrice = get�ֶμӳ��ۼ�(Val(Me.txt�ɱ��۸�.Text))
        End If
                
        If Val(Me.txtָ���ۼ�.Text) > 0 Then
            If dblSalePrice > Val(Me.txtָ���ۼ�.Text) Then dblSalePrice = Val(Me.txtָ���ۼ�.Text)
        End If
        
        Me.txt��ǰ�ۼ�.Text = FormatEx(dblSalePrice, mintPriceDigit, , True)
        
        If mint�ֶμӳ� = 1 Then
            dbl�۸� = mdbl�ӳ��� * 100
            Me.txt�ӳ���.Text = Format(mdbl�ӳ��� * 100, "0.00")
        End If
    End If
End Sub

Private Function get�ֶμӳ��ۼ�(ByVal dbl�ɹ��� As Double) As Double
    Dim blnData As Boolean
    
    mdbl�ӳ��� = 0
    mdbl��۶� = 0
    
    Do Until mrs�ֶμӳ�.EOF
        If dbl�ɹ��� > mrs�ֶμӳ�!��ͼ� And dbl�ɹ��� <= mrs�ֶμӳ�!��߼� Then
            mdbl�ӳ��� = mrs�ֶμӳ�!�ӳ��� / 100
            mdbl��۶� = IIf(IsNull(mrs�ֶμӳ�!��۶�), 0, mrs�ֶμӳ�!��۶�)
            blnData = True
            Exit Do
        End If
        mrs�ֶμӳ�.MoveNext
    Loop
    If blnData = False Then
        MsgBox "û�����ý���Ϊ��" & dbl�ɹ��� & "  �ķֶμӳ����ݣ�����ҩƷĿ¼�����ֶμӳ��ʣ�������"
        get�ֶμӳ��ۼ� = 0
        Exit Function
    Else
        get�ֶμӳ��ۼ� = dbl�ɹ��� * (1 + mdbl�ӳ���) + mdbl��۶�
    End If
End Function

Private Sub txt��ǰ�ۼ�_GotFocus()
    Me.txt��ǰ�ۼ�.SelStart = 0: Me.txt��ǰ�ۼ�.SelLength = 100
End Sub

Private Sub txt��ǰ�ۼ�_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt��ǰ�ۼ�.SelLength = Len(txt��ǰ�ۼ�.Text) Then Exit Sub
            If Len(Mid(txt��ǰ�ۼ�, InStr(1, txt��ǰ�ۼ�.Text, ".") + 1)) >= mintPriceDigit And txt��ǰ�ۼ�.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ǰ�ۼ�_LostFocus()
    Dim dbl�ɱ��� As Double
    Dim dblָ���ۼ� As Double
    Dim dbl�ӳ��� As Double
    Dim dbl����� As Double
    Dim dbl���ۼ� As Double
    
    Me.txt��ǰ�ۼ�.Text = FormatEx(Val(txt��ǰ�ۼ�), mintPriceDigit, , True)
    txtָ���ۼ�.Text = txt��ǰ�ۼ�.Text
    
    dbl���ۼ� = Val(Me.txt��ǰ�ۼ�.Text)
    dbl�ɱ��� = Val(Me.txt�ɱ��۸�.Text)
    dblָ���ۼ� = Val(Me.txtָ���ۼ�.Text)
    
    '������Щ�����ż���ӳ���
    If dbl�ɱ��� > 0 And dblָ���ۼ� > 0 And dbl���ۼ� > 0 And dbl���ۼ� <= dblָ���ۼ� Then
        If mint�ֶμӳ� = 0 Then
            dbl�ӳ��� = dbl���ۼ� / dbl�ɱ��� - 1
            
            If dbl�ӳ��� < 0 Then Exit Sub
            
            dbl�ӳ��� = dbl�ӳ��� * 100
        Else
            dbl�ӳ��� = mdbl�ӳ��� * 100
        End If
        
        Me.txt�ӳ���.Text = Format(dbl�ӳ���, "0.00")
    End If
    
'    If Trim(txt��ǰ�ۼ�.Text) <> "" And Val(Trim(txtָ���ۼ�.Text)) = 0 Then
'        txtָ���ۼ�.Text = txt��ǰ�ۼ�.Text
'    End If
'��ʱ���ݳɱ��ۡ��ӳ��ʡ����������ָ���ۼ��������ۼ۵Ĺ�ʽ
'    Me.txt�ɱ��۸�.Text = FormatEx(Val(Me.txt�ɱ��۸�.Text), mintCostDigit)
'    If Val(Me.txt��ǰ�ۼ�.Text) = 0 And Val(Me.txt�ɱ��۸�.Text) <> 0 Then
'        dblSalePrice = Val(Me.txt�ɱ��۸�.Text) * (1 + Val(Me.txt�ӳ���.Text) / 100)
'        dblSalePrice = dblSalePrice + (Val(Me.txtָ���ۼ�.Text) - dblSalePrice) * (1 - Val(Me.txt�������) / 100)
'        If dblSalePrice > Val(Me.txtָ���ۼ�.Text) Then dblSalePrice = Val(Me.txtָ���ۼ�.Text)
'        Me.txt��ǰ�ۼ�.Text = FormatEx(dblSalePrice, mintPriceDigit)
'    End If

'��������Ĺ�ʽ�õ��ӳ��ʻ�����ʽ
'    If �����ۼ� <= ָ���ۼ� And ������� <> 0 Then
'        If ������� = 1 Then
'           �ӳ��� = ���ۼ� / �ɱ��� - 1
'        Else
'           �ӳ��� = ((���ۼ� - ָ���ۼ� * (1 - �������)) / �������) / �ɱ��� - 1
'        End If
'    End If
 
'��1
'    �ɱ��� = 1
'    ָ���ۼ� = 3
'    �ӳ��� = 0.15
'
'    ������� = 0.6
'
'
'    �ӳ��ۼ� = �ɱ��� * (1 + �ӳ���) = 1 * (1 + 0.15) = 1.15
'    ���ۼ� = �ӳ��ۼ� + (ָ���ۼ� - �ӳ��ۼ�) * (1 - �������) = 1.15 + (3 - 1.15) * (1 - 0.6) = 1.89

'��2
'    �ɱ��� = 1
'    ָ���ۼ� = 3
'    �ӳ��� = 0.20
'
'    ������� = 0.6
'
'
'    �ӳ��ۼ� = �ɱ��� * (1 + �ӳ���) = 1 * (1 + 0.2) = 1.2
'    ���ۼ� = �ӳ��ۼ� + (ָ���ۼ� - �ӳ��ۼ�) * (1 - �������) = 1.2 + (3 - 1.2) * (1 - 0.6) = 1.92

'��3���������=0��
'    �ɱ��� = 1
'    ָ���ۼ� = 3
'    �ӳ��� = 0.20
'
'    ������� = 0
'
'
'    �ӳ��ۼ� = �ɱ��� * (1 + �ӳ���) = 1 * (1 + 0.2) = 1.2
'    ���ۼ� = �ӳ��ۼ� + (ָ���ۼ� - �ӳ��ۼ�) * (1 - �������) = 1.2 + (3 - 1.2) * (1 - 0) = 3

'��4���������=100��
'    �ɱ��� = 1
'    ָ���ۼ� = 3
'    �ӳ��� = 0.20
'
'    ������� = 1
'
'
'    �ӳ��ۼ� = �ɱ��� * (1 + �ӳ���) = 1 * (1 + 0.2) = 1.2
'    ���ۼ� = �ӳ��ۼ� + (ָ���ۼ� - �ӳ��ۼ�) * (1 - �������) = 1.2 + (3 - 1.2) * (1 - 1) = 1.2
End Sub

Private Sub txt���_Change()
    Me.txt������.Text = zlGetDigitSign(lngҩ��id, Trim(Me.txt���.Text))
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��ͬ��λ_GotFocus()
    Me.txt��ͬ��λ.SelStart = 0: Me.txt��ͬ��λ.SelLength = Len(Me.txt��ͬ��λ.Text)
End Sub

Private Sub txt��ͬ��λ_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim reTmp As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt��ͬ��λ.hwnd)
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    
    strTmp = UCase(Trim(Me.txt��ͬ��λ.Text))
    
    If strTmp = "" Then
        Me.txt��ͬ��λ.Tag = "|"
        Call OS.PressKey(vbKeyTab): Exit Sub
    ElseIf strTmp = Split(Me.txt��ͬ��λ.Tag, "|")(1) Then
        Call OS.PressKey(vbKeyTab): Exit Sub
    End If
    
    gstrSql = "Select ����,����,����,id" & _
            " From ��Ӧ��" & _
            " where (���� Like [1] " & _
            "       Or ���� Like [1] " & _
            "       Or ���� Like [1])" & _
            " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
'    Set reTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    Set reTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, Me.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrMatch & strTmp & "%")

    If blnCancel = True Then txt��ͬ��λ.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���

    With reTmp
        If reTmp Is Nothing Then
            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
            Me.txt��ͬ��λ.Text = Split(Me.txt��ͬ��λ.Tag, "|")(1)
            Me.txt��ͬ��λ.SelStart = 0: Me.txt��ͬ��λ.SelLength = Len(Me.txt��ͬ��λ.Text)
            Exit Sub
        End If
        
        txt��ͬ��λ.SetFocus
        Me.txt��ͬ��λ.Text = reTmp!����
        Me.txt��ͬ��λ.Tag = reTmp!ID & "|" & reTmp!����

    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
        Me.txt�����װ = 1
        Me.txtסԺ��װ = 1
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
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If InStr(txt����ϵ��.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    'ֻ�ܴ���һ��С����
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt����ϵ��.SelLength = Len(txt����ϵ��.Text) Then Exit Sub
            If Len(Mid(txt����ϵ��, InStr(1, txt����ϵ��.Text, ".") + 1)) >= 5 And txt����ϵ��.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ӳ���_Change()
    If Val(txt�ӳ���.Text) > 9900 Then txt�ӳ���.Text = 9900
    If Val(txt�ӳ���.Text) < 0 Then txt�ӳ���.Text = 0
End Sub

Private Sub txt�ӳ���_GotFocus()
    Call zlControl.TxtSelAll(txt�ӳ���)
End Sub



Private Sub txt�ӳ���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If InStr(txt�ӳ���.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            Exit Sub
        End If
    End Select
    KeyAscii = 0
End Sub


Private Sub txt�ӳ���_LostFocus()
    Dim cur�۸� As Double

    Me.txt�ӳ���.Text = Format(txt�ӳ���.Text, "0.00")
End Sub

Private Sub txt�����_GotFocus()
    Me.txt�����.SelStart = 0: Me.txt�����.SelLength = 100
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt�����.SelLength = Len(txt�����.Text) Then Exit Sub
            If Len(Mid(txt�����, InStr(1, txt�����.Text, ".") + 1)) >= mintCostDigit And txt�����.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�����_LostFocus()
    Me.txt�����.Text = FormatEx(Val(txt�����), mintCostDigit, , True)
End Sub

Private Sub txt����_Change()
    Me.txt�����.Text = FormatEx(Val(Me.txtָ������.Text) * Val(Me.txt����.Text) / 100, mintCostDigit, , True)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
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

Private Sub txt����_LostFocus()
    Me.txt����.Text = Format(txt����, "0.00000")
End Sub

Private Sub txt�����װ_GotFocus()
    Me.txt�����װ.SelStart = 0: Me.txt�����װ.SelLength = 100
End Sub

Private Sub txt�����װ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If InStr(txt�����װ.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    'ֻ�ܴ���һ��С����
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txt�����װ.SelLength = Len(txt�����װ.Text) Then Exit Sub
            If Len(Mid(txt�����װ, InStr(1, txt�����װ.Text, ".") + 1)) >= 5 And txt�����װ.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���ﵥλ_Change()
    Me.lbl�����װ.Caption = "(1" & Me.txt���ﵥλ.Text & "="
    Call cbo���쵥λ_Click
End Sub

Private Sub txt���ﵥλ_GotFocus()
    Me.txt���ﵥλ.SelStart = 0: Me.txt���ﵥλ.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt���ﵥλ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���ﵥλ_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt��׼�ĺ�_GotFocus()
    Me.txt��׼�ĺ�.SelStart = 0: Me.txt��׼�ĺ�.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt��׼�ĺ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��׼�ĺ�_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txtƴ��_GotFocus()
    Me.txtƴ��.SelStart = 0: Me.txtƴ��.SelLength = 100
End Sub

Private Sub txtƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim Count As Integer
    
    If KeyAscii = vbKeyReturn Then
        If txtDDDֵ.Visible = True Then
            Call OS.PressKey(vbKeyTab)
        Else
            stbSpec.Tab = 1
            If cboҩ������.Enabled = True Then
                cboҩ������.SetFocus
            End If
        End If
        Exit Sub
    End If
    strText = Me.txt����.Text
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13 Then
        If strText <> "" Then
            If KeyAscii = 46 Then
                Count = (Len(strText) - Len(Replace(strText, ".", ""))) / Len(".")
                
                If Count > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Else
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
    strText = ""
    
'    If KeyAscii = vbKeyReturn Then
'        Me.stbSpec.Tab = 1
'        If Me.cboҩ������.Enabled Then
'            Me.cboҩ������.SetFocus
'        Else
'            Me.txtָ������.SetFocus
'        End If
'    End If
End Sub

Private Sub txt�ͻ���װ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ͻ���λ_Change()
    Me.lbl�ͻ���װ.Caption = "(1" & Me.txt�ͻ���λ.Text & "="
    If Trim(txt�ͻ���λ.Text) <> "" Then
        txt�ͻ���װ.Enabled = True
    Else
        txt�ͻ���װ.Enabled = False
        txt�ͻ���װ.Text = ""
    End If
End Sub

Private Sub txt�ͻ���λ_GotFocus()
    Me.txt�ͻ���λ.SelStart = 0: Me.txt�ͻ���λ.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt�ͻ���λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ʒ��_Change()
    Dim strTmp As String
    '���¼�����ƣ���ȥ �������ַ�
    strTmp = MoveSpecialChar(txt��Ʒ��.Text)
    If txt��Ʒ��.Text <> strTmp Then
        txt��Ʒ��.Text = strTmp
    End If
    Me.txtƴ��.Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
    Me.txt���.Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
End Sub

Private Sub txt��Ʒ��_GotFocus()
    Me.txt��Ʒ��.SelStart = 0: Me.txt��Ʒ��.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt��Ʒ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
        Case vbKeyReturn
            Call OS.PressKey(vbKeyTab)
    End Select
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Me.txtƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ʒ��.Text, False, mlng���볤��)
    Me.txt���.Text = zlStr.GetCodeByORCL(Me.txt��Ʒ��.Text, True, mlng���볤��)

End Sub

Private Sub txt��Ʒ��_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt���췧ֵ_GotFocus()
    txt���췧ֵ.SelStart = 0: txt���췧ֵ.SelLength = Len(txt���췧ֵ)
End Sub

Private Sub txt���췧ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
'    If KeyAscii = vbKeyReturn Then
'        Me.stbSpec.Tab = 1
'        If Me.cboҩ������.Enabled Then
'            Me.cboҩ������.SetFocus
'        Else
'            Me.txtָ������.SetFocus
'        End If
'    End If
End Sub

Private Sub txt�ۼ۵�λ_Change()
    Me.lbl����ϵ��.Caption = "(1" & Me.txt�ۼ۵�λ.Text & "="
    If glngSys \ 100 = 8 Then
        Me.txt���ﵥλ = Me.txt�ۼ۵�λ
        Me.txtסԺ��λ = Me.txt�ۼ۵�λ
    End If
    Me.lblסԺ��λChild.Caption = Me.txt�ۼ۵�λ & ")"
    Me.lbl���ﵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
    Me.lblҩ�ⵥλChild.Caption = Me.txt�ۼ۵�λ & ")"
    Me.lbl���쵥λ.Caption = Me.txt�ۼ۵�λ & ")"
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
    Call OS.OpenIme(True)
End Sub

Private Sub txt�ۼ۵�λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ۼ۵�λ_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txt������_GotFocus()
    txt������.MaxLength = Val(zlDatabase.GetPara("������", glngSys, 1023, 7))
    Me.txt������.SelStart = 0: Me.txt������.SelLength = 100
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
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

Private Sub txt��λ��_GotFocus()
    txt��λ��.MaxLength = Val(zlDatabase.GetPara("��λ��", glngSys, 1023, 20))
    Me.txt��λ��.SelStart = 0: Me.txt��λ��.SelLength = 100
End Sub

Private Sub txt��λ��_KeyPress(KeyAscii As Integer)
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

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
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
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtЧ��_GotFocus()
    Me.txtЧ��.SelStart = 0: Me.txtЧ��.SelLength = 100
End Sub

Private Sub txtЧ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        If stbSpec.TabVisible(2) = True Then
            stbSpec.Tab = 2
            cboTemperature.SetFocus
        Else
            Call OS.PressKey(vbKeyTab)
        End If
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtҩ���װ_GotFocus()
    Me.txtҩ���װ.SelStart = 0: Me.txtҩ���װ.SelLength = 100
End Sub

Private Sub txtҩ���װ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
        
    Case Else
        If InStr(txtҩ���װ.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    'ֻ�ܴ���һ��С����
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txtҩ���װ.SelLength = Len(txtҩ���װ.Text) Then Exit Sub
            If Len(Mid(txtҩ���װ, InStr(1, txtҩ���װ.Text, ".") + 1)) >= 5 And txtҩ���װ.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txtҩ�ⵥλ_Change()
    Me.lblҩ���װ.Caption = "(1" & Me.txtҩ�ⵥλ.Text & "="
    Me.lbl�ͻ���λchild.Caption = Me.txtҩ�ⵥλ.Text & ")"
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
    Call OS.OpenIme(True)
End Sub

Private Sub txtҩ�ⵥλ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtҩ�ⵥλ_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txtָ������_Change()
    Me.txt�����.Text = FormatEx(Val(Me.txtָ������.Text) * Val(Me.txt����.Text) / 100, mintCostDigit, , True)
End Sub

Private Sub txtָ������_GotFocus()
    Me.txtָ������.SelStart = 0: Me.txtָ������.SelLength = 100
End Sub

Private Sub txtָ������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txtָ������.SelLength = Len(txtָ������.Text) Then Exit Sub
            If Len(Mid(txtָ������, InStr(1, txtָ������.Text, ".") + 1)) >= mintCostDigit And txtָ������.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub


Private Sub txtָ������_LostFocus()
    Me.txtָ������.Text = FormatEx(Val(txtָ������.Text), mintCostDigit, , True)
End Sub

Private Sub txtָ���ۼ�_GotFocus()
    Me.txtָ���ۼ�.SelStart = 0: Me.txtָ���ۼ�.SelLength = 100
End Sub

Private Sub txtָ���ۼ�_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txtָ���ۼ�.SelLength = Len(txtָ���ۼ�.Text) Then Exit Sub
            If Len(Mid(txtָ���ۼ�, InStr(1, txtָ���ۼ�.Text, ".") + 1)) >= mintPriceDigit And txtָ���ۼ�.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txtָ���ۼ�_LostFocus()
    Me.txtָ���ۼ�.Text = FormatEx(Val(txtָ���ۼ�), mintPriceDigit, , True)
End Sub

Private Sub txtסԺ��װ_GotFocus()
    Me.txtסԺ��װ.SelStart = 0: Me.txtסԺ��װ.SelLength = 100
End Sub

Private Sub txtסԺ��װ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If InStr(txtסԺ��װ.Text, ".") <> 0 And Chr(KeyAscii) = "." Then    'ֻ�ܴ���һ��С����
            KeyAscii = 0
            Exit Sub
        End If
            
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If txtסԺ��װ.SelLength = Len(txtסԺ��װ.Text) Then Exit Sub
            If Len(Mid(txtסԺ��װ, InStr(1, txtסԺ��װ.Text, ".") + 1)) >= 5 And txtסԺ��װ.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txtסԺ��λ_Change()
    Me.lblסԺ��װ.Caption = "(1" & Me.txtסԺ��λ.Text & "="
    Call cbo���쵥λ_Click
End Sub

Private Sub txtסԺ��λ_GotFocus()
    Me.txtסԺ��λ.SelStart = 0: Me.txtסԺ��λ.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txtסԺ��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtסԺ��λ_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub stbSpec_Click(PreviousTab As Integer)
   
   Select Case stbSpec.Tab
    Case 0
        If Me.txt����.Enabled Then Me.txt����.SetFocus
    Case 1
'        If Me.txtָ������.Enabled Then Me.txtָ������.SetFocus
        If Me.cboҩ������.Enabled Then Me.cboҩ������.SetFocus
    End Select
End Sub

Private Function zlSureManufacturer() As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж��Ƿ�ɼ������������̣������̱����ֶο��Ϊ:10��
    '-------------------------------------------------------------
    On Error GoTo errHandle
    zlSureManufacturer = False
    
    gstrSql = "Select Max(����) ���� From ҩƷ������"
'        If .State = adStateOpen Then .Close
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
    
    With rsTemp
        If .EOF Then zlSureManufacturer = True: Exit Function
        If IsNull(rsTemp!����) Then zlSureManufacturer = True: Exit Function
        
        '����������˳�
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetDigitSign(ByVal lngMediId As Long, ByVal strSpec As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ҩƷͨ�����ơ����͵����ֱ����͹��ǰ��λ��ֵ����������ҩƷ��λ��
    '��Σ�strSpellcode-ͨ�����Ƶ�ƴ���룻strDoseCode:���͵����ֱ����, strSpec�������ֵ
    '���أ�ҩƷ����
    '-------------------------------------------------------------
    Dim rsThis As New ADODB.Recordset
    Dim strSpellcode As String, strDoseCode As String
    Dim strChange As String
    Dim intLocate As Integer
    
    On Error GoTo errHandle
    gstrSql = "Select ���� From ������Ŀ���� where ������Ŀid=[1] and ����=1 and ����=1"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strSpellcode = IIf(IsNull(rsThis!����), "", rsThis!����)
    Else
        strSpellcode = ""
    End If
    
    gstrSql = "select P.����� from ҩƷ���� T,ҩƷ���� P where T.ҩƷ����=P.����(+) and ҩ��id=[1]"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strDoseCode = IIf(IsNull(rsThis!�����), "", rsThis!�����)
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Private Sub txtע���̱�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Function CheckUnit() As Boolean
    Dim intOut As Integer, intIN As Integer
    Dim arr��λ, arrϵ��
    Dim str��λ As String, strϵ�� As String
    Dim str��λ_Tmp As String, strϵ��_Tmp As String
    
    '����Ƿ���ڵ�λ����һ������ϵ����һ�µ����
    '����Ƿ����ϵ��һ��������λ���Ʋ�һ�������
    str��λ = txt�ۼ۵�λ.Text & "|" & txtסԺ��λ.Text & "|" & txt���ﵥλ.Text & "|" & txtҩ�ⵥλ.Text
    strϵ�� = txt����ϵ��.Text & "|" & txtסԺ��װ.Text & "|" & txt�����װ.Text & "|" & txtҩ���װ.Text
    
    '���ǵ�������λ�������ۼ۵�λһ�£���ϵ���϶���һ�£����Ա���ֿ��ж�
    '���ۼ۵�λ��ļ��
    For intOut = 2 To 4
        str��λ_Tmp = IIf(intOut = 1, txt�ۼ۵�λ.Text, IIf(intOut = 2, txtסԺ��λ.Text, IIf(intOut = 3, txt���ﵥλ.Text, txtҩ�ⵥλ.Text)))
        strϵ��_Tmp = Val(IIf(intOut = 1, txt����ϵ��.Text, IIf(intOut = 2, txtסԺ��װ.Text, IIf(intOut = 3, txt�����װ.Text, txtҩ���װ.Text))))
        arr��λ = Split(str��λ, "|")
        arrϵ�� = Split(strϵ��, "|")
        For intIN = 2 To 4
            If intIN <> intOut Then
                '��λ��ͬϵ����ͬ
                If str��λ_Tmp = arr��λ(intIN - 1) And (Val(strϵ��_Tmp) <> Val(arrϵ��(intIN - 1))) Then
                    MsgBox IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��λ��" & IIf(intIN = 2, "סԺ", IIf(intIN = 3, "����", "ҩ��")) & "��λһ�£�����ϵ��ȴ����ͬ�����飡", vbInformation, gstrSysName
                    Exit Function
                End If
                If str��λ_Tmp <> arr��λ(intIN - 1) And (Val(strϵ��_Tmp) = Val(arrϵ��(intIN - 1))) Then
                    MsgBox IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��װ��" & IIf(intIN = 2, "סԺ", IIf(intIN = 3, "����", "ҩ��")) & "��װһ�£����䵥λȴ����ͬ�����飡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    Next
    
    '����������λ���ۼ۵�λ��ͬ����ϵ����Ϊ1�����
    '����λ���ۼ۵�λ���м��
    For intOut = 2 To 4
        str��λ_Tmp = IIf(intOut = 1, txt�ۼ۵�λ.Text, IIf(intOut = 2, txtסԺ��λ.Text, IIf(intOut = 3, txt���ﵥλ.Text, txtҩ�ⵥλ.Text)))
        strϵ��_Tmp = Val(IIf(intOut = 1, txt����ϵ��.Text, IIf(intOut = 2, txtסԺ��װ.Text, IIf(intOut = 3, txt�����װ.Text, txtҩ���װ.Text))))
        If str��λ_Tmp = txt�ۼ۵�λ.Text And Val(strϵ��_Tmp) <> 1 Then
            MsgBox IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��λ���ۼ۵�λһ�£�" & IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "ϵ��Ӧ��Ϊ1", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    CheckUnit = True
End Function

Private Function CheckRequest() As Boolean
    Dim dbl�������� As Double
    Dim str�������� As String
    '������췧ֵת��Ϊ���۵�λ���Ƿ�Ϊ����������5λС������ʾ����Ա����ǿ�Ʊ���
    dbl�������� = Val(txt���췧ֵ.Text)
    
    Select Case cbo���쵥λ.ListIndex
    Case 1 'סԺ��λ
        dbl�������� = dbl�������� * Val(txtסԺ��װ.Text)
    Case 2 '���ﵥλ
        dbl�������� = dbl�������� * Val(txt�����װ.Text)
    Case 3 'ҩ�ⵥλ
        dbl�������� = dbl�������� * Val(txtҩ���װ.Text)
    End Select
    txt���췧ֵ.Tag = dbl��������
    
    CheckRequest = True
End Function

Private Sub txtע���̱�_KeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub UploadDrugInfo(ByVal lngDrugId As Long)
'ͬ���ϴ�ҩƷ��Ϣ
    If Not gobjLogisticPlatform Is Nothing And lngDrugId <> 0 Then
        gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngDrugId
    End If
End Sub


Private Sub vsfItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 0 Then Exit Sub
    With vsfItem
        If Col <> .ColIndex("����") Then Exit Sub
        If .TextMatrix(Row, .ColIndex("����")) <> .TextMatrix(Row, .ColIndex("ԭ����")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("����")) = vbRed
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("����")) = vbBlack
        End If
    End With
End Sub

Private Sub vsfItem_EnterCell()
    With vsfItem
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .Col = .ColIndex("����") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub


Private Sub vsfItem_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfItem
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("����") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("����")
            End If
        End If
    End With
End Sub


Private Sub vsfItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsfItem.ColIndex("����") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


