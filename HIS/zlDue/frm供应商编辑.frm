VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm��Ӧ�̱༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ӧ�̱༭"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��Ƭ 
      Caption         =   "������Ƭ(&F)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmd��Ƭ 
      Caption         =   "�����Ƭ(&L)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Index           =   1
      Left            =   -180
      TabIndex        =   36
      Top             =   645
      Width           =   10155
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   8880
      Width           =   13185
   End
   Begin TabDlg.SSTab sstab 
      Height          =   7935
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ(&0)"
      TabPicture(0)   =   "frm��Ӧ�̱༭.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkĩ��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pic����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkCodeLen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "������Ϣ(&1)"
      TabPicture(1)   =   "frm��Ӧ�̱༭.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEdit(9)"
      Tab(1).Control(1)=   "lblEdit(10)"
      Tab(1).Control(2)=   "lblEdit(8)"
      Tab(1).Control(3)=   "lblEdit(7)"
      Tab(1).Control(4)=   "lblEdit(6)"
      Tab(1).Control(5)=   "lblEdit(0)"
      Tab(1).Control(6)=   "lblEdit(4)"
      Tab(1).Control(7)=   "lblEdit(5)"
      Tab(1).Control(8)=   "Lbl���֤��"
      Tab(1).Control(9)=   "Lbl���֤Ч��"
      Tab(1).Control(10)=   "lblִ�պ�"
      Tab(1).Control(11)=   "Lblִ��Ч��"
      Tab(1).Control(12)=   "lbl(0)"
      Tab(1).Control(13)=   "lbl(1)"
      Tab(1).Control(14)=   "Label2"
      Tab(1).Control(15)=   "Label3"
      Tab(1).Control(16)=   "dtp��Ȩ��"
      Tab(1).Control(17)=   "Dtpִ��Ч��"
      Tab(1).Control(18)=   "Dtp���֤Ч��"
      Tab(1).Control(19)=   "TxtEdit(11)"
      Tab(1).Control(20)=   "TxtEdit(10)"
      Tab(1).Control(21)=   "TxtEdit(3)"
      Tab(1).Control(22)=   "TxtEdit(2)"
      Tab(1).Control(23)=   "TxtEdit(8)"
      Tab(1).Control(24)=   "TxtEdit(7)"
      Tab(1).Control(25)=   "TxtEdit(9)"
      Tab(1).Control(26)=   "TxtEdit(4)"
      Tab(1).Control(27)=   "TxtEdit(5)"
      Tab(1).Control(28)=   "TxtEdit(6)"
      Tab(1).Control(29)=   "TxtEdit(15)"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "����������Ϣ(&2)"
      TabPicture(2)   =   "frm��Ӧ�̱༭.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl��ӪƷ��"
      Tab(2).Control(1)=   "lbl��ע"
      Tab(2).Control(2)=   "txt��ӪƷ��"
      Tab(2).Control(3)=   "txt��ע"
      Tab(2).Control(4)=   "Picture2"
      Tab(2).Control(5)=   "Picture3"
      Tab(2).Control(6)=   "Picture4"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "���֤��Ƭ(&3)"
      TabPicture(3)   =   "frm��Ӧ�̱༭.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic��Ƭ(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "ִ�պ���Ƭ(&4)"
      TabPicture(4)   =   "frm��Ӧ�̱༭.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pic��Ƭ(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "��Ȩ����Ƭ(&5)"
      TabPicture(5)   =   "frm��Ӧ�̱༭.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "pic��Ƭ(2)"
      Tab(5).ControlCount=   1
      Begin VB.CheckBox chkCodeLen 
         Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
         Height          =   285
         Left            =   480
         TabIndex        =   87
         Top             =   3840
         Width           =   4290
      End
      Begin VB.PictureBox pic��Ƭ 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   81
         Top             =   480
         Width           =   12375
         Begin VB.Image img��Ƭ 
            Height          =   1650
            Index           =   2
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox pic��Ƭ 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   80
         Top             =   480
         Width           =   12375
         Begin VB.Image img��Ƭ 
            Height          =   1650
            Index           =   1
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox pic��Ƭ 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   79
         Top             =   480
         Width           =   12375
         Begin VB.Image img��Ƭ 
            Height          =   1650
            Index           =   0
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   73
         Top             =   1920
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   14
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   75
            Tag             =   "ҩ��ֱ�����Ϣ�е�֤��"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker dtpҩ��ֱ��� 
            Height          =   300
            Left            =   4680
            TabIndex        =   74
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy��MM��dd��"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label7 
            Caption         =   "ҩ��ֱ�����Ϣ"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label lbl 
            Caption         =   "֤  ��(&V)"
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   77
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   4050
            TabIndex        =   76
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   67
         Top             =   1200
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   13
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   68
            Tag             =   "������֤��Ϣ�е�֤��"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker dtp������֤ 
            Height          =   300
            Left            =   4680
            TabIndex        =   69
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy��MM��dd��"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label6 
            Caption         =   "������֤��Ϣ"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   4050
            TabIndex        =   71
            Top             =   300
            Width           =   630
         End
         Begin VB.Label lbl 
            Caption         =   "֤  ��(&Z)"
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   70
            Top             =   285
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   61
         Top             =   480
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   12
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   62
            Tag             =   "ί��������"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker Dtpί�������� 
            Height          =   300
            Left            =   4680
            TabIndex        =   63
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy��MM��dd��"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label5 
            Caption         =   "������Աί������Ϣ"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "��  ��(&N)"
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   315
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   4050
            TabIndex        =   64
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   240
         ScaleHeight     =   1335
         ScaleWidth      =   7455
         TabIndex        =   54
         Top             =   2520
         Width           =   7455
         Begin VB.CheckBox chkType 
            Caption         =   "��������(&W)"
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   55
            Top             =   360
            Width           =   1410
         End
         Begin VB.CheckBox chkType 
            Caption         =   "ҩƷ(&Y)"
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   59
            Top             =   0
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "����(&M)"
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   58
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "�豸(&J)"
            Height          =   315
            Index           =   2
            Left            =   915
            TabIndex        =   57
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "����(&E)"
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   56
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox pic���� 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1935
         ScaleWidth      =   7455
         TabIndex        =   41
         Top             =   480
         Width           =   7455
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   1
            Left            =   915
            MaxLength       =   10
            TabIndex        =   48
            Tag             =   "����"
            Top             =   1110
            Width           =   1905
         End
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   0
            Left            =   915
            MaxLength       =   80
            TabIndex        =   46
            Tag             =   "����"
            Top             =   750
            Width           =   5655
         End
         Begin VB.TextBox txtParent 
            Height          =   300
            Left            =   915
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   45
            Top             =   0
            Width           =   5385
         End
         Begin VB.CommandButton cmd�ϼ� 
            Caption         =   "&P"
            Height          =   300
            Left            =   6300
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1500
            Width           =   1905
         End
         Begin VB.TextBox txtCode 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "����"
            Text            =   "111111"
            Top             =   420
            Width           =   1755
         End
         Begin VB.TextBox txtUpCode 
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   915
            MaxLength       =   10
            TabIndex        =   47
            TabStop         =   0   'False
            Tag             =   "����"
            Text            =   "11"
            Top             =   375
            Width           =   1905
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�ϼ�(&U)"
            Height          =   180
            Index           =   11
            Left            =   240
            TabIndex        =   53
            Top             =   75
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&D)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   435
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&N)"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   51
            Top             =   795
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&S)"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   50
            Top             =   1155
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "Ժ��(&B)"
            Height          =   180
            Left            =   240
            TabIndex        =   49
            Top             =   1545
            Width           =   630
         End
      End
      Begin VB.TextBox txt��ע 
         Height          =   2265
         Left            =   -73560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   3240
         Width           =   5805
      End
      Begin VB.TextBox txt��ӪƷ�� 
         Height          =   585
         Left            =   -73560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   2520
         Width           =   5805
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   15
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   17
         Tag             =   "��Ȩ��"
         Top             =   2070
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   6
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   13
         Tag             =   "ִ�պ�"
         Top             =   1680
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   5
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "���֤��"
         Top             =   1290
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   4
         Left            =   -69900
         MaxLength       =   16
         TabIndex        =   7
         Tag             =   "�绰"
         Top             =   900
         Width           =   2625
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   9
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "��ַ"
         Top             =   3255
         Width           =   6450
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   7
         Left            =   -69900
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "˰��ǼǺ�"
         Top             =   510
         Width           =   2640
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   8
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "��������"
         Top             =   2880
         Width           =   6450
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   2
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "�ʺ�"
         Top             =   510
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   3
         Left            =   -73740
         MaxLength       =   20
         TabIndex        =   5
         Tag             =   "��ϵ��"
         Top             =   900
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   10
         Left            =   -73740
         MaxLength       =   6
         TabIndex        =   21
         Tag             =   "������"
         Top             =   2460
         Width           =   2055
      End
      Begin VB.TextBox TxtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   -69900
         MaxLength       =   8
         TabIndex        =   24
         Tag             =   "���ö�"
         Top             =   2460
         Width           =   2430
      End
      Begin VB.CheckBox chkĩ�� 
         Caption         =   "ĩ��(&M)"
         Height          =   180
         Left            =   510
         TabIndex        =   33
         Top             =   8040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker Dtp���֤Ч�� 
         Height          =   300
         Left            =   -69900
         TabIndex        =   11
         Top             =   1290
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy��MM��dd��"
         DateIsNull      =   -1  'True
         Format          =   141033475
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker Dtpִ��Ч�� 
         Height          =   300
         Left            =   -69900
         TabIndex        =   15
         Top             =   1680
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy��MM��dd��"
         DateIsNull      =   -1  'True
         Format          =   141099011
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker dtp��Ȩ�� 
         Height          =   300
         Left            =   -69900
         TabIndex        =   19
         Top             =   2070
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy��MM��dd��"
         DateIsNull      =   -1  'True
         Format          =   141033475
         CurrentDate     =   37994
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         Caption         =   "��  ע(&B)"
         Height          =   180
         Left            =   -74640
         TabIndex        =   40
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label lbl��ӪƷ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ӪƷ��(&S)"
         Height          =   180
         Left            =   -74640
         TabIndex        =   39
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ȩ��(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74550
         TabIndex        =   16
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ȩ��(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -70740
         TabIndex        =   18
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   -71670
         TabIndex        =   22
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ԫ"
         Height          =   180
         Index           =   0
         Left            =   -67440
         TabIndex        =   25
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label Lblִ��Ч�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ִ��Ч��(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -70920
         TabIndex        =   14
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label lblִ�պ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�պ�(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74550
         TabIndex        =   12
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Lbl���֤Ч�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤Ч��(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71100
         TabIndex        =   10
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label Lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74730
         TabIndex        =   8
         Top             =   1350
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�绰(&T)"
         Height          =   180
         Index           =   5
         Left            =   -70560
         TabIndex        =   6
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ַ(&A)"
         Height          =   180
         Index           =   4
         Left            =   -74370
         TabIndex        =   28
         Top             =   3315
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "˰��ǼǺ�(&K)"
         Height          =   180
         Index           =   0
         Left            =   -71100
         TabIndex        =   2
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&G)"
         Height          =   180
         Index           =   6
         Left            =   -74730
         TabIndex        =   26
         Top             =   2940
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��  ��(&Z)"
         Height          =   180
         Index           =   7
         Left            =   -74550
         TabIndex        =   0
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ϵ��(&L)"
         Height          =   180
         Index           =   8
         Left            =   -74550
         TabIndex        =   4
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&Y)"
         Height          =   180
         Index           =   10
         Left            =   -74550
         TabIndex        =   20
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ö�(&E)"
         Height          =   180
         Index           =   9
         Left            =   -70740
         TabIndex        =   23
         Top             =   2520
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   10080
      TabIndex        =   30
      Top             =   9120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   11400
      TabIndex        =   31
      Top             =   9120
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog cdl��Ƭ 
      Left            =   7920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblͼƬ˵�� 
      Height          =   210
      Index           =   2
      Left            =   3120
      TabIndex        =   86
      Top             =   9180
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label lblͼƬ˵�� 
      Height          =   210
      Index           =   1
      Left            =   3120
      TabIndex        =   85
      Top             =   9187
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label lblͼƬ˵�� 
      Height          =   210
      Index           =   0
      Left            =   3120
      TabIndex        =   84
      Top             =   9187
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ҩƷ(���ʡ��豸��)��Ӧ�̽����������޸ĵ���.ͬʱ�ɼӳ���������б���ĳ��ȡ�"
      Height          =   180
      Left            =   600
      TabIndex        =   35
      Top             =   345
      Width           =   6930
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frm��Ӧ�̱༭.frx":00A8
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frm��Ӧ�̱༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String         '��ǰ�༭�ĵ�λID
Dim mlng�ϼ�id As Long       '�ϼ���λID
Dim mintSuccess As Integer
Dim mintEditType As gEditType    '�༭����
Dim mblnChange As Boolean
Dim mstrPrivs As String         'Ȩ�޴�
Const mintMaxLen = 8        '���볤��
Dim mblnFist As Boolean

Private Enum picType
    ���֤��Ƭ = 0
    ִ����Ƭ = 1
    ��Ȩ����Ƭ = 2
End Enum

Private Type picCon
    mblnExistPic(0 To 2) As Boolean     '��ǰ�Ƿ���ͼƬ��Ϣ
    mblnIsModify(0 To 2) As Boolean     '����Ƭ��������ʱ��ΪTrue
End Type
Private myPicCon As picCon

Private Sub InitDefaultLen()
    '-----------------------------------------------------------------------------------------------------------
    '����:���ñ༭��Ĭ�ϳ���
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-10-23 14:31:25
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, j As Long
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select ˰��ǼǺ�,���֤��,ִ�պ�,��Ȩ�� From ��Ӧ�� where id=0"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 0 To rsTemp.Fields.Count - 1
        For j = 0 To TxtEdit.UBound
            If rsTemp.Fields(i).Name = TxtEdit(j).Tag Then
                TxtEdit(j).MaxLength = rsTemp.Fields(i).DefinedSize
                Exit For
            End If
        Next
    Next
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub chkCodeLen_Click()
    If chkCodeLen.Visible = False Then Exit Sub
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    If sstab.Tab = 0 Then
        If Me.txtCode.Enabled Then txtCode.SetFocus
    End If
End Sub

Private Sub chkCodeLen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    mblnChange = True
    setCtlEn
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
    Case 3
        '�������һҳ��Ϣ,������
        sstab.Tab = 1
                
        If TxtEdit(2).Enabled Then
            TxtEdit(2).SetFocus
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub CmdCancel_Click()
    Dim blnYes As Boolean
    If mblnChange = False Then
        Unload Me
        Exit Sub
    End If
    ShowMsgbox "���Ѿ������˵�����Ϣ,�������˳��Ļ�," & vbCrLf & "�����ĵ����ݽ����ܱ���,���Ҫ�˳���?", True, blnYes
    If blnYes = True Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub cmdOK_Click()
    Dim intIndex As Integer
    
    If IsValid() = False Then Exit Sub
    If Save��λ() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    mstrID = ""
    For intIndex = 0 To 15
        TxtEdit(intIndex).Text = ""
    Next
    zlChangeCode "��Ӧ��", mlng�ϼ�id, txtUpCode, txtCode, chkCodeLen, Me.Caption
     mblnChange = False
    sstab.Tab = 0
    If TxtEdit(0).Enabled And TxtEdit(0).Visible Then
        TxtEdit(0).SetFocus
    End If
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    
    Dim strTemp As String
    
    strTemp = Trim(txtCode.Text)
     
    If strTemp = "" Then
        ShowMsgbox "�����������!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If InStr(1, strTemp, "'") <> 0 Then
        ShowMsgbox "���벻�����뵥����!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(strTemp) Then
        ShowMsgbox "����������������,������!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If Len(txtUpCode.Text & strTemp) > 8 Then
        ShowMsgbox "���볤�Ȳ��ܳ���8λ,������!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If LenB(txt��ӪƷ��.Text) > 200 Then
        ShowMsgbox "��ӪƷ�ֲ��ܳ���200λ�ַ���100������,������!"
        sstab.Tab = 2
        If txt��ӪƷ��.Enabled Then txt��ӪƷ��.SetFocus
        Exit Function
    End If
    If LenB(txt��ע.Text) > 200 Then
        ShowMsgbox "��ע���ܳ���200λ�ַ���100������,������!"
        sstab.Tab = 2
        If txt��ע.Enabled Then txt��ע.SetFocus
        Exit Function
    End If
    For intIndex = 0 To 15
        strTemp = Trim(TxtEdit(intIndex).Text)
        If intIndex = 0 Then
            If strTemp = "" Then
                ShowMsgbox TxtEdit(intIndex).Tag & "��������!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > TxtEdit(intIndex).MaxLength Then
                ShowMsgbox TxtEdit(intIndex).Tag & "����,���������" & TxtEdit(intIndex).MaxLength / 2 & "�����ֻ�" & TxtEdit(intIndex).MaxLength & "���ַ�!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox TxtEdit(intIndex).Tag & "�������뵥����!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
            
            Select Case TxtEdit(intIndex).Tag
            Case "������", "���ö�"
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox TxtEdit(intIndex).Tag & "����������,������!"
                    If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If TxtEdit(intIndex).Tag = "���ö�" Then
                    If Val(strTemp) > 99999999 Then
                        ShowMsgbox TxtEdit(intIndex).Tag & "������99999999,������!"
                        If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(strTemp) > 999999 Then
                        ShowMsgbox TxtEdit(intIndex).Tag & "������999999,������!"
                        If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                        Exit Function
                    End If
                End If
            End Select
        End If
    Next
    Dim blnTrue As Boolean
    Dim i As Byte
    For i = 0 To 4
        If chkType(i).Value = 1 Then
            blnTrue = True
        End If
    Next
    If blnTrue = False And chkĩ��.Value = 1 Then
        ShowMsgbox "ûѡ�����������ֹ�Ӧ��,��ѡ��!"
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save��λ() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:����ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPriType As Long
    Dim lngID As Long
    Dim strTmp As String
    Dim intIndex As Integer
    Dim blnTran As Boolean
    Dim i As Integer
    Dim rsDepend As New ADODB.Recordset
    
    On Error GoTo errHandle
    strTmp = ""
    For intIndex = 0 To 4
        strTmp = strTmp & IIf(chkType(intIndex).Value = 1, 1, 0)
    Next
    
    If mintEditType = g���� Then
        gstrSQL = "Select ���� From ��Ӧ�� "
    ElseIf mintEditType = g�޸� Then
        gstrSQL = "Select ���� From ��Ӧ�� Where ID <> [1]  "
    End If
    
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "��������Ƿ��ظ�", Val(mstrID))
    Do While Not rsDepend.EOF
        If TxtEdit(0) = rsDepend!���� Then
            MsgBox "�����ظ����������룡", vbInformation, gstrSysName
            Exit Function
        End If
        rsDepend.MoveNext
    Loop
        
    If mstrID = "" Then
        lngID = zldatabase.GetNextId("��Ӧ��")
        gstrSQL = "zl_��Ӧ��_insert ( "
    Else
        lngID = Val(mstrID)
        gstrSQL = "zl_��Ӧ��_update ( "
    End If
    
    '���̲�������:
    '   ID_IN,�ϼ�ID_IN,����_IN,����_IN,����_IN,��ַ_IN,�绰_IN,��������_IN,�ʺ�_IN,��ϵ��_IN,
    '   ˰��ǼǺ�_IN,���֤��_IN,���֤Ч��_IN,ִ�պ�_IN,ִ��Ч��_IN,��Ȩ��_IN,��Ȩ��_IN,��Ӧ������_IN,������_IN,
    '   ���ö�_IN,����ί����_IN ,����ί������_IN,������֤��_IN,������֤����_IN, ҩ��ֱ�����_IN,ҩ��ֱ�������_IN
    '     վ��_In,ĩ��_IN , �ı���볤��,��ӪƷ��_In,��ע_In
    
    gstrSQL = gstrSQL & "" & _
            lngID & "," & _
            IIf(mlng�ϼ�id = 0, "Null", mlng�ϼ�id) & ",'" & _
            txtUpCode.Text & txtCode.Text & "','" & _
            Trim(TxtEdit(0).Text) & "','" & _
            Trim(TxtEdit(1).Text) & "'," & _
            IIf(Trim(TxtEdit(9).Text) = "", "NULL", "'" & Trim(TxtEdit(9).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(4).Text) = "", "NULL", "'" & Trim(TxtEdit(4).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(8).Text) = "", "NULL", "'" & Trim(TxtEdit(8).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(2).Text) = "", "NULL", "'" & Trim(TxtEdit(2).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(3).Text) = "", "NULL", "'" & Trim(TxtEdit(3).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(7).Text) = "", "NULL", "'" & Trim(TxtEdit(7).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(5).Text) = "", "NULL", "'" & Trim(TxtEdit(5).Text) & "'") & "," & _
            IIf(Dtp���֤Ч��.Value = "" Or IsNull(Dtp���֤Ч��.Value), "NULL", "to_Date('" & Format(Dtp���֤Ч��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(6).Text) = "", "NULL", "'" & Trim(TxtEdit(6).Text) & "'") & "," & _
            IIf(Dtpִ��Ч��.Value = "" Or IsNull(Dtpִ��Ч��.Value), "NULL", "to_Date('" & Format(Dtpִ��Ч��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(15).Text) = "", "NULL", "'" & Trim(TxtEdit(15).Text) & "'") & "," & _
            IIf(dtp��Ȩ��.Value = "" Or IsNull(dtp��Ȩ��.Value), "NULL", "to_Date('" & Format(dtp��Ȩ��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            "'" & strTmp & "'," & _
            IIf(Trim(TxtEdit(10).Text) = "", "NULL", Val(TxtEdit(10).Text)) & "," & _
            IIf(Trim(TxtEdit(11).Text) = "", "NULL", Val(TxtEdit(11).Text)) & ","
        gstrSQL = gstrSQL & _
            IIf(Trim(TxtEdit(12).Text) = "", "NULL", "'" & Trim(TxtEdit(12).Text) & "'") & "," & _
            IIf(Dtpί��������.Value = "" Or IsNull(Dtpί��������.Value), "NULL", "to_Date('" & Format(Dtpί��������.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(13).Text) = "", "NULL", "'" & Trim(TxtEdit(13).Text) & "'") & "," & _
            IIf(dtp������֤.Value = "" Or IsNull(dtp������֤.Value), "NULL", "to_Date('" & Format(dtp������֤.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(14).Text) = "", "NULL", "'" & Trim(TxtEdit(14).Text) & "'") & "," & _
            IIf(dtpҩ��ֱ���.Value = "" Or IsNull(dtpҩ��ֱ���.Value), "NULL", "to_Date('" & Format(dtpҩ��ֱ���.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & Me.cmbStationNo.ItemData(Me.cmbStationNo.ListIndex) & "'", "NULL") & "," & chkĩ��.Value & "," & _
            IIf(Me.chkCodeLen.Value = 1, 1, 0) & "," & _
            IIf(Trim(txt��ӪƷ��.Text) = "", "NULL", "'" & txt��ӪƷ��.Text & "'") & "," & _
            IIf(Trim(txt��ע.Text) = "", "NULL", "'" & txt��ע.Text & "'") & _
            ")"
    
    gcnOracle.BeginTrans: blnTran = True
    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '������Ƭ
    For i = 0 To 2
        If myPicCon.mblnIsModify(i) = True Then
            'ֻ�з����˸��Ĳ���Ҫ����
            Call zldatabase.ExecuteProcedure("Zl_��Ӧ����Ƭ_Delete(" & lngID & "," & i & ")", Me.Caption)
            
            If myPicCon.mblnExistPic(i) = True And img��Ƭ(i).Tag <> "" Then
                '����
                If sys.Savelob(100, 23, lngID & "," & i, img��Ƭ(i).Tag) = False Then
                    gcnOracle.RollbackTrans
                    MsgBox "��Ƭ����ʧ�ܡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    gcnOracle.CommitTrans: blnTran = False
    
    Save��λ = True
    Exit Function
errHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function �༭��λ(ByVal FrmMain As Object, ByVal lng�ϼ�id As Long, _
    intEditType As gEditType, Optional strID As String = "", Optional ByVal blnĩ�� As Boolean = False, _
    Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�༭��Ӧ�̵���
    '--�����:frmMain-���õ�������
    '--       lng�ϼ�id-�ϼ�id
    '--       intEditType -�༭����
    '--       strID-�༭�����ĵ�ǰID
    '--       blnĩ��-�Ƿ���δ����Ŀ
    '--������:
    '--��  ��:�༭�ɹ�,����ture,����false
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim intTemp As Byte, i As Integer
    Dim strTemp As String
    Dim strTempFile As String
   

    mintSuccess = 0
    
    mstrID = strID
    mlng�ϼ�id = lng�ϼ�id
    mintEditType = intEditType
    mstrPrivs = strPrivs
    On Error GoTo errHandle
    '��ʼ��Ժ����Ϣ
    gstrSQL = "Select ���, ���� From Zltools.Zlnodelist "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡվ��")
    With cmbStationNo
        .Clear
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!��� & "-" & rsTemp!����
            .ItemData(.NewIndex) = rsTemp!���
            rsTemp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    If mlng�ϼ�id <> 0 Then
        '����ϼ����뼰����
        'by lesfeng 2009-12-2 �����Ż�
        gstrSQL = "Select ����,���� From ��Ӧ�� where id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�ϼ�id)
        If rsTemp.EOF Then
            ShowMsgbox "�ϼ������ѱ�����ɾ��,���������Ӹ÷�����¼���Ŀ!"
            Exit Function
        End If
        txtParent.Text = "[" & Nvl(rsTemp!����, "..") & "]" & Nvl(rsTemp!����, "..")
        txtUpCode.Text = Nvl(rsTemp!����)
        mlng�ϼ�id = lng�ϼ�id
    Else
        txtParent.Text = "��"
        txtUpCode.Text = ""
    End If
    If mintEditType <> g���� Then
        '��ȷ���������������Ŀ
        'by lesfeng 2009-12-2 �����Ż�
        gstrSQL = "Select ID,�ϼ�ID,����,����,����,ĩ��,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,�绰,��������," & _
                  "       �ʺ�,��ϵ��,����ʱ��,����ʱ��,����,������,���ö�,����ί����,����ί������,������֤��,������֤����," & _
                  "       ҩ��ֱ�����,ҩ��ֱ�������,��Ȩ��,��Ȩ��,վ��,��ӪƷ��,��ע " & _
                  "  From ��Ӧ�� where id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���༭�Ĺ�Ӧ��!", Val(strID))
        
        If mintEditType = g�鿴 Then
        Else
            If SetEditPro(Nvl(rsTemp!����)) = False Then Exit Function
        End If
        With rsTemp
            txtCode.Text = Mid(Nvl(!����), Len(txtUpCode.Text) + 1)
            txtCode.MaxLength = Len(txtCode.Text)
            txtCode.Tag = .Fields("����").DefinedSize
            Dim intIndex As Long
            For intIndex = 0 To 11
                strTemp = TxtEdit(intIndex).Tag
                Select Case strTemp
                Case "���ö�"
                        TxtEdit(intIndex).Text = Format(Nvl(.Fields(strTemp), 0), "####0.00;####0.00; ;")
                Case Else
                    TxtEdit(intIndex).Text = Nvl(.Fields(strTemp))
                End Select
            Next
            If IsNull(!���֤Ч��) Then
                Dtp���֤Ч��.Value = ""
            Else
                Dtp���֤Ч��.Value = Format(!���֤Ч��, "yyyy-mm-dd")
            End If
            If IsNull(!ִ��Ч��) Then
                Dtpִ��Ч��.Value = ""
            Else
                Dtpִ��Ч��.Value = Format(!ִ��Ч��, "yyyy-mm-dd")
            End If
            
            TxtEdit(15).Text = Nvl(!��Ȩ��)
            If IsNull(!��Ȩ��) Then
                dtp��Ȩ��.Value = ""
            Else
                dtp��Ȩ��.Value = Format(!��Ȩ��, "yyyy-mm-dd")
            End If
                        
            If IsNull(!����ί������) Then
                Dtpί��������.Value = ""
            Else
                Dtpί��������.Value = Format(!����ί������, "yyyy-mm-dd")
            End If
            
            If IsNull(!������֤����) Then
                dtp������֤.Value = ""
            Else
                dtp������֤.Value = Format(!������֤����, "yyyy-mm-dd")
            End If
            If IsNull(!ҩ��ֱ�������) Then
                dtpҩ��ֱ���.Value = ""
            Else
                dtpҩ��ֱ���.Value = Format(!ҩ��ֱ�������, "yyyy-mm-dd")
            End If
                        
            TxtEdit(12).Text = Nvl(!����ί����)
            TxtEdit(13).Text = Nvl(!������֤��)
            TxtEdit(14).Text = Nvl(!ҩ��ֱ�����)
            
            txt��ӪƷ��.Text = Nvl(!��ӪƷ��)
            txt��ע.Text = Nvl(!��ע)
            
            '����վ����Ϣ
            With cmbStationNo
                For i = 0 To .ListCount - 1
                    If Mid(.List(i), 1, 1) = Nvl(rsTemp!վ��) Then
                        .ListIndex = i
                        Exit For
                    End If
                Next
            End With
            
            If !ĩ�� = 1 Then
                chkĩ��.Value = 1
            Else
                chkĩ��.Value = 0
            End If
            strTemp = Nvl(!����)
            
            '��ȡ����
            If Len(strTemp) >= 4 Then
                For intTemp = 0 To 4
                    If intTemp > Len(strTemp) - 1 Then
                        chkType(intTemp).Value = 0
                    Else
                        chkType(intTemp).Value = Mid(strTemp, intTemp + 1, 1)
                    End If
                Next
            End If
        End With
    Else
        '����
        zlChangeCode "��Ӧ��", mlng�ϼ�id, txtUpCode, txtCode, chkCodeLen, Me.Caption
        If blnĩ�� Then
            chkĩ��.Value = 1
        Else
            chkĩ��.Value = 0
        End If
        For intTemp = 0 To 4
            chkType(intTemp).Value = 0
        Next
    End If
    
    If chkĩ��.Value <> 1 Then
        Set pic����.Container = Me
        pic����.Top = sstab.Top
        chkCodeLen.Top = pic����.Top + pic����.Height + 100
        fra(0).Top = chkCodeLen.Top + chkCodeLen.Height + 100
        cmdCancel.Top = fra(0).Top + fra(0).Height + 100
        cmdOK.Top = cmdCancel.Top
        sstab.Visible = False
        frm��Ӧ�̱༭.Height = cmdCancel.Top + cmdCancel.Height + 600
    Else
        sstab.Visible = True
    End If
    mblnChange = False
    Me.chkCodeLen.Visible = InStr(1, mstrPrivs, "�ı���볤��") <> 0
    If chkĩ��.Value <> 1 Then
        Me.Caption = "����༭"
        Label1.Caption = "�Թ�Ӧ�̷����������.ͬʱ�ɼӳ���������б���ĳ��ȡ�"
    End If
    
    '����ͼƬ
    For i = 0 To 2
        strTempFile = sys.Readlob(100, 23, strID & "," & i)
        img��Ƭ(i).Picture = LoadPicture(strTempFile)
        myPicCon.mblnIsModify(i) = False
        myPicCon.mblnExistPic(i) = (strTempFile <> "")
        lblͼƬ˵��(i) = GetPictureInfo(img��Ƭ(i).Picture)
        'ɾ������ʱ�ļ�
        If lblͼƬ˵��(i) <> "����Ƭ" Then
            Kill strTempFile
        End If
    Next
    
    frm��Ӧ�̱༭.Show 1, FrmMain
    �༭��λ = mintSuccess > 0
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub setCtlEn()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ÿؼ���Enable����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    If mintEditType = g�鿴 Then
        txtCode.Enabled = False
        txtUpCode.Enabled = False
        txtParent.Enabled = False
        cmd�ϼ�.Enabled = False
        cmdOK.Visible = False
        cmbStationNo.Enabled = False
        For intIndex = 0 To 4
            chkType(intIndex).Enabled = False
        Next
        
        For intIndex = 0 To TxtEdit.UBound
            TxtEdit(intIndex).Enabled = False
        Next
        Dtp���֤Ч��.Enabled = False
        Dtpִ��Ч��.Enabled = False
        Dtpί��������.Enabled = False
        dtp������֤.Enabled = False
        dtpҩ��ֱ���.Enabled = False
'        cmbStationNo.Enabled = False
        chkCodeLen.Enabled = False
        dtp��Ȩ��.Enabled = False
        txt��ӪƷ��.Enabled = False
        txt��ע.Enabled = False
        
    End If
    cmdOK.Enabled = mblnChange And Trim(TxtEdit(0).Text) <> "" And Trim(txtCode.Text) <> ""
End Sub

Private Sub cmd�ϼ�_Click()
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    Dim int����  As Integer
    
    gstrSQL = "select ID,�ϼ�ID,����,���� from ��Ӧ��  " & _
        "where ĩ�� <> 1 start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    strID = IIf(mlng�ϼ�id = 0, "", mlng�ϼ�id)
    
    str���� = TxtEdit(0).Text
    str���� = txtUpCode.Text
    blnRe = frm����ѡ��.ShowTree(gstrSQL, strID, str����, str����, mstrID, "��Ӧ��", "���й�Ӧ��")
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        txtParent.Text = str����
        mlng�ϼ�id = Val(strID)
        '���ñ���
        zlChangeCode "��Ӧ��", mlng�ϼ�id, txtUpCode, txtCode, chkCodeLen, Me.Caption
        setCtlEn
    End If
End Sub

Private Sub cmd��Ƭ_Click(Index As Integer)
    Dim intPicIndex As Integer
    
    intPicIndex = sstab.Tab - 3
    
    Select Case Index
        Case 0 '�ļ�
            With cdl��Ƭ
                .CancelError = True
                .Filter = "ͼƬ�ļ�(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
                
                On Error Resume Next
                .ShowOpen
                If Err <> 0 Then
                    'ûѡ���ļ�
                    Err.Clear
                Else
                    img��Ƭ(intPicIndex).Picture = LoadPicture(.FileName)
'                    img��Ƭ.Left = pic����.ScaleLeft
'                    img��Ƭ.Top = pic����.ScaleTop
                    
'                    DoEvents
                    If Err <> 0 Then
                        MsgBox "ͼƬ�ļ���Ч�����ļ������ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    lblͼƬ˵��(intPicIndex) = GetPictureInfo(img��Ƭ(intPicIndex).Picture)
                    img��Ƭ(intPicIndex).Tag = .FileName
                    myPicCon.mblnExistPic(intPicIndex) = True
                    myPicCon.mblnIsModify(intPicIndex) = True
                End If
            End With
        Case 1 '���
            myPicCon.mblnExistPic(intPicIndex) = False
            myPicCon.mblnIsModify(intPicIndex) = True
            Call ��ʾ��ͼƬ(intPicIndex)
    End Select
    
    
End Sub

Private Sub ��ʾ��ͼƬ(ByVal intPicIndex As Integer)
    '��ͼƬ������ʾ��ͼƬ��Ϣ
    If myPicCon.mblnExistPic(intPicIndex) = False Then
        img��Ƭ(intPicIndex).Picture = Nothing
        img��Ƭ(intPicIndex).Tag = ""
        lblͼƬ˵��(intPicIndex) = "����Ƭ"
    End If
End Sub
Private Sub dtp��Ȩ��_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtp��Ȩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Dtpί��������_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtpί��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Dtp���֤Ч��_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtp���֤Ч��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpҩ��ֱ���_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtpҩ��ֱ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Dtpִ��Ч��_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtpִ��Ч��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp������֤_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtp������֤_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
    '��ʼվ��
'    cmbStationNo.Visible = gSystemPara.bln����վ�� And chkĩ��.Value = 1
'    lblStationNo.Visible = cmbStationNo.Visible
    
    If Me.TxtEdit(0).Enabled Then Me.TxtEdit(0).SetFocus
    Call Ȩ�޿���
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFist = True
    Call InitDefaultLen
End Sub



Private Sub Form_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To 2
        img��Ƭ(i).Move pic��Ƭ(i).ScaleLeft, pic��Ƭ(i).ScaleTop, pic��Ƭ(i).ScaleWidth, pic��Ƭ(i).ScaleHeight
    Next
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
    If sstab.Tab >= 3 And sstab.Tab <= 5 Then
        cmd��Ƭ(0).Enabled = True
        cmd��Ƭ(1).Enabled = True
    Else
        cmd��Ƭ(0).Enabled = False
        cmd��Ƭ(1).Enabled = False
    End If
    
    lblͼƬ˵��(0).Visible = (sstab.Tab = 3)
    lblͼƬ˵��(1).Visible = (sstab.Tab = 4)
    lblͼƬ˵��(2).Visible = (sstab.Tab = 5)
End Sub

Private Sub TxtCode_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCode, KeyAscii, m����ʽ
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 0 Then
        TxtEdit(1).Text = zlCommFun.SpellCode(TxtEdit(0).Text)
    End If
    setCtlEn
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Dim blnOpen As Boolean
    
    Select Case TxtEdit(Index).Tag
    Case "������", "���ö�", "����"
            blnOpen = False
    Case Else
            blnOpen = True
    End Select
    SetTxtGotFocus TxtEdit(Index), blnOpen
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case TxtEdit(Index).Tag
        Case "����"
            If cmbStationNo.Visible And cmbStationNo.Enabled Then
                cmbStationNo.SetFocus
            ElseIf chkType(0).Enabled And chkType(0).Visible Then
                chkType(0).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case TxtEdit(Index).Tag
    Case "������"
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m����ʽ
    Case "���ö�"
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m���ʽ
    Case "�ʺ�"
        If LenB(StrConv(TxtEdit(Index).Text, vbFromUnicode)) >= 50 And (KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    Case Else
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m�ı�ʽ
    End Select
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If TxtEdit(Index).Tag = "���ö�" Then
        TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "####0.00;-####0.00; ;")
    End If
    ImeLanguage False
End Sub

Private Sub txtParent_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
End Sub

Private Sub Ȩ�޿���()
    'Ȩ�޿���
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    
    blnҩƷ = InStr(1, mstrPrivs, "ҩƷ��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���ʹ�Ӧ��") <> 0
    bln�豸 = InStr(1, mstrPrivs, "�豸��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "������Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���Ĺ�Ӧ��") <> 0
    
    chkType(0).Enabled = blnҩƷ And mintEditType <> g�鿴
    chkType(1).Enabled = bln���� And mintEditType <> g�鿴
    chkType(2).Enabled = bln�豸 And mintEditType <> g�鿴
    chkType(3).Enabled = bln���� And mintEditType <> g�鿴
    chkType(4).Enabled = bln���� And mintEditType <> g�鿴
End Sub
Private Function SetEditPro(ByVal str���� As String) As Boolean
    '���ñ༭Ȩ��
    
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    
    blnҩƷ = InStr(1, mstrPrivs, "ҩƷ��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���ʹ�Ӧ��") <> 0
    bln�豸 = InStr(1, mstrPrivs, "�豸��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "������Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���Ĺ�Ӧ��") <> 0
    
    Err = 0: On Error GoTo ErrHand:
    SetEditPro = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

