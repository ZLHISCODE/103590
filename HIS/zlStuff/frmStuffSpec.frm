VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������Ϲ��༭"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmStuffSpec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "���������Ʒ��(&A)"
      Height          =   350
      Left            =   2280
      TabIndex        =   59
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "������������(&B)"
      Height          =   350
      Left            =   4275
      TabIndex        =   58
      Top             =   7680
      Width           =   1695
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3720
      Left            =   1200
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   8760
      Visible         =   0   'False
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   6562
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   109
      Top             =   7560
      Width           =   8880
   End
   Begin VB.PictureBox picFound 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2835
      ScaleHeight     =   210
      ScaleWidth      =   5505
      TabIndex        =   104
      Top             =   885
      Width           =   5505
      Begin VB.Label lblFound 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ע���ù������2002��12��20�գ���2003��8��10��ͣ�á�"
         Height          =   225
         Left            =   105
         TabIndex        =   65
         Top             =   0
         Width           =   4860
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf���� 
      Height          =   1845
      Left            =   360
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   8640
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
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   107
      Top             =   570
      Width           =   8775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�����˳�(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   56
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7785
      TabIndex        =   57
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      Picture         =   "frmStuffSpec.frx":030A
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1100
   End
   Begin TabDlg.SSTab stbSpec 
      Height          =   6705
      Left            =   120
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   720
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   11827
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ(&1)"
      TabPicture(0)   =   "frmStuffSpec.frx":0454
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "�۸���Ϣ(&2)"
      TabPicture(1)   =   "frmStuffSpec.frx":0470
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk����ʹ��"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmd����"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra��������"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chk���ηѱ�"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cbo�������"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cbo��������"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cbo������Ŀ"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fra(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt������Ŀ"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl������Ŀ"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl(20)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl(18)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl(19)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CheckBox chk����ʹ�� 
         Caption         =   "����ʹ��(&G)"
         Height          =   285
         Left            =   -68280
         TabIndex        =   125
         Top             =   2400
         Width           =   1290
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Left            =   -67200
         TabIndex        =   122
         TabStop         =   0   'False
         Tag             =   "����"
         ToolTipText     =   "��*��ѡ����"
         Top             =   1950
         Width           =   255
      End
      Begin VB.Frame fra�������� 
         Caption         =   "��������"
         Height          =   1230
         Left            =   -70200
         TabIndex        =   98
         Top             =   2880
         Width           =   3780
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   17
            Left            =   2820
            MaxLength       =   5
            TabIndex        =   102
            Tag             =   "������"
            Top             =   375
            Width           =   630
         End
         Begin VB.CheckBox chk�ⷿ 
            Caption         =   "���Ŀⷿ����(&W)"
            Height          =   420
            Left            =   105
            TabIndex        =   99
            Top             =   315
            Width           =   1665
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "���ϲ��ŷ���(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   105
            TabIndex        =   100
            Top             =   750
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ч��(&7)       ��"
            Height          =   180
            Left            =   2190
            TabIndex        =   101
            Top             =   435
            Width           =   1440
         End
      End
      Begin VB.Frame fra 
         Caption         =   "��������"
         Height          =   2370
         Index           =   1
         Left            =   5370
         TabIndex        =   35
         Top             =   456
         Width           =   3195
         Begin VB.ComboBox cbo�洢���� 
            Height          =   300
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Tag             =   "������Դ"
            Top             =   1965
            Width           =   2970
         End
         Begin VB.ComboBox cbo���ʷ��� 
            Height          =   300
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "������Դ"
            Top             =   1290
            Width           =   2970
         End
         Begin VB.ComboBox cbo������Դ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Tag             =   "������Դ"
            Top             =   690
            Width           =   1950
         End
         Begin VB.ComboBox cbo��Դ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Tag             =   "��Դ���"
            Top             =   330
            Width           =   1950
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�洢����(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   25
            Left            =   150
            TabIndex        =   42
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ʷ���(&J)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   150
            TabIndex        =   40
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Դ����(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   150
            TabIndex        =   38
            Top             =   750
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Դ���(&Q)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   150
            TabIndex        =   36
            Top             =   390
            Width           =   990
         End
      End
      Begin VB.Frame Fra2 
         Caption         =   "��������"
         Height          =   3640
         Left            =   5355
         TabIndex        =   44
         Top             =   2925
         Width           =   3195
         Begin VB.CheckBox chkֲ��Ĳ� 
            Caption         =   "ֲ���ԺĲ�"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   1260
         End
         Begin VB.CheckBox chkInstrument 
            Caption         =   "��е�����ĵ���"
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   285
            Width           =   1575
         End
         Begin VB.CheckBox chkCode 
            Caption         =   "�������(&7)"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1507
            Width           =   1365
         End
         Begin VB.CheckBox chkCostly 
            Caption         =   "��ֵ����(&6)"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1214
            Width           =   1485
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   3030
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txt��ѡ�� 
            Height          =   300
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   54
            Top             =   2565
            Width           =   2085
         End
         Begin VB.CheckBox chkԭ�� 
            Caption         =   "ԭ��(&3)"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   1500
         End
         Begin VB.CheckBox chk�޾��Բ��� 
            Caption         =   "�޾�����(&4)"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   603
            Width           =   1500
         End
         Begin VB.CheckBox Chkһ���Բ��� 
            Caption         =   "һ��������(&5)"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   922
            Width           =   1485
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   7
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   53
            Tag             =   "���Ч��"
            Top             =   2175
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "Ժ��"
            Height          =   180
            Left            =   120
            TabIndex        =   63
            Top             =   3105
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѡ��(&V)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1920
            TabIndex        =   61
            Top             =   2235
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ч��(&7)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   52
            Top             =   2235
            Width           =   990
         End
      End
      Begin VB.CheckBox chk���ηѱ� 
         Caption         =   "���ηѱ�(&M)"
         Height          =   285
         Left            =   -70080
         TabIndex        =   97
         Top             =   2400
         Width           =   1290
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Tag             =   "Ӧ�ö���"
         Top             =   1500
         Width           =   2115
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Tag             =   "ҽ������"
         Top             =   1125
         Width           =   2115
      End
      Begin VB.ComboBox cbo������Ŀ 
         Height          =   300
         Left            =   -69045
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Tag             =   "������Ŀ"
         Top             =   750
         Width           =   2115
      End
      Begin VB.Frame fra 
         Height          =   6120
         Index           =   0
         Left            =   90
         TabIndex        =   108
         Top             =   450
         Width           =   5190
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   285
            Left            =   4750
            TabIndex        =   124
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   2978
            Width           =   285
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   22
            Left            =   1125
            MaxLength       =   250
            TabIndex        =   34
            Tag             =   "˵��"
            Top             =   5700
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   20
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "��ʼ���"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   19
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "ƴ������"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   18
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "��Ʒ��"
            Top             =   1050
            Width           =   3945
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "�������(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   3795
            TabIndex        =   14
            Top             =   2205
            Width           =   1335
         End
         Begin VB.TextBox txtע��֤�� 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   26
            Top             =   4120
            Width           =   3915
         End
         Begin VB.CheckBox chk���ٲ��� 
            Caption         =   "���ٲ���(&S)"
            Height          =   210
            Left            =   3795
            TabIndex        =   33
            Top             =   5367
            Width           =   1290
         End
         Begin MSComCtl2.DTPicker dtp���֤Ч�� 
            Height          =   345
            Left            =   1335
            TabIndex        =   32
            Top             =   5300
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   137363457
            CurrentDate     =   39227
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   30
            Tag             =   "���֤��"
            Top             =   4930
            Width           =   3915
         End
         Begin VB.TextBox txt��׼�ĺ� 
            Height          =   300
            Left            =   1125
            MaxLength       =   40
            TabIndex        =   22
            Top             =   3350
            Width           =   3915
         End
         Begin VB.TextBox txtע���̱� 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   24
            Top             =   3730
            Width           =   3915
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   4575
            MaxLength       =   1
            TabIndex        =   18
            Tag             =   "��ʶ����"
            Top             =   2595
            Width           =   465
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "��ʶ����"
            Top             =   2595
            Width           =   1605
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��������(&I)"
            Height          =   210
            Left            =   2430
            TabIndex        =   13
            Top             =   2205
            Width           =   1335
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "����ϵ��"
            Text            =   "1"
            Top             =   2160
            Width           =   870
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   1125
            MaxLength       =   60
            TabIndex        =   20
            Tag             =   "������"
            Top             =   2970
            Width           =   3615
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   3
            Tag             =   "���"
            Top             =   660
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "������"
            Top             =   300
            Width           =   2160
         End
         Begin VB.ComboBox cbo��λ 
            Height          =   300
            Index           =   0
            ItemData        =   "frmStuffSpec.frx":048C
            Left            =   1125
            List            =   "frmStuffSpec.frx":048E
            TabIndex        =   8
            Tag             =   "ɢװ��λ"
            Text            =   "֧"
            Top             =   1770
            Width           =   1245
         End
         Begin VB.ComboBox cbo��λ 
            Height          =   300
            Index           =   1
            ItemData        =   "frmStuffSpec.frx":0490
            Left            =   3810
            List            =   "frmStuffSpec.frx":0492
            TabIndex        =   10
            Tag             =   "��װ��λ"
            Text            =   "֧"
            Top             =   1770
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker dtpע��֤��Ч�� 
            Height          =   345
            Left            =   1335
            TabIndex        =   28
            Top             =   4478
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   137297921
            CurrentDate     =   39227
         End
         Begin VB.Label lblע��֤ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ע��֤��Ч��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   4560
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "˵��(&S)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   30
            Left            =   480
            TabIndex        =   121
            Top             =   5760
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(���)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   4530
            TabIndex        =   115
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(ƴ��)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   2490
            TabIndex        =   114
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ʒ������(&P)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   113
            Top             =   1500
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʒ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   112
            Top             =   1110
            Width           =   720
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ע��֤��(&T)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   4180
            Width           =   990
         End
         Begin VB.Label lblIn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���֤Ч��(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   5382
            Width           =   1170
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���֤��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   4990
            Width           =   720
         End
         Begin VB.Label lbl��׼�ĺ� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��׼�ĺ�(&W)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   3410
            Width           =   990
         End
         Begin VB.Label lblע���̱� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ע���̱�(&E)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   3790
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��ʶ����(&D)"
            Height          =   180
            Index           =   16
            Left            =   3570
            TabIndex        =   17
            Top             =   2655
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��ʶ����(&Z)"
            Height          =   180
            Index           =   17
            Left            =   120
            TabIndex        =   15
            Top             =   2685
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ϵ��(&X)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   2220
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������(&M)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   300
            TabIndex        =   19
            Tag             =   "������"
            Top             =   3030
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���(&G)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   480
            TabIndex        =   2
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������(&N)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ɢװ��λ(&U)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   1830
            Width           =   990
         End
         Begin VB.Label lblסԺ��λ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��װ��λ(&K)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2805
            TabIndex        =   9
            Top             =   1830
            Width           =   990
         End
      End
      Begin VB.Frame fra 
         Height          =   4515
         Index           =   2
         Left            =   -74850
         TabIndex        =   64
         Top             =   465
         Width           =   4365
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   21
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   89
            Tag             =   "��ֵ˰��"
            Top             =   3840
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   11
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   82
            Tag             =   "ָ���ۼ�"
            Top             =   2628
            Width           =   3030
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   73
            Tag             =   "ָ������"
            Top             =   1419
            Width           =   1455
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   12
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   84
            Tag             =   "ָ�������"
            Text            =   "13.0435"
            Top             =   4545
            Visible         =   0   'False
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   16
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   71
            Tag             =   "��ǰ�ۼ�"
            Top             =   1016
            Width           =   3030
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   76
            Tag             =   "�ɹ�����"
            Text            =   "100"
            Top             =   1822
            Width           =   1455
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   10
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   79
            Tag             =   "�����"
            Top             =   2225
            Width           =   1455
         End
         Begin VB.ComboBox cbo�۸����� 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Tag             =   "�۸�����"
            Top             =   210
            Width           =   3090
         End
         Begin VB.TextBox txtEdit 
            Enabled         =   0   'False
            Height          =   300
            Index           =   14
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   88
            Tag             =   "�������"
            Text            =   "100"
            Top             =   3434
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   13
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   86
            Tag             =   "�ӳ���"
            Text            =   "15.00"
            Top             =   3031
            Width           =   2790
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   15
            Left            =   1170
            MaxLength       =   14
            TabIndex        =   69
            Tag             =   "�ɱ��۸�"
            Top             =   613
            Width           =   3030
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   4080
            TabIndex        =   120
            Top             =   3900
            Width           =   90
         End
         Begin VB.Label lblPercent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   4080
            TabIndex        =   119
            Top             =   3494
            Width           =   90
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
            Left            =   4080
            TabIndex        =   118
            Top             =   3091
            Width           =   90
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
            Left            =   4080
            TabIndex        =   117
            Top             =   4605
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ˰��(&Z)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   90
            TabIndex        =   116
            Top             =   3900
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ָ���ۼ�(&K)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   90
            TabIndex        =   81
            Top             =   2688
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ָ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   90
            TabIndex        =   72
            Top             =   1479
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ָ������(&E)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   90
            TabIndex        =   83
            Top             =   4605
            Visible         =   0   'False
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
            Left            =   2700
            TabIndex        =   77
            Top             =   1882
            Width           =   90
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ɹ�����(&X)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   90
            TabIndex        =   75
            Top             =   1882
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����(&T)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   78
            Top             =   2285
            Width           =   810
         End
         Begin VB.Label lbl���۵�λ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ԫ/Ƭ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2700
            TabIndex        =   74
            Top             =   1479
            Width           =   735
         End
         Begin VB.Label lbl���۵�λ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ԫ/Ƭ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2700
            TabIndex        =   80
            Top             =   2285
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�۸�����(&P)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   90
            TabIndex        =   66
            Top             =   270
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   90
            TabIndex        =   87
            Top             =   3494
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ�ۼ�(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   90
            TabIndex        =   70
            Top             =   1076
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ӳ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   90
            TabIndex        =   85
            Top             =   3091
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ɱ��۸�(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   90
            TabIndex        =   68
            Top             =   673
            Width           =   990
         End
      End
      Begin VB.TextBox txt������Ŀ 
         Height          =   300
         Left            =   -69045
         MaxLength       =   40
         TabIndex        =   96
         ToolTipText     =   "��*��ѡ����"
         Top             =   1920
         Width           =   2115
      End
      Begin VB.Label lbl������Ŀ 
         Caption         =   "������Ŀ(&F)"
         Height          =   255
         Left            =   -70125
         TabIndex        =   123
         Top             =   1943
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   20
         Left            =   -70125
         TabIndex        =   90
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   18
         Left            =   -70125
         TabIndex        =   92
         Top             =   1185
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ�ö���(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   -70125
         TabIndex        =   94
         Top             =   1575
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8400
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":0494
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":0A2E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":0FC8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSpec.frx":1562
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblƷ��˵�� 
      Caption         =   "����:0201    Ʒ����һ�������         Ӣ������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   750
      TabIndex        =   111
      Top             =   240
      Width           =   7965
   End
   Begin VB.Label lbl 
      Caption         =   $"frmStuffSpec.frx":1AFC
      Height          =   390
      Index           =   0
      Left            =   -150
      TabIndex        =   103
      Top             =   8640
      Width           =   7125
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmStuffSpec.frx":1B83
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "frmStuffSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng����ID As Long
Dim mstr����ID As String         '��ǰ�༭�Ĳ���ID
Private mlng����id As Long       '��ǰѡ��ķ���id

Dim mintSuccess As Integer
Dim mintEditType As gEditType    '�༭����
Dim mblnChange As Boolean
Dim mstrPrivs As String         'Ȩ�޴�
Dim mblnFrist As Boolean        '��һ������ϵͳʱ
Dim mintCount As Integer
Dim mstr���� As String
Dim mintUnit As Integer     '0-ɢװ��λ,1-��װ��λ
Dim mintCodeLength As Integer   '����ĳ���,�����ݿ��ж�ȡ�����ĳ���
Private Const mlngModule = 1711
Private mblnLoad As Boolean      '����ֻactiveһ��
Private mintSet���� As Integer  '���÷�������
Private mblnInStrument As Boolean '�Ƿ���װ������ϵͳ
Private mstrע��֤�� As String   '��¼�޸�֮ǰ��ע��֤��
Private mstrע��֤��Ч�� As String  '��¼�޸�֮ǰ��ע��֤��Ч��
Private mintע���޸Ĳ��� As Integer '0-ֻ�޸ĵ�ǰ���1-ͬ���޸�Ʒ��������ע��֤�ź�ע��֤��Ч��

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Public Function ShowEditCard(ByVal frmMain As Object, intEditType As gEditType, ByVal lng����ID As Long, ByVal lng����id As Long, _
    Optional str����ID As String = "", Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�༭��������
    '--�����:frmMain-���õ�������
    '--       intEditType -�༭����
    '--       lng����ID-����ID(Ʒ��ID)
    '--       str����ID-�༭�����ĵ�ǰID
    '--       strPrivs-Ȩ�޴�
    '--������:
    '--��  ��:�༭�ɹ�,����ture,����false
    '--����:���˺�
    '--����:2007/05/25
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    mintEditType = intEditType: mstrPrivs = strPrivs: mstr����ID = str����ID: mlng����ID = lng����ID: mlng����id = lng����id
    mintSuccess = 0
    
    frmStuffSpec.Show 1, frmMain
    
    ShowEditCard = mintSuccess > 0
End Function

Private Function GetDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������������
    '--�����:
    '--������:
    '--��  ��:���ڷ���true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    gstrSQL = "Select ����||'-'||���� From ������Դ���� Order By ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption

    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgBox "δ���ò�����Դ���ࣨ�ֵ������"
            Exit Function
        End If
        Me.cbo������Դ.Clear
        Do While Not .EOF
            Me.cbo������Դ.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    If Me.cbo������Դ.ListCount > 0 Then Me.cbo������Դ.ListIndex = 0
    
     
    gstrSQL = "Select ����||'-'||���� From �������� where ����=1 Order By ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     
    With rsTemp
        Me.cbo��������.Clear
        If .RecordCount = 0 Then
            ShowMsgBox "δ�����������ĵ�ҽ�����ͣ��ֵ������"
            Exit Function
        End If
        Do While Not .EOF
            Me.cbo��������.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    
    '���˺�:2007/05/25:���Ӳ��ʷ���
    gstrSQL = "Select ����||'-'||���� as ����,���� From ���ϲ��ʷ���  order by ���� "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo���ʷ���.Clear
        Do While Not .EOF
            Me.cbo���ʷ���.AddItem zlStr.nvl(!����)
            .MoveNext
        Loop
        If cbo���ʷ���.ListCount <> 0 Then cbo���ʷ���.ListIndex = 0
    End With
    
    '���˺�:2007/05/25:���ϴ洢����
    gstrSQL = "Select ����||'-'||���� as ����,���� From ���ϴ洢���� order by ���� "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo�洢����.Clear
        Do While Not .EOF
            Me.cbo�洢����.AddItem zlStr.nvl(!����)
            .MoveNext
        Loop
        If cbo�洢����.ListCount <> 0 Then cbo�洢����.ListIndex = 0
    End With
    
    gstrSQL = "Select ����||'-'||���� as ����,���� From ���ϻ�Դ���  order by ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo��Դ.Clear
        Do While Not .EOF
            Me.cbo��Դ.AddItem zlStr.nvl(!����)
            .MoveNext
        Loop
        cbo��Դ.ListIndex = 0
    End With
    
    If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
    
    gstrSQL = "" & _
        "   Select ID,'['||����||']'||���� as ����" & _
        "   From ������Ŀ" & _
        "   where ĩ��=1 and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
        "   Order By ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgBox "δ������ϸ��������Ŀ��"
            Exit Function
        End If
        Me.cbo������Ŀ.Clear
        Do While Not .EOF
            Me.cbo������Ŀ.AddItem !����: Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.NewIndex) = !Id
            .MoveNext
        Loop
        If Me.cbo������Ŀ.ListCount > 0 Then Me.cbo������Ŀ.ListIndex = 0
    End With
    
    mintUnit = Get���۵�λ
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    
   'mstrFormat = GetFmtString(mintUnit) 'IIf(mintUnit = 1, "#####0.0000;-#####0.0000; ;", "#####0.0000000;-#####0.0000000; ;")
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbo������Դ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkֲ��Ĳ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo���ʷ���_Change()
    mblnChange = True
    
End Sub

Private Sub cbo���ʷ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cbo�洢����_Change()
    mblnChange = True
End Sub

Private Sub cbo�洢����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub

Private Sub cbo��λ_Click(Index As Integer)
    Call cbo��λ_Change(Index)
End Sub

Private Sub cbo��λ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub cbo�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo��Դ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub cbo�۸�����_Click()
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If txtEdit(i).Tag = "�������" Then
            txtEdit(14).Enabled = InStr(1, mstrPrivs, ";ָ���۸����;") <> 0 And Not (cbo�۸�����.Text = "����")
            If txtEdit(14).Enabled Then
                txtEdit(14).BackColor = &H80000005
            Else
                txtEdit(14).BackColor = &H8000000F
            End If
            Exit For
        End If
    Next
End Sub


Private Sub cbo�۸�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub


Private Sub cbo������Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub



Private Sub chkCostly_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandle
    If chkCostly.Value = 0 Then
        strSql = "select count(*) rec from ҩƷ�շ���¼ a, �շ���¼������Ϣ b where a.ҩƷid=[1] and a.id=b.�շ�id "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr����ID)
        If rsTmp!rec > 0 Then
            rsTmp.Close
            If MsgBox("ȡ������ֵ���ϡ����Խ�ʹ�������⹺��⡱�в�����ʾ��¼�롰��ֵ���ϡ���Ϣ����ȷ����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                chkCostly.Value = 1
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk����_Click()
    If mintEditType = g�鿴 Then Exit Sub
    If chk����.Enabled = False Then Exit Sub
    
    If chk����.Value = 1 Then
        chk�������.Enabled = True
    Else
        chk�������.Enabled = False
        chk�������.Value = 0
    End If
End Sub

Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
    
End Sub

Private Sub chk���ٲ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

 

Private Sub chk�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk���ηѱ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk�޾��Բ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk�ⷿ_Click()
    Dim blnEnable As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '�ڿⷿ������ǰ���£�������ϲ���û�п�棬����������Ƿ����
    
    '    gstrSQL = "" & _
    '            "   Select nvl(Count(*),0) " & _
    '            "   From ҩƷ��� A,��������˵�� B" & _
    '            "  Where A.ҩƷID=[1]" & _
    '            "       And A.�ⷿID=B.����ID And (B.�������� Like '���ϲ���' Or B.�������� Like '%�Ƽ���' )"
    '
    '    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
    '
    '    With rsTemp
    '        blnEnable = True
    '        If .Fields(0).Value <> 0 Then
    '            blnEnable = False
    '        End If
    '    End With
    If Me.chk�ⷿ.Value = 0 Then
        Me.chk����.Value = 0: Me.chk����.Enabled = False
        'Me.chkЧ��.Value = 0: Me.chkЧ��.Enabled = False
        Me.txtEdit(GetTxtIdx("������")).Text = "": Me.txtEdit(GetTxtIdx("������")).Enabled = False
    Else
        Me.chk����.Enabled = True
        Me.txtEdit(GetTxtIdx("������")).Enabled = True
    End If
    SetCtlBackColor txtEdit(GetTxtIdx("������"))
End Sub

Private Function GetTxtIdx(ByVal strName As String) As Integer
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�ı��������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If Trim(txtEdit(i).Tag) = strName Then
            GetTxtIdx = i
            Exit Function
        End If
    Next
    GetTxtIdx = -1
End Function

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:�Ϸ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    Dim strName As String
    Dim bln��ǿ�ƿ���ָ���۸� As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    bln��ǿ�ƿ���ָ���۸� = ISCHECK��ǿ�ƿ���ָ���۸�
    
    ISValied = False
    
    '���ø������������Ƿ������޸�
    If mintEditType = g�޸� Then
        '����������->�������ã��������/סԺ���ü�¼���������޸�
        If Me.chk����.Value = 1 And chk����.Tag = 0 Then
            gstrSQL = "Select 1 " & _
                " From (Select 1 From ������ü�¼ Where �շ�ϸĿid = [1] " & _
                " Union All " & _
                " Select 1 From סԺ���ü�¼ Where �շ�ϸĿid = [1]) " & _
                " Where Rownum < 2 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
            
            If Not rsTemp.EOF Then
                MsgBox "�ù���Ѿ����������ü�¼�����Բ����޸ġ��������á����ԣ�" & vbCrLf & _
                "��ȡ����ѡ���ٱ��棡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '��������->����������  ���ҩƷ�շ���¼���������޸�
        If Me.chk����.Value = 0 And chk����.Tag = 1 Then
            If Not Me.cbo�۸�����.Enabled Then
                MsgBox "�ù���Ѿ��������շ���¼�����Բ����޸ġ��������á����ԣ�" & vbCrLf & _
                "�빴ѡ���ٱ��棡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    For i = 0 To txtEdit.UBound
        strName = txtEdit(i).Tag
        strTmp = Trim(txtEdit(i).Text)
        Select Case strName
        Case "������", "���", "����ϵ��", "����"
            If strTmp = "" Then
                ShowMsgBox strName & "δ���룬������" & strName & "��"
                Me.stbSpec.Tab = 0
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case "�ɹ�����"  ',"ָ�������""�ɱ��۸�",
                '���˺�:��Ҫ�ǽ���ɱ��۸����Ϊ������,���磺����.����ѵ�
                '����:9569 2006-11-20
                If Val(strTmp) = 0 And txtEdit(i).Enabled Then
                    ShowMsgBox strName & "Ϊ0��δ���룬������" & strName & "��"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
        Case "ָ�������"
            '���˺�:ȡ��ָ������ʵ�����,��ʾ�Ƿ��������
            '��ݸ����ҽԺ������ҩƷ�����ĵ�ָ�����ʺͼӳ�������Ϊ0.ҽԺ�Բ���ҩƷ��������ҽԺʵ�гɱ�������,Ŀǰ����ֱ����Ŀ¼�ｫ�ӳ�������Ϊ0,���ǿ���������ʱ���޸�Ϊ0.
            If strTmp = "" And txtEdit(i).Enabled Then
                If MsgBox(strName & "δ���룬�ҽ��Զ�����Ϊ0��" & vbCrLf & "�Ƿ�������棿", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
        Case "ָ������", "ָ���ۼ�"
            If bln��ǿ�ƿ���ָ���۸� = False Then
                If Val(strTmp) = 0 And txtEdit(i).Enabled Then
                    ShowMsgBox strName & "Ϊ0��δ���룬������" & strName & "��"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
'        Case "ƴ������"
'            Me.txtEdit(i).Text = zlStr.GetCodeByORCL(Me.txtEdit(GetTxtIdx("��Ʒ��")).Text, 0)
'        Case "��ʼ���"
'            Me.txtEdit(i).Text = zlStr.GetCodeByORCL(Me.txtEdit(GetTxtIdx("��Ʒ��")).Text, 1)
        Case Else
            
        End Select
        
        If txtEdit(i).MaxLength <> 0 Then
            If LenB(StrConv(strTmp, vbFromUnicode)) > txtEdit(i).MaxLength Then
                ShowMsgBox strName & "����,���" & txtEdit(i).MaxLength & "���ַ�(" & txtEdit(i).MaxLength / 2 & "������)��"
                Me.stbSpec.Tab = 0
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        Select Case strName
        Case "����ϵ��", "ָ������", "ָ���ۼ�", "�ɱ��۸�"
            If Val(strTmp) > 1000000 Then
                ShowMsgBox strName & "�������ֵ1000000��"
                Me.stbSpec.Tab = IIf(strName = "����ϵ��", 0, 1)
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
            If strName = "����ϵ��" And Val(strTmp) <= 0 Then
                ShowMsgBox strName & "��������㣡"
                Me.stbSpec.Tab = IIf(strName = "����ϵ��", 0, 1)
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
            
        Case "ָ�������", "�������", "�ɹ�����", "��ֵ˰��"
            If Val(strTmp) > 100 Then
                ShowMsgBox strName & "���ܳ���100��"
                Me.stbSpec.Tab = 1
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case "��ǰ�ۼ�"
            If Me.cbo�۸�����.ItemData(cbo�۸�����.ListIndex) = 0 Then
                If Abs(Val(strTmp)) > 1000000 Then
                    ShowMsgBox "��ǰ�ۼ۳������Χ-1000000~1000000��"
                    Me.stbSpec.Tab = 1
                    If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                If bln��ǿ�ƿ���ָ���۸� = False Then
                    If Val(strTmp) > Val(Me.txtEdit(GetTxtIdx("ָ���ۼ�"))) Then
                        ShowMsgBox "�ۼ۲��ܸ���ָ�����ۼۣ�"
                        Me.stbSpec.Tab = 1
                        If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                        Exit Function
                    End If
                End If
            End If
        Case Else
        End Select
    Next
    
    '��������������Ҫ��������
    If chkCode.Value = 1 Then
        If chk�ⷿ.Value = 0 Or chk����.Value = 0 Then
            Me.stbSpec.Tab = 1
            MsgBox "�����������������������Ϊ��������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If chk����.Value = 0 And chk�������.Value = 1 Then
        ShowMsgBox "�Ǹ��ٲ��ϲ������ú������,����!:"
         Me.stbSpec.Tab = 1
         If chk�������.Enabled = True Then chk�������.SetFocus
        Exit Function
    End If
    
    If LenB(StrConv(Me.txtע���̱�.Text, vbFromUnicode)) > 50 Then
        MsgBox "ע���̱곬�������50���ַ���25�����֣�", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txtע���̱�.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txt��׼�ĺ�.Text, vbFromUnicode)) > 40 Then
        MsgBox "��׼�ĺų��������40���ַ���20�����֣�", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt��׼�ĺ�.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txtע��֤��.Text, vbFromUnicode)) > 50 Then
        MsgBox "ע��֤�ų��������50���ַ���25�����֣�", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txtע��֤��.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Me.txt��ѡ��.Text, vbFromUnicode)) > 20 Then
        MsgBox "��ѡ�볬�������20���ַ���10�����֣�", vbInformation, gstrSysName
         Me.stbSpec.Tab = 1
        txt��ѡ��.SetFocus
        Exit Function
    End If
    If Trim(Me.cbo��λ(0).Text) = "" Then ShowMsgBox "������ɢװ��λ��": Me.stbSpec.Tab = 0: Me.cbo��λ(0).SetFocus: Exit Function
    If LenB(StrConv(Me.cbo��λ(0).Text, vbFromUnicode)) > 6 Then ShowMsgBox "ɢװ��λ����(���6���ַ���3������)��": Me.stbSpec.Tab = 0: Me.cbo��λ(0).SetFocus: Exit Function
    If Trim(Me.cbo��λ(1).Text) = "" Then ShowMsgBox "�������װ��λ��": Me.stbSpec.Tab = 0: Me.cbo��λ(1).SetFocus: Exit Function
    If LenB(StrConv(Me.cbo��λ(1).Text, vbFromUnicode)) > 6 Then ShowMsgBox "��װ��λ����(���6���ַ���3������)��": Me.stbSpec.Tab = 0: Me.cbo��λ(1).SetFocus: Exit Function
    ISValied = True

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���濨Ƭ����
    '--�����:
    '--������:
    '--��  ��:����ɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl��ǰ�ۼ� As Double, dblָ���ۼ� As Double, dbl�ɱ��۸� As Double, dblָ������ As Double
    Dim strվ�� As String
    Dim lng����ID As Long
    Dim lng����id As Long
    Dim str��Դ As String
    Dim str��Դ As String
    Dim strValues As String
    
    str��Դ = Trim(cbo��Դ.Text)
    If str��Դ <> "" Then
        str��Դ = Mid(str��Դ, InStr(1, str��Դ, "-") + 1)
    End If
    
    str��Դ = Trim(cbo������Դ.Text)
    If str��Դ <> "" Then
        str��Դ = Mid(str��Դ, InStr(1, str��Դ, "-") + 1)
    End If
    
    err = 0
    On Error GoTo ErrHand:
    
    '------------------------------------------
    '���ݱ���
    If mintUnit <> 0 Then
        dblָ���ۼ� = Round(Val(txtEdit(11).Text) / Val(txtEdit(GetTxtIdx("����ϵ��")).Text), g_С��λ��.obj_���С��.���ۼ�С��)
        dbl��ǰ�ۼ� = Round(Val(txtEdit(16).Text) / Val(txtEdit(GetTxtIdx("����ϵ��")).Text), g_С��λ��.obj_���С��.���ۼ�С��)
        dbl�ɱ��۸� = Round(Val(txtEdit(15).Text) / Val(txtEdit(GetTxtIdx("����ϵ��")).Text), g_С��λ��.obj_���С��.�ɱ���С��)
        dblָ������ = Round(Val(txtEdit(8).Text) / Val(txtEdit(GetTxtIdx("����ϵ��")).Text), g_С��λ��.obj_���С��.�ɱ���С��)
    Else
        dbl��ǰ�ۼ� = Round(Val(txtEdit(16).Text), g_С��λ��.obj_���С��.���ۼ�С��)
        dblָ���ۼ� = Round(Val(txtEdit(11).Text), g_С��λ��.obj_���С��.���ۼ�С��)
        dbl�ɱ��۸� = Round(Val(txtEdit(15).Text), g_С��λ��.obj_���С��.�ɱ���С��)
        dblָ������ = Round(Val(txtEdit(8).Text), g_С��λ��.obj_���С��.�ɱ���С��)
    End If
    If mintEditType = g���� Then
        lng����ID = sys.NextId("�շ���ĿĿ¼")
        gstrSQL = "zl_��������_Insert("
    Else
        lng����ID = Val(mstr����ID)
        gstrSQL = "zl_��������_UPdate("
    End If
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '�洢������ز���
    ' zl_��������_Insert  or zl_��������_UPdate�Ĳ���
    '  ����id_In       In ��������.����id%Type,
    gstrSQL = gstrSQL & mlng����ID & ","
    '  ����id_In       In ��������.����id%Type,
    gstrSQL = gstrSQL & lng����ID & ","
    '  ����_In         In �շ���ĿĿ¼.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(GetTxtIdx("������")).Text) & "',"
    '  ���_In         In �շ���ĿĿ¼.���%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(GetTxtIdx("���")).Text) & "',"
    '  ����_In         In �շ���ĿĿ¼.����%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("������")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  ��ʶ����_In     In �շ���ĿĿ¼.��ʶ����%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("��ʶ����")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  ��ʶ����_In     In �շ���ĿĿ¼.��ʶ����%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("��ʶ����")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  ��ѡ��_In       In �շ���ĿĿ¼.��ѡ��%Type := Null,
    strValues = Trim(txt��ѡ��.Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  ������Դ_In     In ��������.������Դ%Type := Null,
    gstrSQL = gstrSQL & IIf(str��Դ = "", "NULL", "'" & str��Դ & "'") & ","
    '  ��Դ���_In     In ��������.��Դ���%Type := Null,
    gstrSQL = gstrSQL & IIf(str��Դ = "", "NULL", "'" & str��Դ & "'") & ","
    '  ɢװ��λ_In     In �շ���ĿĿ¼.���㵥λ%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo��λ(0).Text) = "", "NULL", "'" & Trim(cbo��λ(0).Text) & "'") & ","
    '  ��װ��λ_In     In ��������.��װ��λ%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo��λ(1).Text) = "", "NULL", "'" & Trim(cbo��λ(1).Text) & "'") & ","
    '  ����ϵ��_In     In ��������.����ϵ��%Type := Null,
    strValues = Val(txtEdit(GetTxtIdx("����ϵ��")).Text):
    gstrSQL = gstrSQL & strValues & ","
    '  �Ƿ���_In     In �շ���ĿĿ¼.�Ƿ���%Type := Null,
    gstrSQL = gstrSQL & IIf(cbo�۸�����.ItemData(cbo�۸�����.ListIndex) = 0, 0, 1) & ","
    '  ָ��������_In   In ��������.ָ��������%Type := Null,
    gstrSQL = gstrSQL & dblָ������ & ","
    '  ����_In         In ��������.����%Type := 95,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("�ɹ�����")).Text) & ","
    '  ָ�����ۼ�_In   In ��������.ָ�����ۼ�%Type := Null,
    gstrSQL = gstrSQL & dblָ���ۼ� & ","
    '  ָ�������_In   In ��������.ָ�������%Type := Null,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("ָ�������")).Text) & ","
    '  ��������_In     In �շ���ĿĿ¼.��������%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo��������.Text) = "", "NULL", "'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'") & ","
    '  �������_In     In �շ���ĿĿ¼.�������%Type := Null,
    gstrSQL = gstrSQL & cbo�������.ItemData(cbo�������.ListIndex) & ","
    '  ���ηѱ�_In     In �շ���ĿĿ¼.���ηѱ�%Type := 0,
    gstrSQL = gstrSQL & IIf(chk���ηѱ�.Value = 1, 1, 0) & ","
    '  �ⷿ����_In     In ��������.�ⷿ����%Type := Null,
    gstrSQL = gstrSQL & IIf(chk�ⷿ.Value = 1, 1, 0) & ","
    '  ���÷���_In     In ��������.���÷���%Type := Null,
    gstrSQL = gstrSQL & IIf(chk����.Value = 1, 1, 0) & ","
    '  ���Ч��_In     In ��������.���Ч��%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("������")).Text)
    gstrSQL = gstrSQL & IIf(Val(strValues) <> 0, Val(strValues), "NULL") & ","
    '  ���Ч��_In     In ��������.���Ч��%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("���Ч��")).Text)
    gstrSQL = gstrSQL & IIf(Val(strValues) <> 0, Val(strValues), "NULL") & ","
    '  �޾��Բ���_In   In ��������.�޾��Բ���%Type := Null,
    gstrSQL = gstrSQL & IIf(chk�޾��Բ���.Value = 1, 1, 0) & ","
    '  һ���Բ���_In   In ��������.һ���Բ���%Type := Null,
    gstrSQL = gstrSQL & IIf(Chkһ���Բ���.Value = 1, 1, 0) & ","
    '  ԭ����_In       In ��������.ԭ����%Type := Null,
    gstrSQL = gstrSQL & IIf(chkԭ��.Value = 1, 1, 0) & ","
    '  ���������_In   In ��������.���������%Type := 0,
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("�������")).Text) & ","
    '  �ɱ���_In       In ��������.�ɱ���%Type := 0,
    gstrSQL = gstrSQL & dbl�ɱ��۸� & ","
    '  ��������_In     In ��������.��������%Type := Null,
    gstrSQL = gstrSQL & chk����.Value & ","
    '  �������_In     In ��������.�������%Type := 0,
    gstrSQL = gstrSQL & IIf(chk����.Value = 1, chk�������.Value, 0) & ","
    '  ��ǰ�ۼ�_In     In �շѼ�Ŀ.�ּ�%Type := 0,
    gstrSQL = gstrSQL & dbl��ǰ�ۼ� & ","
    '  ����id_In       In �շѼ�Ŀ.������Ŀid%Type := Null,
    gstrSQL = gstrSQL & cbo������Ŀ.ItemData(cbo������Ŀ.ListIndex) & ","
    '  ��׼�ĺ�_In     In ��������.��׼�ĺ�%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txt��׼�ĺ�.Text) = "", "NULL", "'" & Trim(txt��׼�ĺ�.Text) & "'") & ","
    '  ע���̱�_In     In ��������.ע���̱�%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txtע���̱�.Text) = "", "NULL", "'" & Trim(txtע���̱�.Text) & "'") & ","
    '  ע��֤��_In     In ��������.ע��֤��%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(txtע��֤��.Text) = "", "NULL", "'" & Trim(txtע��֤��.Text) & "'") & ","
    '  ���֤��_In     In ��������.���֤��%Type := Null,
    strValues = Trim(txtEdit(GetTxtIdx("���֤��")).Text)
    gstrSQL = gstrSQL & IIf(strValues = "", "NULL", "'" & strValues & "'") & ","
    '  ���֤��Ч��_In In ��������.���֤��Ч��%Type := Null,
    If dtp���֤Ч��.Value = "" Or IsNull(dtp���֤Ч��.Value) Then
        gstrSQL = gstrSQL & "NULL" & ","
    Else
        gstrSQL = gstrSQL & "To_date('" & Format(dtp���֤Ч��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    End If
    '  ���ʷ���_In     In ��������.���ʷ���%Type := Null,
    gstrSQL = gstrSQL & IIf(Trim(cbo���ʷ���.Text) = "", "NULL", "'" & Mid(Me.cbo���ʷ���.Text, InStr(1, Me.cbo���ʷ���.Text, "-") + 1) & "'") & ","
    '  �洢����_In     In ��������.�洢����%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(cbo�洢����.Text) = "", "NULL", "'" & Mid(Me.cbo�洢����.Text, InStr(1, Me.cbo�洢����.Text, "-") + 1) & "'") & ","
    '  ���ٲ���_In     In ��������.���ٲ���%Type := 0
    gstrSQL = gstrSQL & IIf(chk���ٲ���.Value = 1, "1", "0") & ","
    '  վ��_In         In �շ���ĿĿ¼.վ��%Type := Null
    gstrSQL = gstrSQL & IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & strվ�� & "'", "NULL") & ","
    '  Ʒ��_In         In �շ���Ŀ����.����%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("��Ʒ��")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("��Ʒ��")).Text) & "'") & ","
    '  ƴ��_In         In �շ���Ŀ����.����%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("ƴ������")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("ƴ������")).Text) & "'") & ","
    '  ���_In         In �շ���Ŀ����.����%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("��ʼ���")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("��ʼ���")).Text) & "'") & ","
    '  ��ֵ˰��_In     In ��������.��ֵ˰��%Type := Null
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("��ֵ˰��")).Text) & ","
    '  ˵��_In         In �շ���ĿĿ¼.˵��%Type := Null
    gstrSQL = gstrSQL & IIf(Trim(txtEdit(GetTxtIdx("˵��")).Text) = "", "NULL", "'" & Trim(txtEdit(GetTxtIdx("˵��")).Text) & "'") & ","
    '  ��ֵ����        In ��������.��ֵ����%Type := Null
    gstrSQL = gstrSQL & IIf(chkCostly.Value = 1, 1, 0) & ","
    '  �������        In ��������.�Ƿ��������%Type := Null
    gstrSQL = gstrSQL & IIf(chkCode.Value = 1, 1, 0) & ",'"
    '   ������Ŀ
    gstrSQL = gstrSQL & txt������Ŀ.Text & "',"
    '   ��е������
    gstrSQL = gstrSQL & IIf(chkInstrument.Value = 1, chkInstrument.Value, 0) & ","
    '  ע��֤��Ч��_In In ��������.ע��֤��Ч��%Type := Null
    If dtpע��֤��Ч��.Value = "" Or IsNull(dtpע��֤��Ч��.Value) Then
        gstrSQL = gstrSQL & "NULL,"
    Else
        gstrSQL = gstrSQL & "To_date('" & Format(dtpע��֤��Ч��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    End If
    '  �Ƿ�ֲ��Ĳ�
    gstrSQL = gstrSQL & IIf(chkֲ��Ĳ�.Value = 1, 1, 0) & ","
    
    If mintEditType = g�޸� Then
        gstrSQL = gstrSQL & mintע���޸Ĳ��� & ","
    End If
    '  �ӳ���
    gstrSQL = gstrSQL & Val(txtEdit(GetTxtIdx("�ӳ���")).Text) & ","
    '  ����ʹ��_In
    gstrSQL = gstrSQL & IIf(chk����ʹ��.Value = 1, 1, 0)
    gstrSQL = gstrSQL & ")"

    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chk�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Chkһ���Բ���_Click()
    If mintEditType = g�鿴 Then Exit Sub
    If Chkһ���Բ���.Value = 1 Then
        txtEdit(7).Enabled = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0
    Else
        'ֻ��һ���Բ��ϲ������Ч�ڡ�
        txtEdit(7).Enabled = False
        txtEdit(7).Text = ""
    End If
    
    SetCtlBackColor txtEdit(7)
End Sub

Private Sub Chkһ���Բ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkԭ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt������Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkInstrument_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkCostly_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub chkCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub cmbStationNo_Change()
    mblnChange = True
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    stbSpec.Tab = 1
    If cbo�۸�����.Enabled Then cbo�۸�����.SetFocus
End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    Call cmdOK_Click
End Sub

Private Sub cmd����_Click()
    On Error GoTo ErrHandle
    Dim strSql As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    Call Sel����("")
End Sub
Private Sub cmdOK_Click()
    
    Dim i As Long
    '�����ҳ����������Ƿ���ȷ
    If ISValied = False Then Exit Sub
    

    If mintEditType <> g���� And mintEditType <> g�޸� Then
        Unload Me
        Exit Sub
    End If
    
    If mintEditType = g�޸� Then
        If mstrע��֤�� <> Trim(txtע��֤��) Or mstrע��֤��Ч�� <> CStr(Format(dtpע��֤��Ч��.Value, "yyyy-mm-dd")) Then
            If MsgBox("�Ƿ�ѡ�ע��֤�š��͡�ע��֤��Ч�ڡ�ͬ���޸ĵ���Ʒ�������й��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mintע���޸Ĳ��� = 1
            Else
                mintע���޸Ĳ��� = 0
            End If
        Else
            mintע���޸Ĳ��� = 0
        End If
    End If
    
    If SaveData = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    '�������ֵ
    Call SaveReg
    
    If mintEditType = g���� Then
        If ActiveControl Is cmdOK Then   '��ͨģʽ
            Unload Me
        ElseIf ActiveControl Is cmdSaveAddSpec Then        '�������ӹ��
            For i = 0 To cbo��λ(0).ListCount
                If Trim(cbo��λ(0).Text) = cbo��λ(0).List(i) Then
                    cbo��λ(0).ListIndex = i: i = -1: Exit For
                End If
            Next
            If i >= 0 Then
                cbo��λ(0).AddItem Trim(cbo��λ(0).Text)
                cbo��λ(0).ListIndex = cbo��λ(0).NewIndex
            End If
            
            Call InitCardData(False)
            
            Me.stbSpec.Tab = 0
            If txtEdit(GetTxtIdx("���")).Enabled Then txtEdit(GetTxtIdx("���")).SetFocus
        ElseIf ActiveControl Is cmdSaveAddItem Then '��������Ʒ��
            Unload Me
            If frmStuffBreed.ShowEditCard(frmStuffMgr, g����, "", mlng����id, gstrPrivs) = False Then
                Exit Sub
            End If
        End If
    Else
        Unload Me
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub
Private Sub cmd����_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Function GetMaxCode() As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����
    '--�����:
    '--������:
    '--��  ��:�����
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCode As ADODB.Recordset
    Dim strTemp As String
    Dim intCodeType As Integer
    Dim str���� As String
    
    On Error GoTo ErrHandle
    intCodeType = Val(zlDatabase.GetPara("�������ģʽ", glngSys, mlngModule))
    gstrSQL = "Select ���� From ������ĿĿ¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    strTemp = zlStr.nvl(rsTemp!����)
    
    If intCodeType = 0 Or Len(strTemp) >= 17 Then
        'ȡ������
        gstrSQL = "Select Nvl(����, '00000000000000') As ����" & vbNewLine & _
                        "From (Select ����" & vbNewLine & _
                        "       From �շ���ĿĿ¼ A, �������� B" & vbNewLine & _
                        "       Where a.��� = '4' And a.Id = b.����id" & vbNewLine & _
                        "       Order By Length(����) Desc, ���� Desc)" & vbNewLine & _
                        "Where Rownum = 1"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            str���� = zlCommFun.IncStr(!����)
            GetMaxCode = str����
        End With
   
    Else

        gstrSQL = "Select a.Id, c.����, c.����, c.���" & vbNewLine & _
                        "From ������ĿĿ¼ A, �������� B, �շ���ĿĿ¼ C" & vbNewLine & _
                        "Where a.Id = b.����id And b.����id = c.Id And a.����id In (Select ID From ���Ʒ���Ŀ¼ Where ���� = 7) And a.Id =[1] " & vbNewLine & _
                        "Order By ID"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        
        If Len(strTemp) >= 14 Or 14 - Len(strTemp) - IIf(intCodeType = 1, 1, 0) = 0 Then
            str���� = "01"
            str���� = IIf(intCodeType = 1, "4", "") & strTemp & str����
        Else
            str���� = Mid("00000000000000", 1, 14 - Len(strTemp) - IIf(intCodeType = 1, 1, 0))
            str���� = IIf(intCodeType = 1, "4", "") & strTemp & str����
            str���� = zlCommFun.IncStr(str����)
        End If
        
        GetMaxCode = str����
    
        Do While True
            rsTemp.Filter = ""
            rsTemp.Filter = "����='" & GetMaxCode & "'"
            If rsTemp.RecordCount = 0 Then
                Exit Do
            End If
            GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "Select ���� From �շ���ĿĿ¼ "
    Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    Do While True
        rsCode.Filter = ""
        rsCode.Filter = "����='" & GetMaxCode & "'"
        If rsCode.RecordCount = 0 Then
            Exit Do
        End If
        GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    Loop

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Set����()
    '�ⷿ������������
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mintSet���� = 0 Then
        gstrSQL = "Select b.�ⷿ����, b.���÷���" & _
                   " From �������� B, (Select Max(a.Id) As ID From �շ���ĿĿ¼ A, �������� B Where a.Id = b.����id And b.����id = [1]) C" & _
                   " Where b.����id = c.Id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ⷿ��������", mlng����ID)
        
        If rsTemp.RecordCount > 0 Then
            chk�ⷿ.Value = IIf(IsNull(rsTemp!�ⷿ����), "0", rsTemp!�ⷿ����)
            chk����.Value = IIf(IsNull(rsTemp!���÷���), "0", rsTemp!���÷���)
            chk�ⷿ.Enabled = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0
            chk����.Enabled = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0
        End If
    ElseIf mintSet���� = 1 Then
        chk�ⷿ.Value = 1
        chk����.Value = 0
        chk�ⷿ.Enabled = False
        chk����.Enabled = False
    ElseIf mintSet���� = 2 Then
        chk�ⷿ.Value = 1
        chk����.Value = 1
        chk�ⷿ.Enabled = False
        chk����.Enabled = False
    ElseIf mintSet���� = 3 Then
        chk�ⷿ.Value = 0
        chk����.Value = 0
        chk�ⷿ.Enabled = False
        chk����.Enabled = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitCardData(Optional bln��λ As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ��Ƭ����
    '--�����:bln��λ-�Ƿ����»�ȡ��λ
    '--������:
    '--��  ��:���سɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim str���㵥λ As String
    Dim rsTemp As New ADODB.Recordset
    Dim dbl����ϵ�� As Double
    '--�ָ�����ֵ
    On Error GoTo ErrHandle
    
    Call LoadReg    '����ע����Ϣֵ
    If bln��λ Then
        '�����ǰ��λ
        gstrSQL = " Select distinct a.���㵥λ,b.��װ��λ From �շ���ĿĿ¼ a,�������� b where a.id=b.����id and b.����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        cbo��λ(0).Clear
        cbo��λ(1).Clear
        
        With rsTemp
            .Sort = "���㵥λ"
            str���㵥λ = ""
            Do While Not .EOF
                If str���㵥λ <> zlStr.nvl(!���㵥λ) Then
                    str���㵥λ = zlStr.nvl(!���㵥λ)
                    cbo��λ(0).AddItem zlStr.nvl(!���㵥λ)
                End If
                .MoveNext
            Loop
            .Sort = "��װ��λ"
            str���㵥λ = ""
            Do While Not .EOF
                If str���㵥λ <> zlStr.nvl(!��װ��λ) Then
                    str���㵥λ = zlStr.nvl(!��װ��λ)
                    cbo��λ(1).AddItem zlStr.nvl(!��װ��λ)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    'ȷ��������Ϣ
    gstrSQL = "Select id, ����,����,���㵥λ,վ�� from ������ĿĿ¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
   If Not rsTemp.EOF Then
        '��ȡվ����Ϣ
        With cmbStationNo
            For i = 1 To .ListCount - 1
                If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!վ��) Then
                    .ListIndex = i: Exit For
                End If
            Next
        End With
        
        
        lblƷ��˵��.Caption = "Ʒ����Ϣ��[" & zlStr.nvl(rsTemp!����) & "] " & zlStr.nvl(rsTemp!����)
        str���㵥λ = zlStr.nvl(rsTemp!���㵥λ)
        gstrSQL = "Select ���� From ������Ŀ���� where ������Ŀid=[1] and ����=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        If rsTemp.RecordCount = 0 Then
            lblƷ��˵��.Caption = lblƷ��˵��.Caption & Space(8) & "Ӣ������:"
        Else
            lblƷ��˵��.Caption = lblƷ��˵��.Caption & Space(8) & "Ӣ������:" & zlStr.nvl(rsTemp!����)
        End If
        For i = 0 To cbo��λ(0).ListCount
            If str���㵥λ = cbo��λ(0).List(i) Then
                If cbo��λ(0).ListIndex < 0 Then
                    cbo��λ(0).ListIndex = i
                End If
                i = -1
                Exit For
            End If
        Next
        If i <> -1 Then
            cbo��λ(0).AddItem str���㵥λ
            cbo��λ(0).ListIndex = cbo��λ(0).NewIndex
        End If
        If cbo��λ(0).ListIndex >= 0 Then
            str���㵥λ = Trim(cbo��λ(0).Text)
        End If
   Else
        ShowMsgBox "������ָ����Ʒ��,���ܼ���!"
        Exit Function
   End If
   
    '--ȡȱʡ������
   cbo������Ŀ.Tag = Val(zlDatabase.GetPara("������Ŀ��Ӧ", glngSys, mlngModule))
    For mintCount = 0 To Me.cbo������Ŀ.ListCount - 1
        If Me.cbo������Ŀ.ItemData(mintCount) = Val(Me.cbo������Ŀ.Tag) Then
            Me.cbo������Ŀ.ListIndex = mintCount: Exit For
        End If
    Next
       
   If mintEditType = g���� Then
        '����ʱ��������ȡ����ţ���չ���������
        Call ��ȡ�ϴ�¼������Ϣ(mlng����ID)
        Me.txtEdit(GetTxtIdx("������")).Text = "": Me.txtEdit(GetTxtIdx("���")).Text = "": Me.txtEdit(GetTxtIdx("������")).Text = "": Me.lblFound.Caption = ""
        gstrSQL = "Select ���� from �շ���ĿĿ¼ where 1=2"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        mintCodeLength = rsTemp.Fields("����").DefinedSize
        txtEdit(0).MaxLength = mintCodeLength
        Me.txtEdit(0).Text = GetMaxCode
        'Ĭ�Ϲ��
        With cbo��λ(0)
            For i = 0 To .ListCount
                If .List(i) = str���㵥λ Then
                    .ListIndex = i
                End If
            Next
            If .ListIndex >= 0 Then
                If .List(.ListIndex) <> str���㵥λ Then
                    .AddItem str���㵥λ
                End If
            End If
        End With
        dtp���֤Ч��.Value = sys.Currentdate
        dtp���֤Ч��.Value = ""
        dtpע��֤��Ч��.Value = sys.Currentdate
        dtpע��֤��Ч��.Value = ""
        
        Call Set����
        
        Exit Function
   End If

   '�������ȡ��Ƭ����
    '----------����װ��-------------------------------------
    gstrSQL = "select I.���� as ������,I.����,I.���,I.���� as ������,S.��Դ���,S.������Դ, " & _
             "        I.���㵥λ,S.����ϵ��,S.��װ��λ,I.�Ƿ���,S.ָ�������� as ָ������,S.���� as �ɹ�����,S.ָ�����ۼ� as ָ���ۼ�," & _
             "        S.ָ�������,S.�ӳ���,S.��������� as �������,S.�ɱ��� as �ɱ��۸�, " & _
             "        I.��ʶ����,I.��ʶ����,i.������Ŀ,I.��ѡ��,I.��������,I.�������,I.���ηѱ�, " & _
             "        S.�ⷿ����,S.���÷���,S.���Ч�� as ������,S.���Ч��,S.�޾��Բ���," & _
             "        S.һ���Բ���,S.ԭ����,S.��׼�ĺ�,S.ע���̱�,S.ע��֤��,s.ע��֤��Ч��,S.��ֵ����,S.�Ƿ��������," & IIf(mblnInStrument, " s.��е�����ĵ���, ", "") & _
             "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��," & _
             "        S.��������,S.�������,S.���֤��,S.���֤��Ч��,S.���ʷ���,S.�洢����,S.���ٲ���,I.վ��,S.��ֵ˰��,I.˵��,S.�Ƿ�ֲ��Ĳ�,S.�Ƿ����  " & _
             "  from �շ���ĿĿ¼ I,�������� S " & _
             "  where I.ID=S.����ID and I.id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
             
    
    Dim strFieldsName As String
    
    With rsTemp
        mintCodeLength = .Fields("������").DefinedSize
        txtEdit(0).MaxLength = mintCodeLength
        If .RecordCount > 0 Then
            txt��׼�ĺ�.Text = zlStr.nvl(!��׼�ĺ�)
            txtע���̱�.Text = zlStr.nvl(!ע���̱�)
            txtע��֤��.Text = zlStr.nvl(!ע��֤��)
            mstrע��֤�� = zlStr.nvl(!ע��֤��)
            txt��ѡ��.Text = zlStr.nvl(!��ѡ��)
            For i = 0 To txtEdit.UBound
                strFieldsName = txtEdit(i).Tag
                Select Case strFieldsName
                Case "���Ч��", "������"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), "######")
                Case "ָ������", "�ɱ��۸�"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!����ϵ��, 1)), mFMT.FM_�ɱ���)
                Case "ָ���ۼ�"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!����ϵ��, 1)), mFMT.FM_���ۼ�)
                Case "�ɹ�����", "ָ�������", "�������", "�ӳ���"
                        txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBCJL)
                Case "ƴ��", "���", "�����", "��ǰ�ۼ�"
                Case "���֤��"
                    txtEdit(i).Text = zlStr.nvl(!���֤��)
                Case "���֤��Ч��"
                Case "��ֵ˰��"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBJCL)
                Case Else
                    '����Ʒ��,ƴ������,��ʼ��롱��������
                    If InStr("��Ʒ��;ƴ������;��ʼ���", strFieldsName) = 0 Then txtEdit(i).Text = zlStr.nvl(.Fields(strFieldsName))
                End Select
            Next
            chk���ٲ���.Value = IIf(Val(zlStr.nvl(!���ٲ���)) = 1, 1, 0)
            chk�������.Value = IIf(Val(zlStr.nvl(!�������)) = 1, 1, 0)
            Me.chk�������.Enabled = (Me.chk����.Value = 1)
            
            If IsNull(!���֤��Ч��) Then
                dtp���֤Ч��.Value = ""
            Else
                dtp���֤Ч��.Value = Format(!���֤��Ч��, "yyyy-mm-dd")
            End If
            If IsNull(!ע��֤��Ч��) Then
                dtpע��֤��Ч��.Value = ""
                mstrע��֤��Ч�� = ""
            Else
                dtpע��֤��Ч��.Value = Format(!ע��֤��Ч��, "yyyy-mm-dd")
                mstrע��֤��Ч�� = Format(!ע��֤��Ч��, "yyyy-mm-dd")
            End If
            
            '���㵥λ
            For mintCount = 0 To Me.cbo��λ(0).ListCount - 1
                If cbo��λ(0).List(mintCount) = zlStr.nvl(!���㵥λ) Then
                    cbo��λ(0).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo��λ(0).ListIndex < 0 Then
                If zlStr.nvl(!���㵥λ) <> "" Then
                    cbo��λ(0).AddItem zlStr.nvl(!���㵥λ)
                    cbo��λ(0).ListIndex = cbo��λ(0).NewIndex
                End If
            End If
            '��װ��λ
            For mintCount = 0 To Me.cbo��λ(1).ListCount - 1
                If cbo��λ(1).List(mintCount) = zlStr.nvl(!��װ��λ) Then
                    cbo��λ(1).ListIndex = mintCount
                    Exit For
                End If
            Next
            If cbo��λ(1).ListIndex < 0 Then
                If zlStr.nvl(!��װ��λ) <> "" Then
                    cbo��λ(1).AddItem zlStr.nvl(!���㵥λ)
                    cbo��λ(1).ListIndex = cbo��λ(1).NewIndex
                End If
            End If
            dbl����ϵ�� = zlStr.nvl(!����ϵ��, 1)
            '--������Դ
            For mintCount = 0 To Me.cbo��Դ.ListCount - 1
                If Mid(Me.cbo��Դ.List(mintCount), InStr(1, Me.cbo��Դ.List(mintCount), "-") + 1) = zlStr.nvl(!��Դ���) Then
                    Me.cbo��Դ.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo������Դ.ListCount - 1
                If Mid(Me.cbo������Դ.List(mintCount), InStr(1, Me.cbo������Դ.List(mintCount), "-") + 1) = zlStr.nvl(!������Դ) Then
                    Me.cbo������Դ.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo���ʷ���.ListCount - 1
                If Mid(Me.cbo���ʷ���.List(mintCount), InStr(1, Me.cbo���ʷ���.List(mintCount), "-") + 1) = zlStr.nvl(!���ʷ���) Then
                    Me.cbo���ʷ���.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo�洢����.ListCount - 1
                If Mid(Me.cbo�洢����.List(mintCount), InStr(1, Me.cbo�洢����.List(mintCount), "-") + 1) = zlStr.nvl(!�洢����) Then
                    Me.cbo�洢����.ListIndex = mintCount: Exit For
                End If
            Next
            
            
          'ʱ��
            For mintCount = 0 To Me.cbo�۸�����.ListCount - 1
                If cbo�۸�����.ItemData(mintCount) = zlStr.nvl(!�Ƿ���, 0) Then
                    cbo�۸�����.ListIndex = mintCount
                    Exit For
                End If
            Next
            
            lbl���۵�λ(0).Caption = "Ԫ/" & IIf(mintUnit = 0, zlStr.nvl(!���㵥λ), zlStr.nvl(!��װ��λ))
            lbl���۵�λ(1).Caption = "Ԫ/" & IIf(mintUnit = 0, zlStr.nvl(!���㵥λ), zlStr.nvl(!��װ��λ))
            
            cbo�۸�����.ListIndex = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
            
            Chkһ���Բ���.Value = zlStr.nvl(!һ���Բ���, 0): Call Chkһ���Բ���_Click
            chk�޾��Բ���.Value = zlStr.nvl(!�޾��Բ���, 0)
            chkԭ��.Value = zlStr.nvl(!ԭ����, 0)
            chkֲ��Ĳ�.Value = zlStr.nvl(!�Ƿ�ֲ��Ĳ�, 0)
            
            For mintCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(mintCount), InStr(1, Me.cbo��������.List(mintCount), "-") + 1) = IIf(IsNull(!��������), "", !��������) Then
                    Me.cbo��������.ListIndex = mintCount: Exit For
                End If
            Next
            
            Me.cbo�������.ListIndex = IIf(IsNull(!�������), 0, !�������)
            Me.chk���ηѱ�.Value = IIf(IsNull(!���ηѱ�), 0, !���ηѱ�)
            
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "ע���ù����" & Format(!����ʱ��, "YYYY��MM��DD��") & "������" & Format(!����ʱ��, "YYYY��MM��DD��") & "ͣ��"
            Else
                Me.lblFound.Caption = ""
            End If
            
            Me.chk�ⷿ.Value = zlStr.nvl(!�ⷿ����, 0)
            Me.chk����.Value = zlStr.nvl(!���÷���, 0)
            Me.chk����.Value = zlStr.nvl(!��������, 0)
            Me.chkCostly.Value = zlStr.nvl(!��ֵ����, 0)
            Me.chkCode.Value = zlStr.nvl(!�Ƿ��������, 0)
            Me.chk����ʹ��.Value = zlStr.nvl(!�Ƿ����, 0)
            If mblnInStrument = True Then
                Me.chkInstrument.Value = zlStr.nvl(!��е�����ĵ���, 0)
            End If

            Me.chk����.Tag = Me.chk����.Value
            txt������Ŀ.Text = IIf(IsNull(!������Ŀ), "", !������Ŀ)
            
            If Me.chk�ⷿ.Value = 0 Then
                Me.chk����.Enabled = False: Me.chk����.Value = 0
            Else
                Me.chk����.Enabled = True
                Me.chk����.Value = Me.chk����.Tag
            End If
            
            '��ȡվ����Ϣ
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!վ��) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
            
        End If
        
   End With
         
    '��ȡ��Ʒ���ͼ���
    gstrSQL = "select ����,����,����,���� from �շ���Ŀ���� where �շ�ϸĿid=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
    With rsTemp
        Me.txtEdit(19).MaxLength = .Fields("����").DefinedSize
        Me.txtEdit(20).MaxLength = .Fields("����").DefinedSize
        Me.txtEdit(GetTxtIdx("��Ʒ��")).MaxLength = .Fields("����").DefinedSize
        Do While Not .EOF
            If !���� = 3 And !���� = 1 Then
                Me.txtEdit(GetTxtIdx("��Ʒ��")).Text = IIf(IsNull(!����), "", !����)
                Me.txtEdit(GetTxtIdx("ƴ������")).Text = IIf(IsNull(!����), "", !����)
            End If
            If !���� = 3 And !���� = 2 Then
                Me.txtEdit(GetTxtIdx("��Ʒ��")).Text = IIf(IsNull(!����), "", !����)
                Me.txtEdit(GetTxtIdx("��ʼ���")) = IIf(IsNull(!����), "", !����)
            End If
            .MoveNext
        Loop
    End With
         
    '��ȡ��ʾ��ǰ�ۼ�
    If Me.cbo�۸�����.ListIndex <> 0 Then
        'ʱ�۲��ϣ�ȡ�����/���������Ϊ��۸��޿��ʱȡ�۱���
        gstrSQL = "select Decode(K.�������,0,P.�ּ�,K.�����/Nvl(K.�������,1)) as �ּ�,P.������Ŀid" & _
                " from �շѼ�Ŀ P," & _
                "     (Select nvl(Sum(ʵ�ʽ��),0) as �����,nvl(Sum(ʵ������),0) as �������" & _
                "      From ҩƷ��� Where ҩƷID=[1]) K" & _
                " where P.�շ�ϸĿid=[1] and (Sysdate Between p.ִ������ And p.��ֹ���� or Sysdate>=p.ִ������ And p.��ֹ���� Is Null)" & _
                GetPriceClassString("P")
    Else
        '��ʱ�۲��ϵ��ۣ�ȡ��۸��¼�еļ۸�
        gstrSQL = "select P.�ּ�,P.������Ŀid" & _
                " from �շѼ�Ŀ P" & _
                " where P.�շ�ϸĿid=[1] and (Sysdate Between p.ִ������ And p.��ֹ���� or Sysdate>=p.ִ������ And p.��ֹ���� Is Null)" & _
                GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtEdit(16).Text = Format(!�ּ� * IIf(mintUnit = 0, 1, dbl����ϵ��), mFMT.FM_���ۼ�)
            For mintCount = 0 To Me.cbo������Ŀ.ListCount - 1
                If Me.cbo������Ŀ.ItemData(mintCount) = !������Ŀid Then
                    Me.cbo������Ŀ.ListIndex = mintCount: Exit For
                End If
            Next
        End If
    End With

    If Val(mstr����ID) <> 0 Then
        '--��ִ֤������
        gstrSQL = "Select ID from �շѼ�Ŀ where �շ�ϸĿid=[1] and nvl(�䶯ԭ��,0)=0" & _
                GetPriceClassString("")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
        
        Do While Not rsTemp.EOF
                gstrSQL = "zl_�����շ���¼_Adjust(" & Val(zlStr.nvl(rsTemp!Id)) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                rsTemp.MoveNext
        Loop
    End If
    
    '�����Ƿ��з�����ȷ�����������ԡ��ɱ��۸����ۼ۸���޸ķ�
    gstrSQL = " Select nvl(Count(*),0) From ҩƷ�շ���¼ Where ҩƷID=[1] And rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
    
    With rsTemp
        If .Fields(0).Value > 0 Then
            Me.cbo�۸�����.Enabled = False
            Me.txtEdit(15).Enabled = False
            Me.txtEdit(16).Enabled = False
'            Me.cbo������Ŀ.Enabled = False
        Else
            Me.cbo�۸�����.Enabled = cbo�۸�����.Enabled
            Me.txtEdit(15).Enabled = Me.txtEdit(15).Enabled  '�ɱ���
            Me.txtEdit(16).Enabled = cbo�۸�����.Enabled       '��ǰ�ۼ�
'            Me.cbo������Ŀ.Enabled = cbo�۸�����.Enabled
        End If
        SetCtlBackColor txtEdit(15)
        SetCtlBackColor txtEdit(16)
    End With
    
    If Me.chk����.Value = 1 Then
'        Me.chk����.Enabled = Me.cbo�۸�����.Enabled
        chk����.Tag = 1
    Else
        chk����.Tag = 0
    End If
    
    If Val(mstr����ID) <> 0 Then
        '�������δִ�еļ۸�,�򲻳����޸���ؼ۸�
        gstrSQL = "Select �ּ� from �շѼ�Ŀ where �շ�ϸĿid=[1] and nvl(�䶯ԭ��,0)=0" & _
                GetPriceClassString("")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
        
        If rsTemp.EOF = False Then
            Me.cbo�۸�����.Enabled = False
            'Me.txtEDIT(15).Enabled = False
            Me.txtEdit(16).Enabled = False
            Me.cbo������Ŀ.Enabled = False
        End If
    End If
    '�����Ƿ��п�棬ȷ�����������Կ��޸ķ�
    
    gstrSQL = "" & _
        "   Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
        "   Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And B.�������� In ('���Ŀ�','���ʿⷿ', '����ⷿ')"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
        
    With rsTemp
        
        If .Fields(0).Value > 0 Then
            Me.chk�ⷿ.Enabled = False
        Else
            Me.chk�ⷿ.Enabled = True
        End If
    End With
    
    If Me.chk�ⷿ.Value = 1 Then
        gstrSQL = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                 " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '���ϲ���' Or B.�������� Like '%�Ƽ���')"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstr����ID))
                
        With rsTemp
            If .Fields(0).Value > 0 Then
                Me.chk����.Enabled = False
                If Me.chk�ⷿ.Enabled Then Me.chk�ⷿ.Enabled = IIf(chk����.Value = 1, False, True)
            Else
                Me.chk����.Enabled = True
            End If
        End With
    End If
    
    InitCardData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ��ȡ�ϴ�¼������Ϣ(ByVal lng����ID As Long)
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ϴ�¼��Ĺ����Ϣ
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, dbl����ϵ�� As Double
    gstrSQL = "Select ID as ����id" & _
              "  From �շ���ĿĿ¼ A," & _
              "      (Select Max(a.����ʱ��) As ����ʱ�� From �շ���ĿĿ¼ A, �������� B Where a.Id = b.����id And b.����id = [1]) B " & _
              "  Where a.����ʱ�� = b.����ʱ�� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then Exit Sub
    lng����ID = Val(zlStr.nvl(rsTemp!����ID))
    If lng����ID = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
   '�������ȡ��Ƭ����
    '----------����װ��-------------------------------------
    gstrSQL = "select I.���� as ������,I.����,I.���,I.���� as ������,S.��Դ���,S.������Դ, " & _
             "        I.���㵥λ,S.����ϵ��,S.��װ��λ,I.�Ƿ���,S.ָ�������� as ָ������,S.���� as �ɹ�����,S.ָ�����ۼ� as ָ���ۼ�," & _
             "        S.ָ�������,S.�ӳ���,S.��������� as �������,S.�ɱ��� as �ɱ��۸�, " & _
             "        I.��ʶ����,I.��ʶ����,I.��������,I.�������,I.���ηѱ�, " & _
             "        S.�ⷿ����,S.���÷���,S.���Ч�� as ������,S.���Ч��,S.�޾��Բ���,S.һ���Բ���,S.ԭ����,S.��׼�ĺ�,S.ע���̱�," & _
             "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,S.��������,S.�������,s.ע��֤��Ч��,S.���֤��,S.���֤��Ч��,S.���ʷ���,S.�洢����,I.վ��,S.��ֵ˰��,I.˵��,S.�Ƿ�ֲ��Ĳ� " & _
             "  from �շ���ĿĿ¼ I,�������� S " & _
             "  where I.ID=S.����ID and I.id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
             
    
    Dim strFieldsName As String
    
    With rsTemp
        If .RecordCount > 0 Then
            txt��׼�ĺ�.Text = zlStr.nvl(!��׼�ĺ�)
            txtע���̱�.Text = zlStr.nvl(!ע���̱�)
                  
            For i = 0 To txtEdit.UBound
                strFieldsName = txtEdit(i).Tag
                Select Case strFieldsName
                Case "���Ч��", "������"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), "######")
                Case "ָ������", "�ɱ��۸�"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!����ϵ��, 1)), mFMT.FM_�ɱ���)
                Case "ָ���ۼ�"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0) * IIf(mintUnit = 0, 1, zlStr.nvl(!����ϵ��, 1)), mFMT.FM_���ۼ�)
                Case "�ɹ�����", "ָ�������", "�������", "�ӳ���"
                        txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBCJL)
                Case "ƴ��", "���", "�����", "��ǰ�ۼ�"
                Case "���֤��"
                    txtEdit(i).Text = zlStr.nvl(!���֤��)
                Case "���", "������"
                    txtEdit(i).Text = ""
                Case "��ֵ˰��"
                    txtEdit(i).Text = Format(zlStr.nvl(.Fields(strFieldsName), 0), GFM_VBJCL)
                Case Else
                    '����Ʒ��,ƴ������,��ʼ��롱��������
                    If InStr("��Ʒ��;ƴ������;��ʼ���", strFieldsName) = 0 Then txtEdit(i).Text = zlStr.nvl(.Fields(strFieldsName))
                End Select
            Next
            
            If IsNull(!���֤��Ч��) Then
                dtp���֤Ч��.Value = ""
            Else
                dtp���֤Ч��.Value = Format(!���֤��Ч��, "yyyy-mm-dd")
            End If
            If IsNull(!ע��֤��Ч��) Then
                dtpע��֤��Ч��.Value = ""
            Else
                dtpע��֤��Ч��.Value = Format(!ע��֤��Ч��, "yyyy-mm-dd")
            End If
            
            '���㵥λ
            For mintCount = 0 To Me.cbo��λ(0).ListCount - 1
                If cbo��λ(0).List(mintCount) = zlStr.nvl(!���㵥λ) Then
                    cbo��λ(0).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo��λ(0).ListIndex < 0 Then
                If zlStr.nvl(!���㵥λ) <> "" Then
                    cbo��λ(0).AddItem zlStr.nvl(!���㵥λ)
                    cbo��λ(0).ListIndex = cbo��λ(0).NewIndex
                End If
            End If
            '��װ��λ
            For mintCount = 0 To Me.cbo��λ(1).ListCount - 1
                If cbo��λ(1).List(mintCount) = zlStr.nvl(!��װ��λ) Then
                    cbo��λ(1).ListIndex = mintCount
                    Exit For
                End If
            Next
            
            If cbo��λ(1).ListIndex < 0 Then
                If zlStr.nvl(!��װ��λ) <> "" Then
                    cbo��λ(1).AddItem zlStr.nvl(!���㵥λ)
                    cbo��λ(1).ListIndex = cbo��λ(1).NewIndex
                End If
            End If
            dbl����ϵ�� = zlStr.nvl(!����ϵ��, 1)
            '--������Դ
            For mintCount = 0 To Me.cbo��Դ.ListCount - 1
                If Mid(Me.cbo��Դ.List(mintCount), InStr(1, Me.cbo��Դ.List(mintCount), "-") + 1) = zlStr.nvl(!��Դ���) Then
                    Me.cbo��Դ.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo������Դ.ListCount - 1
                If Mid(Me.cbo������Դ.List(mintCount), InStr(1, Me.cbo������Դ.List(mintCount), "-") + 1) = zlStr.nvl(!������Դ) Then
                    Me.cbo������Դ.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo���ʷ���.ListCount - 1
                If Mid(Me.cbo���ʷ���.List(mintCount), InStr(1, Me.cbo���ʷ���.List(mintCount), "-") + 1) = zlStr.nvl(!���ʷ���) Then
                    Me.cbo���ʷ���.ListIndex = mintCount: Exit For
                End If
            Next
            
            '--������Դ
            For mintCount = 0 To Me.cbo�洢����.ListCount - 1
                If Mid(Me.cbo�洢����.List(mintCount), InStr(1, Me.cbo�洢����.List(mintCount), "-") + 1) = zlStr.nvl(!�洢����) Then
                    Me.cbo�洢����.ListIndex = mintCount: Exit For
                End If
            Next
            
            
          'ʱ��
            For mintCount = 0 To Me.cbo�۸�����.ListCount - 1
                If cbo�۸�����.ItemData(mintCount) = zlStr.nvl(!�Ƿ���, 0) Then
                    cbo�۸�����.ListIndex = mintCount
                    Exit For
                End If
            Next
            
            lbl���۵�λ(0).Caption = "Ԫ/" & IIf(mintUnit = 0, zlStr.nvl(!���㵥λ), zlStr.nvl(!��װ��λ))
            lbl���۵�λ(1).Caption = "Ԫ/" & IIf(mintUnit = 0, zlStr.nvl(!���㵥λ), zlStr.nvl(!��װ��λ))
            
            cbo�۸�����.ListIndex = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
            
            Chkһ���Բ���.Value = zlStr.nvl(!һ���Բ���, 0)
            chk�޾��Բ���.Value = zlStr.nvl(!�޾��Բ���, 0)
            chkԭ��.Value = zlStr.nvl(!ԭ����, 0)
            chkֲ��Ĳ�.Value = zlStr.nvl(!�Ƿ�ֲ��Ĳ�, 0)
            
            For mintCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(mintCount), InStr(1, Me.cbo��������.List(mintCount), "-") + 1) = IIf(IsNull(!��������), "", !��������) Then
                    Me.cbo��������.ListIndex = mintCount: Exit For
                End If
            Next
            
            If InStr(1, mstrPrivs, ";�������;") <> 0 Then
                Me.cbo�������.ListIndex = IIf(IsNull(!�������), 0, !�������)
            Else
                cbo�������.Enabled = False
            End If
            
            Me.chk���ηѱ�.Value = IIf(IsNull(!���ηѱ�), 0, !���ηѱ�)
                 
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "ע���ù����" & Format(!����ʱ��, "YYYY��MM��DD��") & "������" & Format(!����ʱ��, "YYYY��MM��DD��") & "ͣ��"
            Else
                Me.lblFound.Caption = ""
            End If
            
            Me.chk�ⷿ.Value = zlStr.nvl(!�ⷿ����, 0)
            Me.chk����.Value = zlStr.nvl(!���÷���, 0)
            Me.chk����.Value = zlStr.nvl(!��������, 0)
            Me.chk�������.Value = Val(zlStr.nvl(!�������))
             
            Me.chk����.Tag = Me.chk����.Value
            
            If Me.chk�ⷿ.Value = 0 Then
                Me.chk����.Enabled = False: Me.chk����.Value = 0
            Else
                Me.chk����.Enabled = True
                Me.chk����.Value = Me.chk����.Tag
            End If
            '��ȡվ����Ϣ
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!վ��) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
        End If
        
   End With
         
    '��ȡ��ʾ��ǰ�ۼ�
    If Me.cbo�۸�����.ListIndex <> 0 Then
        'ʱ�۲��ϣ�ȡ�����/���������Ϊ��۸��޿��ʱȡ�۱���
        gstrSQL = "select Decode(K.�������,0,P.�ּ�,K.�����/Nvl(K.�������,1)) as �ּ�,P.������Ŀid" & _
                " from �շѼ�Ŀ P," & _
                "     (Select nvl(Sum(ʵ�ʽ��),0) as �����,nvl(Sum(ʵ������),0) as �������" & _
                "      From ҩƷ��� Where ҩƷID=[1]) K" & _
                " where P.�շ�ϸĿid=[1] and (Sysdate Between p.ִ������ And p.��ֹ���� or Sysdate>=p.ִ������ And p.��ֹ���� Is Null)" & _
                GetPriceClassString("P")
    Else
        '��ʱ�۲��ϵ��ۣ�ȡ��۸��¼�еļ۸�
        gstrSQL = "select P.�ּ�,P.������Ŀid" & _
                " from �շѼ�Ŀ P" & _
                " where P.�շ�ϸĿid=[1] and (Sysdate Between p.ִ������ And p.��ֹ���� or Sysdate>=p.ִ������ And p.��ֹ���� Is Null)" & _
                GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtEdit(16).Text = Format(!�ּ� * IIf(mintUnit = 0, 1, dbl����ϵ��), mFMT.FM_���ۼ�)
            For mintCount = 0 To Me.cbo������Ŀ.ListCount - 1
                If Me.cbo������Ŀ.ItemData(mintCount) = !������Ŀid Then
                    Me.cbo������Ŀ.ListIndex = mintCount: Exit For
                End If
            Next
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetPopedom()
    Dim intI As Long
    Dim blnStuffModify As Boolean
    
    chkCode.Visible = zlStr.IsHavePrivs(mstrPrivs, "�����������")
    
    If mintEditType = g���� Or mintEditType = g�޸� Then
   
        If mintEditType = g�޸� Then
            '���Ƿ������޸�������ĵ�����Ϣ
            blnStuffModify = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0
            For intI = 0 To txtEdit.UBound
                 txtEdit(intI).Enabled = blnStuffModify
            Next
            
            txt��׼�ĺ�.Enabled = blnStuffModify
            txtע���̱�.Enabled = blnStuffModify
            txtע��֤��.Enabled = blnStuffModify
            
            SetCtlBackColor txt��׼�ĺ�
            SetCtlBackColor txtע���̱�
            
            cbo��λ(0).Enabled = blnStuffModify
            cbo��λ(1).Enabled = blnStuffModify
            
            chkԭ��.Enabled = blnStuffModify
            chk���ηѱ�.Enabled = blnStuffModify
            cbo������Դ.Enabled = blnStuffModify
            chk����ʹ��.Enabled = blnStuffModify
            
            cmd����.Enabled = blnStuffModify
            cbo��Դ.Enabled = blnStuffModify
            cbo�������.Enabled = blnStuffModify
            cbo������Դ.Enabled = blnStuffModify
            cbo�洢����.Enabled = blnStuffModify
            dtp���֤Ч��.Enabled = blnStuffModify
            dtpע��֤��Ч��.Enabled = blnStuffModify
            chk�ⷿ.Enabled = blnStuffModify
            chk����.Enabled = blnStuffModify
            
            chk����.Enabled = blnStuffModify
            
            chk�������.Enabled = blnStuffModify
            cbo���ʷ���.Enabled = blnStuffModify
            chk���ٲ���.Enabled = blnStuffModify
            
            chkֲ��Ĳ�.Enabled = blnStuffModify
            Chkһ���Բ���.Enabled = blnStuffModify
            chkCostly.Enabled = blnStuffModify
            chkCode.Enabled = blnStuffModify
            chk�޾��Բ���.Enabled = blnStuffModify
            fra��������.Enabled = blnStuffModify
            txt��ѡ��.Enabled = blnStuffModify
            cmbStationNo.Enabled = blnStuffModify
            chkInstrument.Enabled = blnStuffModify
            SetCtlBackColor txt��ѡ��
            
        Else
            txt��׼�ĺ�.Enabled = True
            txtע���̱�.Enabled = True
            txt��ѡ��.Enabled = True
            SetCtlBackColor txt��ѡ��
        End If
        
        Me.txtEdit(9).Enabled = InStr(1, mstrPrivs, ";�������;") <> 0     '����
        Me.txtEdit(12).Enabled = InStr(1, mstrPrivs, ";ָ���۸����;") <> 0       'ָ�������
        Me.txtEdit(8).Enabled = Me.txtEdit(12).Enabled                          'ָ������
        Me.txtEdit(11).Enabled = Me.txtEdit(12).Enabled                          'ָ���ۼ�
        Me.txtEdit(13).Enabled = Me.txtEdit(12).Enabled                          '�ӳ���
        Me.txtEdit(14).Enabled = Me.txtEdit(12).Enabled
        
        Me.cbo�۸�����.Enabled = InStr(1, mstrPrivs, ";�ۼ۹���;") <> 0
        Me.txtEdit(15).Enabled = InStr(1, mstrPrivs, ";�ɱ��۹���;") <> 0                  '�ɱ��۸�
        Me.txtEdit(16).Enabled = Me.cbo�۸�����.Enabled                 '��ǰ�ۼ�
        Me.cbo������Ŀ.Enabled = InStr(1, mstrPrivs, ";����������Ŀ;") <> 0
        Me.cbo��������.Enabled = InStr(1, mstrPrivs, ";ҽ������Ŀ¼;") <> 0
        Me.cbo�������.Enabled = InStr(1, mstrPrivs, ";�������;") <> 0
        
        For intI = 0 To txtEdit.UBound
            SetCtlBackColor txtEdit(intI)
        Next
    
        Exit Sub
    Else
        txt������Ŀ.Enabled = False
        cmd����.Enabled = False
    End If
    For intI = 0 To txtEdit.UBound
        txtEdit(intI).Enabled = False
        SetCtlBackColor txtEdit(intI)
    Next
    
    txt��׼�ĺ�.Enabled = False
    txtע���̱�.Enabled = False
    txtע��֤��.Enabled = False
    txt��ѡ��.Enabled = False
    SetCtlBackColor txt��ѡ��
    SetCtlBackColor txt��׼�ĺ�
    SetCtlBackColor txtע���̱�
    SetCtlBackColor txtע��֤��
    
    cbo��λ(0).Enabled = False
    cbo��λ(1).Enabled = False
    
    chkԭ��.Enabled = False
    chk���ηѱ�.Enabled = False
    cbo������Դ.Enabled = False
    chk����ʹ��.Enabled = False
    
    cmd����.Enabled = False
    cbo��Դ.Enabled = False
    cbo�۸�����.Enabled = False
    cbo��������.Enabled = False
    cbo������Ŀ.Enabled = False
    cbo�������.Enabled = False
    cbo���ʷ���.Enabled = False
    cbo�洢����.Enabled = False
    dtp���֤Ч��.Enabled = False
    dtpע��֤��Ч��.Enabled = False
    chk�ⷿ.Enabled = False
    chk����.Enabled = False
    chk����.Enabled = False
    chk�������.Enabled = False
    chk���ٲ���.Enabled = False
    Chkһ���Բ���.Enabled = False
    chkֲ��Ĳ�.Enabled = False
    chkCostly.Enabled = False
    chkCode.Enabled = False
    chkInstrument.Enabled = False
    chk�޾��Բ���.Enabled = False
    fra��������.Enabled = False
    cmbStationNo.Enabled = False
    cmdOK.Visible = False
    cmdCancel.Caption = "�ر�(&C)"
End Sub

Private Sub dtp���֤Ч��_Change()
    mblnChange = True
End Sub

Private Sub dtp���֤Ч��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub dtpע��֤��Ч��_Change()
    mblnChange = True
End Sub

Private Sub dtpע��֤��Ч��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    If mblnLoad = True Then Exit Sub
    If mblnFrist = False Then Exit Sub
    mblnFrist = True
    
    '��ʼվ��
    cmbStationNo.Visible = gSystem_Para.bln����վ��
    lblStationNo.Visible = cmbStationNo.Visible
    gstrSQL = "Select Count(1) ��е������ From all_Tab_Columns Where Table_Name = '��������' And Column_Name = '��е�����ĵ���'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���װ����ϵͳ")
    If rsTemp!��е������ = 0 Then 'û�й���װ����ϵͳ
        chkInstrument.Visible = False
    End If
    
    gstrSQL = "Select Count(1) ����ϵͳ  From zlSystems Where ��� = 400 And ����� = 100"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���װ����ϵͳ")
    mblnInStrument = False
    If rsTemp!����ϵͳ > 0 Then
        mblnInStrument = True
    Else
        chkInstrument.Visible = False
    End If
    
    mintSet���� = Val(zlDatabase.GetPara("���ķ��������Զ�����", glngSys, mlngModule, 0))
    '----------������ϵ�ж�-------------------------------------
    If GetDepend = False Then
        Unload Me
        Exit Sub
    End If
    
    '----------����Ȩ�޿���-------------------------------------
    Call SetPopedom
    
    '----------��ʼ��Ƭ����-------------------------------------
    Call InitCardData
    
    '----------Ĭ�ϵ�һѡ�-----------------------------------
    If mintEditType = g�޸� Then
      If InStr(1, mstrPrivs, ";����Ʒ������;") <> 0 Then
          Me.stbSpec.Tab = 0
      Else
          Me.stbSpec.Tab = 1
      End If
    Else
      Me.stbSpec.Tab = 0
    End If
    mblnLoad = True
     
    If InStr(1, mstrPrivs, ";������������;") = 0 Then
       chk����.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub txt������Ŀ_GotFocus()
    txt������Ŀ.SelStart = 0
    txt������Ŀ.SelLength = Len(txt������Ŀ)
    txt������Ŀ.SetFocus
End Sub

Private Sub txt������Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Then
        txt������Ŀ.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Form_Load()
    Dim aryTemp() As String
    Dim strSql As String
    Dim rsrecord As ADODB.Recordset
    
    mblnFrist = True
    On Error GoTo ErrHandle
            
    If mintEditType <> g���� Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    strSql = "select ���,���� from zlnodelist"
    Set rsrecord = zlDatabase.OpenSQLRecord(strSql, "վ���ѯ")
    With cmbStationNo
        .AddItem ""
        Do While Not rsrecord.EOF
            .AddItem rsrecord!��� & "-" & rsrecord!����
            rsrecord.MoveNext
        Loop
    End With
    
    '----------------װ���ѡ�Ļ�������----------------------
    With Me.cbo�۸�����
        .Clear
        aryTemp = Split("0-����;1-ʱ��", ";")
        For mintCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(mintCount): .ItemData(.NewIndex) = mintCount
        Next
        .ListIndex = 0
    End With
    
    With Me.cbo�������
        aryTemp = Split("0-��Ӧ���ڲ���;1-����;2-סԺ;3-�����סԺ", ";")
        For mintCount = LBound(aryTemp) To UBound(aryTemp)
            .AddItem aryTemp(mintCount): .ItemData(.NewIndex) = mintCount
        Next
        If InStr(1, mstrPrivs, ";�������;") <> 0 Then
            .ListIndex = 3
        Else
            .ListIndex = 0
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnLoad = False
End Sub

Private Sub txtEDIT_Change(Index As Integer)
    Dim strTag As String
    strTag = txtEdit(Index).Tag
    
    Select Case strTag
        Case "�ɹ�����"
            '��������=ָ������*����/100
            Me.txtEdit(10).Text = Format(Val(Me.txtEdit(8).Text) * Val(Me.txtEdit(Index).Text) / 100, mFMT.FM_�ɱ���)
        Case "ָ������"
            '��������=ָ������*����/100
            Me.txtEdit(10).Text = Format(Val(Me.txtEdit(Index).Text) * Val(Me.txtEdit(9).Text) / 100, mFMT.FM_�ɱ���)
        Case "��ʶ����", "��ʶ����"
                txtEdit(Index).Text = UCase(txtEdit(Index).Text)
                txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
        Case "��Ʒ��"
            'ƴ�������
            Me.txtEdit(19).Text = zlStr.GetCodeByORCL(Me.txtEdit(Index).Text, 0, Me.txtEdit(19).MaxLength)
            Me.txtEdit(20).Text = zlStr.GetCodeByORCL(Me.txtEdit(Index).Text, 1, Me.txtEdit(20).MaxLength)
        Case "�ӳ���"
            If Val(txtEdit(Index).Text) > 9900 Then txtEdit(Index).Text = 9900
            If Val(txtEdit(Index).Text) < 0 Then txtEdit(Index).Text = 0
    End Select
End Sub

Private Sub txtEDIT_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strTag As String
    Dim intDigit As Integer
    Dim strKey As String
    
    If Index = 0 Then
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        End Select
        KeyAscii = 0
        Exit Sub
    End If
    
    strKey = txtEdit(Index).Text
    strTag = txtEdit(Index).Tag
    Select Case strTag
        Case "����ϵ��", "���Ч��", "������"
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m����ʽ
        Case "ָ������", "�ɹ�����", "�����", "ָ���ۼ�", "ָ�������", "�ӳ���", "�������", "�ɱ��۸�", "��ǰ�ۼ�", "��ֵ˰��"
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
            
            If strTag = "ָ������" Or strTag = "�����" Or strTag = "�ɱ��۸�" Or strTag = "ָ���ۼ�" Or strTag = "��ǰ�ۼ�" Then
                Select Case strTag
                    Case "ָ������", "�����", "�ɱ��۸�"
                        intDigit = Len(Mid(mFMT.FM_�ɱ���, InStr(1, mFMT.FM_�ɱ���, ".") + 1))
                    Case "ָ���ۼ�", "��ǰ�ۼ�"
                        intDigit = Len(Mid(mFMT.FM_�ɱ���, InStr(1, mFMT.FM_���ۼ�, ".") + 1))
                End Select
                
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If txtEdit(Index).SelLength = Len(strKey) Then Exit Sub
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Case Else
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    End Select
    
    If strTag = "������" Then     '����
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        If KeyAscii <> vbKeyReturn Then Exit Sub
        If txtEdit(Index).Text <> "" Then
            Call Sel����(txtEdit(Index).Text)
        End If
        Exit Sub
    End If
    
    If strTag = "���" Then
        If InStr("^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Exit Sub
    End If
    
    If strTag = "��Ʒ��" Then
        If InStr("^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Exit Sub
    End If
    
    If strTag = "���" Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Sel����(ByVal strKey As String)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:ѡ�����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim vRect  As RECT, lngH As Long
    Dim objTxt As Object
    Dim blnCancel As Boolean
    
    Dim strTemp As String
    Set objTxt = txtEdit(GetTxtIdx("������"))
    
    strTemp = strKey
    
    strTemp = GetMatchingSting(strTemp)
    If strKey = "" Then
        gstrSQL = "Select Rownum as ID, ����,����,���� From ���������� Order By ���� "
    Else
        gstrSQL = "Select Rownum as ID,����,����,���� From ���������� where ���� Like [1]  Or ���� Like [1] Or ���� Like [1] Order By ���� "
    End If
    
    vRect = zlControl.GetControlRect(objTxt.hwnd)
    lngH = objTxt.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����ѡ����", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strTemp)
   
   '     frmParent=��ʾ�ĸ�����
   '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
   '     bytStyle=ѡ�������
   '       Ϊ0ʱ:�б���:ID,��
   '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
   '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
   '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
   '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
   '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
   '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
   '             bytStyle=1ʱ,�����Ǳ��������
   '     strNote=ѡ������˵������
   '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
   '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
   '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
   '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
   '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
   '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    If blnCancel Then
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    If rsTemp Is Nothing Then
        If mstr���� <> strKey And strKey <> "" Then
                If Asc(strKey) > 0 Then
                    MsgBox "û���ҵ�ƥ��������̣����������룡", vbInformation, gstrSysName
                    If objTxt.Enabled Then objTxt.SetFocus
                    mstr���� = ""
                    Exit Sub
                End If
        
                If MsgBox("û���ҵ���ص������̣����Ӹ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If objTxt.Enabled Then objTxt.SetFocus
                    mstr���� = ""
                Else
                    If zlSureManufacturer = False Then
                        MsgBox "�����̵ı��볬�����޷��Զ����ӡ�" & vbCrLf & "�������ѡ�����еĲ��������̣�", vbInformation, gstrSysName
                        objTxt.Text = "": mstr���� = "": Exit Sub
                    Else
                        Dim str���� As String, str���� As String
                        str���� = strKey
                        If AutoAdd������(str����, str����, Me.Caption) = False Then
                            mstr���� = ""
                            If objTxt.Enabled Then objTxt.SetFocus
                            Exit Sub
                        Else
                            mstr���� = strKey
                        End If
                        Call OS.PressKey(vbKeyTab): Exit Sub
                    End If
                End If
        End If
        Exit Sub
    End If
    objTxt.Text = zlStr.nvl(rsTemp!����)
    If objTxt.Enabled Then objTxt.SetFocus
    Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtEDIT_LostFocus(Index As Integer)
    Dim cur�۸� As Double
    Dim strTag   As String
    Dim dbl�ӳ��� As Double
    Dim dbl����� As Double
    strTag = txtEdit(Index).Tag
    Select Case strTag
        Case "ָ������", "�����"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_�ɱ���)
        Case "ָ���ۼ�", "��ǰ�ۼ�"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_���ۼ�)
            
            If strTag = "��ǰ�ۼ�" Then
                If Val(txtEdit(16).Text) <> 0 Then
                    txtEdit(11).Text = txtEdit(16).Text
                End If
                '������Щ�����ż���ӳ���  15-�ɱ���,11-ָ���ۼ�,16-���ۼ�,14-�������,13-�ӳ���,12-ָ�������
                If Val(Trim(txtEdit(15).Text)) > 0 And Val(Trim(txtEdit(11).Text)) > 0 And Val(Trim(txtEdit(16).Text)) > 0 And Val(Trim(txtEdit(16).Text)) <= Val(Trim(txtEdit(11).Text)) And Val(Trim(txtEdit(14).Text)) / 100 <> 0 Then
                    If Val(Trim(txtEdit(14).Text)) / 100 = 1 Then
                        dbl�ӳ��� = Val(Trim(txtEdit(16).Text)) / Val(Trim(txtEdit(15).Text)) - 1
                    Else
                        dbl�ӳ��� = ((Val(Trim(txtEdit(16).Text)) - Val(Trim(txtEdit(11).Text)) * (1 - Val(Trim(txtEdit(14).Text)))) / Val(Trim(txtEdit(14).Text))) / Val(Trim(txtEdit(15).Text)) - 1
                    End If
                    
                    If dbl�ӳ��� < 0 Then Exit Sub
                    
                    dbl�ӳ��� = dbl�ӳ��� * 100
                    
                    txtEdit(13).Text = Format(dbl�ӳ���, "0.00")
                    
                    'ͨ���ӳ��ʼ���ָ�������
                    dbl����� = dbl�ӳ���
                    Call Calc(dbl�����, False)
                    
                    txtEdit(12).Text = Format(dbl�����, "0.00000")
                End If
            End If
        Case "�������"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBCJL)
        Case "�ɱ��۸�"
            Dim dblSalePrice As Double
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), mFMT.FM_�ɱ���)
            If Val(txtEdit(16).Text) = 0 And Val(txtEdit(Index).Text) <> 0 Then
                '��ǰ�ۼ�Ϊ��ʱ,�������ۼ�
                '�ɱ���*(1+�ӳ���)
                dblSalePrice = Val(txtEdit(Index).Text) * (1 + Val(Me.txtEdit(13).Text) / 100)  '
                '�ɱ���*��1+�ӳ��ʣ�+(ָ���ۼ� -���ɱ���*��1+�ӳ��ʣ���)*(1-���������)
                dblSalePrice = dblSalePrice + (Val(Me.txtEdit(11).Text) - dblSalePrice) * (1 - Val(Me.txtEdit(14)) / 100)
                
                If Val(txtEdit(11).Text) <> 0 Then
                    If dblSalePrice > Val(Me.txtEdit(11).Text) Then
                        '����ָ���ۼ�,��ָ������
                        dblSalePrice = Val(Me.txtEdit(11).Text)
                    End If
                End If
                Me.txtEdit(16).Text = Format(dblSalePrice, mFMT.FM_���ۼ�)
            End If
            
            If Val(txtEdit(15).Text) <> 0 And Val(txtEdit(8).Text) = 0 Then
                txtEdit(8).Text = txtEdit(15).Text
            End If
        Case "�ӳ���"
            '���¼���ָ������ʺͼӳ���
            cur�۸� = Val(txtEdit(13).Text)
            Call Calc(cur�۸�, False)
            
            '�ӳ���
            Me.txtEdit(13).Text = Format(txtEdit(13).Text, GFM_VBJCL)
            'ָ�������
            Me.txtEdit(12).Text = Format(cur�۸�, GFM_VBCJL)
        Case "ָ�������"
            '���¼���ָ������ʺͼӳ���
            
            cur�۸� = Val(txtEdit(12).Text) 'ָ�������
            
            If cur�۸� < 100 Then
                Call Calc(cur�۸�, True)
                'ָ�������
                Me.txtEdit(Index).Text = Format(txtEdit(Index).Text, GFM_VBCJL)
                
                '�ӳ���
                Me.txtEdit(13).Text = Format(cur�۸�, GFM_VBJCL)
            Else
                '���������ָ������ʴ��ڵ���100������������Ҫ�Ӽӳ��ʷ������
                cur�۸� = Val(txtEdit(13).Text)
                Call Calc(cur�۸�, False)
                Me.txtEdit(Index).Text = Format(cur�۸�, GFM_VBCJL)
            End If
        Case "�ɹ�����"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBKL)
        Case "��ֵ˰��"
            txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), GFM_VBJCL)
    End Select
    
    '�ر����뷨
    ImeLanguage False
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
    Dim strTag As String
    strTag = txtEdit(Index).Tag
    zlControl.TxtSelAll txtEdit(Index)
    OS.OpenIme True
    Select Case strTag
        Case "����", "���", "������", "��Ʒ��"
            '�����뷨
            ImeLanguage True
        Case "��ʶ����"
            OS.OpenIme False
    End Select
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            OS.PressKey vbKeyTab
        End If
End Sub

Private Sub cbo��λ_GotFocus(Index As Integer)
    Me.cbo��λ(Index).SelStart = 0: Me.cbo��λ(Index).SelLength = 100
    ImeLanguage True
End Sub

Private Sub cbo��λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
           Exit Sub
        Case Else
            zlControl.TxtCheckKeyPress cbo��λ(Index), KeyAscii, m�ı�ʽ
        End Select
        Exit Sub
    End If
End Sub

Private Sub cbo��λ_LostFocus(Index As Integer)
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    ImeLanguage False
    strTmp = cbo��λ(Index).Text
    blnAdd = True
    For i = 0 To cbo��λ(Index).ListCount - 1
        If cbo��λ(Index).List(i) = Trim(strTmp) Then
            blnAdd = False
            Exit For
        End If
    Next
    If blnAdd And strTmp <> "" Then
        cbo��λ(Index).AddItem strTmp
    End If
    If Index <> 0 Then Exit Sub
    Me.lbl���۵�λ(0).Caption = "Ԫ/" & cbo��λ(Index).Text
    Me.lbl���۵�λ(1).Caption = "Ԫ/" & cbo��λ(Index).Text

End Sub


Private Sub cbo��λ_Change(Index As Integer)
    If mintUnit = 0 Then
        If Index = 1 Then Exit Sub
    Else
        If Index = 0 Then Exit Sub
    End If
    
    Me.lbl���۵�λ(0).Caption = "Ԫ/" & cbo��λ(Index).Text
    Me.lbl���۵�λ(1).Caption = "Ԫ/" & cbo��λ(Index).Text
End Sub


Private Sub stbSpec_Click(PreviousTab As Integer)
    If Me.msf����.Visible Then stbSpec.Tab = 0: Me.msf����.SetFocus: Exit Sub
    
    Select Case stbSpec.Tab
    Case 0
        If Me.txtEdit(0).Enabled Then Me.txtEdit(0).SetFocus
    Case 1
        If Me.txtEdit(8).Enabled Then Me.txtEdit(8).SetFocus
        If Me.cbo�۸�����.Enabled Then Me.cbo�۸�����.SetFocus
    End Select
End Sub

Private Function zlSureManufacturer() As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж��Ƿ�ɼ������������̣������̱����ֶο��Ϊ:10��
    '-------------------------------------------------------------
    Dim strTemp  As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    zlSureManufacturer = False
    gstrSQL = "Select Max(����) ���� From ����������"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        If .EOF Then zlSureManufacturer = True: Exit Function
        If IsNull(!����) Then zlSureManufacturer = True: Exit Function
        
        '����������˳�
        strTemp = .Fields(0).Value
        mintCount = Len(strTemp)
        strTemp = strTemp + 1
        If Len(strTemp) > 10 Then Exit Function
        If mintCount >= Len(strTemp) Then
            strTemp = String(mintCount - Len(strTemp), "0") & strTemp
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
  
Private Sub txt��ѡ��_Change()
    mblnChange = True
End Sub

Private Sub txt��ѡ��_GotFocus()
    Call OS.OpenIme(False)
    zlControl.TxtSelAll txt��ѡ��
End Sub

Private Sub txt��ѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmbStationNo.Visible = True Then
        OS.PressKey vbKeyTab
    Else
        stbSpec.Tab = 1
        If cbo�۸�����.Enabled Then cbo�۸�����.SetFocus
    End If
End Sub

Private Sub txt��ѡ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��ѡ��, KeyAscii, m�ı�ʽ
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
Private Sub txtע���̱�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub


Private Sub txtע���̱�_LostFocus()
    Call OS.OpenIme(False)
End Sub
Private Sub txtע���̱�_GotFocus()
    Me.txtע���̱�.SelStart = 0: Me.txtע���̱�.SelLength = 100
    Call OS.OpenIme(True)
End Sub
Private Sub SaveReg()
    '����:������ص�ע����Ϣ
    Dim strReg As String
    Call zlDatabase.SetPara("�ϴ�ָ�������", Val(Me.txtEdit(12)), glngSys, mlngModule)
    Call zlDatabase.SetPara("�ϴμӳ���", Val(Me.txtEdit(13)), glngSys, mlngModule)
End Sub
Private Sub LoadReg()
    '����:����ע����Ϣֵ
    Dim strReg As String
    Dim blnHavePriv As Boolean
    blnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "��������") And zlStr.IsHavePrivs(mstrPrivs, "ָ���۸����")
    
    strReg = zlDatabase.GetPara("�ϴ�ָ�������", glngSys, mlngModule)
    txtEdit(12).Text = Format(IIf(Val(strReg) = 0, 13.0435, Val(strReg)), GFM_VBCJL)
    strReg = zlDatabase.GetPara("�ϴμӳ���", glngSys, mlngModule)
    txtEdit(13).Text = Format(IIf(Val(strReg) = 0, 15, Val(strReg)), GFM_VBCJL)
End Sub


Private Sub txtע��֤��_Change()
    mblnChange = True
End Sub

Private Sub txtע��֤��_GotFocus()
    zlControl.TxtSelAll txtע��֤��
    OS.OpenIme False
End Sub

Private Sub txtע��֤��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtע��֤��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtע��֤��, KeyAscii, m�ı�ʽ
End Sub

