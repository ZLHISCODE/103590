VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   Icon            =   "frmProSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6495
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   120
      Top             =   6720
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5275
      TabIndex        =   121
      Top             =   6720
      Width           =   1100
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "��������(&1)"
      TabPicture(0)   =   "frmProSetup.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��ʾģʽ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frm����ͼƬ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmRect"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra����ˢ��"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "������������(&2)"
      TabPicture(1)   =   "frmProSetup.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraҩ��"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "�Ŷ���������(&3)"
      TabPicture(2)   =   "frmProSetup.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra����ҩ"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "��ʾ��������(&4)"
      TabPicture(3)   =   "frmProSetup.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frm��ʾ"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fra��ʾʱ��"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame fra����ˢ�� 
         Caption         =   " ����ˢ�� "
         Height          =   1215
         Left            =   120
         TabIndex        =   116
         Top             =   4200
         Width           =   6015
         Begin VB.TextBox txt������ʾʱ�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1440
            TabIndex        =   122
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt������ѯʱ�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1440
            TabIndex        =   118
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label lbl������ѯʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "������ʾʱ��"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   124
            Top             =   765
            Width           =   1080
         End
         Begin VB.Label lbl������ѯʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "��(��Χ��1-60)"
            Height          =   180
            Index           =   2
            Left            =   2760
            TabIndex        =   123
            Top             =   765
            Width           =   1260
         End
         Begin VB.Label lbl������ѯʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "��(��Χ��1-60)"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   119
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lbl������ѯʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "������ѯʱ��"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   117
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   96
         Top             =   3300
         Width           =   6015
         Begin VB.TextBox txt�ѹ���_���� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   4200
            TabIndex        =   114
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt�ѹ���_�и� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2280
            TabIndex        =   112
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt�ѹ���_�п� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            TabIndex        =   110
            Top             =   2235
            Width           =   855
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   10
            Left            =   240
            TabIndex        =   108
            Top             =   1800
            Width           =   5595
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   9
            Left            =   240
            TabIndex        =   103
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd�ѹ�����ɫ 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   107
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd�ѹ������� 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   105
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt�ѹ���_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   102
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt�ѹ���_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   100
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk��ʾ�ѹ��� 
            Caption         =   "��ʾ�ѹ���"
            Height          =   180
            Left            =   240
            TabIndex        =   97
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   34
            Left            =   3720
            TabIndex        =   115
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "�иߣ�"
            Height          =   180
            Index           =   33
            Left            =   1800
            TabIndex        =   113
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "�п�"
            Height          =   180
            Index           =   32
            Left            =   240
            TabIndex        =   111
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "���"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   31
            Left            =   240
            TabIndex        =   109
            Top             =   1920
            Width           =   360
         End
         Begin VB.Shape shp�ѹ�����ɫ 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl�ѹ������� 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   106
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   30
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   29
            Left            =   2520
            TabIndex        =   101
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   28
            Left            =   240
            TabIndex        =   99
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   27
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame fra����ҩ 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   27
         Top             =   420
         Width           =   6015
         Begin VB.TextBox txt����ҩ_���� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   4200
            TabIndex        =   45
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ_�и� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2280
            TabIndex        =   43
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ_�п� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            TabIndex        =   41
            Top             =   2235
            Width           =   855
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   7
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   4755
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   6
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd����ҩ��ɫ 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   38
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd����ҩ���� 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   33
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt����ҩ_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   31
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk��ʾ����ҩ 
            Caption         =   "��ʾ����ҩ"
            Height          =   180
            Left            =   240
            TabIndex        =   28
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   26
            Left            =   3720
            TabIndex        =   46
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "�иߣ�"
            Height          =   180
            Index           =   25
            Left            =   1800
            TabIndex        =   44
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "�п�"
            Height          =   180
            Index           =   24
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "���"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   23
            Left            =   240
            TabIndex        =   40
            Top             =   1920
            Width           =   360
         End
         Begin VB.Shape shp����ҩ��ɫ 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   37
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   22
            Left            =   240
            TabIndex        =   35
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   21
            Left            =   2520
            TabIndex        =   32
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   20
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   19
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame fra��ʾʱ�� 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   84
         Top             =   3180
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   5
            Left            =   240
            TabIndex        =   91
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmdʱ����ɫ 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   95
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdʱ������ 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   93
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox chk��ʾʱ�� 
            Caption         =   "��ʾʱ��"
            Height          =   180
            Left            =   240
            TabIndex        =   85
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txtʱ��_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   88
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txtʱ��_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   90
            Top             =   555
            Width           =   1695
         End
         Begin VB.Shape shpʱ����ɫ 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lblʱ������ 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1320
            TabIndex        =   94
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   17
            Left            =   240
            TabIndex        =   92
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   16
            Left            =   2520
            TabIndex        =   89
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   15
            Left            =   240
            TabIndex        =   87
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   14
            Left            =   240
            TabIndex        =   86
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame frm��ʾ 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   8
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   4755
         End
         Begin VB.TextBox txt��ʾ_���� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   4695
         End
         Begin VB.TextBox txt��ʾ_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   8
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt��ʾ_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   7
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk��ʾ��ʾ 
            Caption         =   "��ʾ��ʾ"
            Height          =   180
            Left            =   240
            TabIndex        =   2
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmd��ʾ���� 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd��ʾ��ɫ 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   12
            Top             =   1320
            Width           =   975
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   4
            Left            =   240
            TabIndex        =   9
            Top             =   960
            Width           =   5595
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   18
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   13
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   12
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   11
            Left            =   2520
            TabIndex        =   4
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   10
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ʾ���� 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   11
            Top             =   1410
            Width           =   990
         End
         Begin VB.Shape shp��ʾ��ɫ 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
      End
      Begin VB.Frame fraҩ�� 
         Caption         =   " ҩ�� "
         Height          =   1935
         Left            =   -74880
         TabIndex        =   16
         Top             =   420
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmdҩ����ɫ 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   26
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdҩ������ 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtҩ��_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   21
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txtҩ��_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   19
            Top             =   555
            Width           =   1695
         End
         Begin VB.Shape shpҩ����ɫ 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lblҩ������ 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   25
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   5
            Left            =   2520
            TabIndex        =   20
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame frm�������� 
         Caption         =   " ����"
         Height          =   3915
         Left            =   -74880
         TabIndex        =   54
         Top             =   2460
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   840
            Index           =   3
            Left            =   240
            TabIndex        =   70
            Top             =   2880
            Width           =   5595
            Begin VB.CheckBox chk���д��ڵ������� 
               Caption         =   "���ڵ�������"
               Height          =   180
               Left            =   240
               TabIndex        =   71
               Top             =   0
               Width           =   1455
            End
            Begin VB.CommandButton cmd������ɫ_���� 
               Caption         =   "������ɫ"
               Height          =   350
               Left            =   3960
               TabIndex        =   74
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmd��������_���� 
               Caption         =   "��������"
               Height          =   350
               Left            =   240
               TabIndex        =   72
               Top             =   360
               Width           =   975
            End
            Begin VB.Shape shp������ɫ_���� 
               BackColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   300
               Left            =   5040
               Top             =   390
               Width           =   375
            End
            Begin VB.Label lbl��������_���� 
               AutoSize        =   -1  'True
               Caption         =   "΢���ź�;12"
               Height          =   180
               Left            =   1350
               TabIndex        =   73
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.Frame fraLine 
            Height          =   960
            Index           =   2
            Left            =   240
            TabIndex        =   65
            Top             =   1800
            Width           =   5595
            Begin VB.CheckBox chk���������������� 
               Caption         =   "������������"
               Height          =   180
               Left            =   240
               TabIndex        =   66
               Top             =   0
               Width           =   1455
            End
            Begin VB.CommandButton cmd������ɫ_���� 
               Caption         =   "������ɫ"
               Height          =   350
               Left            =   3960
               TabIndex        =   69
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmd��������_���� 
               Caption         =   "��������"
               Height          =   350
               Left            =   240
               TabIndex        =   67
               Top             =   360
               Width           =   975
            End
            Begin VB.Shape shp������ɫ_���� 
               BackColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   300
               Left            =   5040
               Top             =   390
               Width           =   375
            End
            Begin VB.Label lbl��������_���� 
               AutoSize        =   -1  'True
               Caption         =   "΢���ź�;12"
               Height          =   180
               Left            =   1350
               TabIndex        =   68
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   1
            Left            =   240
            TabIndex        =   60
            Top             =   960
            Width           =   5595
         End
         Begin VB.TextBox txt����_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   59
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt����_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   57
            Top             =   555
            Width           =   1695
         End
         Begin VB.CommandButton cmd������ɫ_ͨ�� 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   4200
            TabIndex        =   64
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd��������_ͨ�� 
            Caption         =   "��������"
            Height          =   350
            Left            =   240
            TabIndex        =   62
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   9
            Left            =   240
            TabIndex        =   61
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   8
            Left            =   2520
            TabIndex        =   58
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   360
         End
         Begin VB.Shape shp������ɫ_ͨ�� 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl��������_ͨ�� 
            AutoSize        =   -1  'True
            Caption         =   "΢���ź�;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   63
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl����_λ�� 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame frmRect 
         Caption         =   " Һ����λ�ã��ֱ���Ϊ��λ��"
         Height          =   1150
         Left            =   120
         TabIndex        =   75
         Top             =   2940
         Width           =   6015
         Begin VB.TextBox txtҺ����_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   76
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox txtҺ����_�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   3720
            TabIndex        =   79
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox txtҺ����_��� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   80
            Top             =   710
            Width           =   1935
         End
         Begin VB.TextBox txtҺ����_�߶� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   3720
            TabIndex        =   82
            Top             =   710
            Width           =   1935
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   3
            Left            =   3360
            TabIndex        =   78
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   285
            TabIndex        =   77
            Top             =   345
            Width           =   360
         End
         Begin VB.Label lbl��ǩ 
            AutoSize        =   -1  'True
            Caption         =   "��ȣ�"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   81
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�߶ȣ�"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   83
            Top             =   750
            Width           =   540
         End
      End
      Begin VB.Frame frm����ͼƬ 
         Caption         =   " ����ͼƬ "
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   2100
         Width           =   6015
         Begin VB.TextBox txtͼƬλ�� 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   300
            Width           =   4695
         End
         Begin VB.CommandButton cmdͼƬλ�� 
            Caption         =   "��"
            Height          =   270
            Left            =   5520
            TabIndex        =   53
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   315
            Width           =   270
         End
         Begin VB.Label lbl����ͼƬ 
            AutoSize        =   -1  'True
            Caption         =   "λ��"
            Height          =   180
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame fra��ʾģʽ 
         Height          =   1575
         Left            =   120
         TabIndex        =   47
         Top             =   420
         Width           =   6015
         Begin VB.CheckBox chk�ര��ģʽ 
            Caption         =   "�ര��ģʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   48
            Top             =   0
            Width           =   1215
         End
         Begin VB.ListBox lst��ҩ���� 
            Appearance      =   0  'Flat
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   1080
            IMEMode         =   3  'DISABLE
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   360
            Width           =   5520
         End
      End
   End
   Begin MSComDlg.CommonDialog cdl���� 
      Left            =   120
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrWins As String      '����Ĵ��ڴ�
Private mstrReg As String       '����ע���·��

Public Sub ShowMe(ByVal frmParent As Form, ByVal strWins As String)
    mstrWins = strWins
    
    Me.Show 1, frmParent
End Sub

Private Sub chk�ര��ģʽ_Click()
    lst��ҩ����.Enabled = chk�ര��ģʽ.Value
End Sub

Private Sub chk���д��ڵ�������_Click()
    cmd��������_����.Enabled = chk���д��ڵ�������.Value
    cmd������ɫ_����.Enabled = chk���д��ڵ�������.Value
End Sub

Private Sub chk����������������_Click()
    cmd��������_����.Enabled = chk����������������.Value
    cmd������ɫ_����.Enabled = chk����������������.Value
End Sub

Private Sub chk��ʾ����ҩ_Click()
    txt����ҩ_��.Enabled = chk��ʾ����ҩ.Value
    txt����ҩ_��.Enabled = chk��ʾ����ҩ.Value
    cmd����ҩ����.Enabled = chk��ʾ����ҩ.Value
    cmd����ҩ��ɫ.Enabled = chk��ʾ����ҩ.Value
    txt����ҩ_�п�.Enabled = chk��ʾ����ҩ.Value
    txt����ҩ_�и�.Enabled = chk��ʾ����ҩ.Value
    txt����ҩ_����.Enabled = chk��ʾ����ҩ.Value
End Sub

Private Sub chk��ʾʱ��_Click()
    txtʱ��_��.Enabled = chk��ʾʱ��.Value
    txtʱ��_��.Enabled = chk��ʾʱ��.Value
    cmdʱ������.Enabled = chk��ʾʱ��.Value
    cmdʱ����ɫ.Enabled = chk��ʾʱ��.Value
End Sub

Private Sub chk��ʾ��ʾ_Click()
    txt��ʾ_��.Enabled = chk��ʾ��ʾ.Value
    txt��ʾ_��.Enabled = chk��ʾ��ʾ.Value
    cmd��ʾ����.Enabled = chk��ʾ��ʾ.Value
    cmd��ʾ��ɫ.Enabled = chk��ʾ��ʾ.Value
    txt��ʾ_����.Enabled = chk��ʾ��ʾ.Value
End Sub

Private Sub chk��ʾ�ѹ���_Click()
    txt�ѹ���_��.Enabled = chk��ʾ�ѹ���.Value
    txt�ѹ���_��.Enabled = chk��ʾ�ѹ���.Value
    cmd�ѹ�������.Enabled = chk��ʾ�ѹ���.Value
    cmd�ѹ�����ɫ.Enabled = chk��ʾ�ѹ���.Value
    txt�ѹ���_�п�.Enabled = chk��ʾ�ѹ���.Value
    txt�ѹ���_�и�.Enabled = chk��ʾ�ѹ���.Value
    txt�ѹ���_����.Enabled = chk��ʾ�ѹ���.Value
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    '���ܣ�����������õ�ע���
    Dim strWin As String
    Dim i As Integer
    
    SaveSetting "ZLSOFT", mstrReg, "����ģʽ", chk�ര��ģʽ.Value
    
    For i = 0 To Me.lst��ҩ����.ListCount - 1
        If lst��ҩ����.Selected(i) Then
            strWin = strWin & IIf(strWin = "", "", ",") & lst��ҩ����.List(i)
        End If
    Next
    SaveSetting "ZLSOFT", mstrReg, "�ര��", strWin
    
    SaveSetting "ZLSOFT", mstrReg, "ͼƬλ��", txtͼƬλ��.Text
    
    SaveSetting "ZLSOFT", mstrReg, "Һ����_��", txtҺ����_��.Text
    SaveSetting "ZLSOFT", mstrReg, "Һ����_��", txtҺ����_��.Text
    SaveSetting "ZLSOFT", mstrReg, "Һ����_���", txtҺ����_���.Text
    SaveSetting "ZLSOFT", mstrReg, "Һ����_�߶�", txtҺ����_�߶�.Text
    
    SaveSetting "ZLSOFT", mstrReg, "������ѯʱ��", txt������ѯʱ��.Text
    SaveSetting "ZLSOFT", mstrReg, "������ʾʱ��", txt������ʾʱ��.Text
    
    SaveSetting "ZLSOFT", mstrReg, "ҩ��_��", txtҩ��_��.Text
    SaveSetting "ZLSOFT", mstrReg, "ҩ��_��", txtҩ��_��.Text
    SaveSetting "ZLSOFT", mstrReg, "ҩ����ɫ", shpҩ����ɫ.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "����_��", txt����_��.Text
    SaveSetting "ZLSOFT", mstrReg, "����_��", txt����_��.Text
    SaveSetting "ZLSOFT", mstrReg, "����������������", chk����������������.Value
    SaveSetting "ZLSOFT", mstrReg, "���д��ڵ�������", chk���д��ڵ�������.Value
    SaveSetting "ZLSOFT", mstrReg, "������ɫ_ͨ��", shp������ɫ_ͨ��.FillColor
    SaveSetting "ZLSOFT", mstrReg, "������ɫ_����", shp������ɫ_����.FillColor
    SaveSetting "ZLSOFT", mstrReg, "������ɫ_����", shp������ɫ_����.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "��ʾ����ҩ", chk��ʾ����ҩ.Value
    SaveSetting "ZLSOFT", mstrReg, "����ҩ_��", txt����ҩ_��.Text
    SaveSetting "ZLSOFT", mstrReg, "����ҩ_��", txt����ҩ_��.Text
    SaveSetting "ZLSOFT", mstrReg, "����ҩ_�п�", txt����ҩ_�п�.Text
    SaveSetting "ZLSOFT", mstrReg, "����ҩ_�и�", txt����ҩ_�и�.Text
    SaveSetting "ZLSOFT", mstrReg, "����ҩ_����", txt����ҩ_����.Text
    SaveSetting "ZLSOFT", mstrReg, "����ҩ��ɫ", shp����ҩ��ɫ.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "��ʾ�ѹ���", chk��ʾ�ѹ���.Value
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���_��", txt�ѹ���_��.Text
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���_��", txt�ѹ���_��.Text
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���_�п�", txt�ѹ���_�п�.Text
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���_�и�", txt�ѹ���_�и�.Text
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���_����", txt�ѹ���_����.Text
    SaveSetting "ZLSOFT", mstrReg, "�ѹ�����ɫ", shp�ѹ�����ɫ.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "��ʾ��ʾ", chk��ʾ��ʾ.Value
    SaveSetting "ZLSOFT", mstrReg, "��ʾ_��", txt��ʾ_��.Text
    SaveSetting "ZLSOFT", mstrReg, "��ʾ_��", txt��ʾ_��.Text
    SaveSetting "ZLSOFT", mstrReg, "��ʾ_����", txt��ʾ_����.Text
    SaveSetting "ZLSOFT", mstrReg, "��ʾ��ɫ", shp��ʾ��ɫ.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "��ʾʱ��", chk��ʾʱ��.Value
    SaveSetting "ZLSOFT", mstrReg, "ʱ��_��", txtʱ��_��.Text
    SaveSetting "ZLSOFT", mstrReg, "ʱ��_��", txtʱ��_��.Text
    SaveSetting "ZLSOFT", mstrReg, "ʱ����ɫ", shpʱ����ɫ.FillColor
    
    Unload Me
End Sub

Private Sub cmd����ҩ��ɫ_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp����ҩ��ɫ.FillColor = cdl����.Color
End Sub

Private Sub cmd������ɫ_����_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp������ɫ_����.FillColor = cdl����.Color
End Sub

Private Sub cmd������ɫ_ͨ��_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp������ɫ_ͨ��.FillColor = cdl����.Color
End Sub

Private Sub cmd������ɫ_����_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp������ɫ_����.FillColor = cdl����.Color
End Sub

Private Sub cmd��������_����_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "��������_����", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "���д�������_����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "���д�������_б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "���д�������_�ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "��������_����", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "���д�������_����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "���д�������_б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "���д�������_�ֺ�", cdl����.FontSize
    
    lbl��������_����.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd��������_ͨ��_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "��������_ͨ��", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "����ͨ������_����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "����ͨ������_б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "����ͨ������_�ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "��������_ͨ��", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "����ͨ������_����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "����ͨ������_б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "����ͨ������_�ֺ�", cdl����.FontSize
    
    lbl��������_ͨ��.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd��������_����_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "��������_����", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "������������_����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "������������_б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "������������_�ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "��������_����", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "������������_����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "������������_б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "������������_�ֺ�", cdl����.FontSize
    
    lbl��������_����.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmdʱ����ɫ_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shpʱ����ɫ.FillColor = cdl����.Color
End Sub

Private Sub cmdʱ������_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "ʱ������", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "ʱ�����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "ʱ��б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "ʱ���ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "ʱ������", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "ʱ�����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "ʱ��б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "ʱ���ֺ�", cdl����.FontSize
    
    lblʱ������.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd��ʾ��ɫ_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp��ʾ��ɫ.FillColor = cdl����.Color
End Sub

Private Sub cmd��ʾ����_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "��ʾ����", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "��ʾ����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "��ʾб��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "��ʾ�ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "��ʾ����", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "��ʾ����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "��ʾб��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "��ʾ�ֺ�", cdl����.FontSize
    
    lbl��ʾ����.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmdͼƬλ��_Click()
    With cdl����
        .CancelError = True
        .Filter = "Pictures (*.jpg)|*.jpg"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            'ûѡ���ļ�
            err.Clear
        Else
            txtͼƬλ��.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdҩ����ɫ_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shpҩ����ɫ.FillColor = cdl����.Color
End Sub

Private Sub cmdҩ������_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "ҩ������", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "ҩ������", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "ҩ��б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "ҩ���ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "ҩ������", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "ҩ������", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "ҩ��б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "ҩ���ֺ�", cdl����.FontSize
    
    lblҩ������.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd����ҩ����_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "����ҩ����", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "����ҩ����", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "����ҩб��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "����ҩ�ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "����ҩ����", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "����ҩ����", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "����ҩб��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "����ҩ�ֺ�", cdl����.FontSize
    
    lbl����ҩ����.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd�ѹ�����ɫ_Click()
    cdl����.Color = shpҩ����ɫ.FillColor
    cdl����.ShowColor
    shp�ѹ�����ɫ.FillColor = cdl����.Color
End Sub

Private Sub cmd�ѹ�������_Click()
    On Error GoTo errHandle
    
    cdl����.Flags = cdlCFBoth
    cdl����.CancelError = False  '�ѵ�ȡ������������
    
    cdl����.FontName = GetSetting("ZLSOFT", mstrReg, "�ѹ�������", "΢���ź�")
    cdl����.FontBold = GetSetting("ZLSOFT", mstrReg, "�ѹ��Ŵ���", "False")
    cdl����.FontItalic = GetSetting("ZLSOFT", mstrReg, "�ѹ���б��", "False")
    cdl����.FontSize = GetSetting("ZLSOFT", mstrReg, "�ѹ����ֺ�", "12")
    
    cdl����.ShowFont

    '��������
    SaveSetting "ZLSOFT", mstrReg, "�ѹ�������", cdl����.FontName
    SaveSetting "ZLSOFT", mstrReg, "�ѹ��Ŵ���", cdl����.FontBold
    SaveSetting "ZLSOFT", mstrReg, "�ѹ���б��", cdl����.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "�ѹ����ֺ�", cdl����.FontSize
    
    lbl�ѹ�������.Caption = cdl����.FontName & "," & IIf(cdl����.FontBold, "����,", "") & IIf(cdl����.FontItalic, "б��,", "") & cdl����.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    '·����ʼ��
    mstrReg = "����ģ��\ҩ���Ŷӽк�\Һ������Pro"
    
    '��ʼ����ҩ����
    Call LoadWins
    
    '�ָ���������
    Call LoadLocalSettings
End Sub

Private Sub LoadWins()
    '���ܣ���ʼ����ҩ����
    Dim i As Integer
    
    For i = 0 To UBound(Split(mstrWins, ","))
        Me.lst��ҩ����.AddItem Split(mstrWins, ",")(i)
    Next
End Sub

Private Sub LoadLocalSettings()
    '���ܣ��ָ���������
    Dim strWin As String
    Dim i As Integer
    
    '�ָ�����ģʽ
    chk�ര��ģʽ.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "����ģʽ", "0")) = 1, 1, 0)
    
    '�ָ�ѡ�з�ҩ����
    strWin = GetSetting("ZLSOFT", mstrReg, "�ര��", "")
    
    If strWin <> "" Then
        For i = 0 To Me.lst��ҩ����.ListCount - 1
            If InStr(1, strWin, lst��ҩ����.List(i)) > 0 Then
                lst��ҩ����.Selected(i) = True
            End If
        Next
    End If
    
    lst��ҩ����.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "����ģʽ", "0")) = 1)
    
    '�ָ�����ͼƬ
    txtͼƬλ��.Text = GetSetting("ZLSOFT", mstrReg, "ͼƬλ��", "")
    
    '�ָ�Һ����λ��
    txtҺ����_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "Һ����_��", "0"))
    txtҺ����_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "Һ����_��", "0"))
    txtҺ����_���.Text = Val(GetSetting("ZLSOFT", mstrReg, "Һ����_���", "1024"))
    txtҺ����_�߶�.Text = Val(GetSetting("ZLSOFT", mstrReg, "Һ����_�߶�", "768"))
    
    '�ָ�����ˢ��
    txt������ѯʱ��.Text = Val(GetSetting("ZLSOFT", mstrReg, "������ѯʱ��", "1"))
    If Val(txt������ѯʱ��.Text) < 1 Then
        txt������ѯʱ��.Text = 1
    ElseIf Val(txt������ѯʱ��.Text) > 60 Then
        txt������ѯʱ��.Text = 60
    End If
    
    '�ָ���ʾˢ��
    txt������ʾʱ��.Text = Val(GetSetting("ZLSOFT", mstrReg, "������ʾʱ��", "1"))
    If Val(txt������ʾʱ��.Text) < 1 Then
        txt������ʾʱ��.Text = 1
    ElseIf Val(txt������ʾʱ��.Text) > 60 Then
        txt������ʾʱ��.Text = 60
    End If
    
    '�ָ�ҩ������
    txtҩ��_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "ҩ��_��", "0"))
    txtҩ��_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "ҩ��_��", "0"))
    shpҩ����ɫ.FillColor = GetSetting("ZLSOFT", mstrReg, "ҩ����ɫ", vbBlack)
    
    lblҩ������.Caption = GetSetting("ZLSOFT", mstrReg, "ҩ������", "΢���ź�")
    lblҩ������.Caption = lblҩ������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "ҩ������", "False"), ";����", "")
    lblҩ������.Caption = lblҩ������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "ҩ��б��", "False"), ";б��", "")
    lblҩ������.Caption = lblҩ������.Caption & IIf(lblҩ������.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "ҩ���ֺ�", "12")
    
    '�ָ���������
    txt����_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "����_��", "0"))
    txt����_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "����_��", "0"))
    shp������ɫ_ͨ��.FillColor = GetSetting("ZLSOFT", mstrReg, "������ɫ_ͨ��", vbBlack)
    
    lbl��������_ͨ��.Caption = GetSetting("ZLSOFT", mstrReg, "��������_ͨ��", "΢���ź�")
    lbl��������_ͨ��.Caption = lbl��������_ͨ��.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "����ͨ������_����", "False"), ";����", "")
    lbl��������_ͨ��.Caption = lbl��������_ͨ��.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "����ͨ������_б��", "False"), ";б��", "")
    lbl��������_ͨ��.Caption = lbl��������_ͨ��.Caption & IIf(lbl��������_ͨ��.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "����ͨ������_�ֺ�", "12")
    
    chk����������������.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "����������������", "0")) = 1, 1, 0)
    shp������ɫ_����.FillColor = GetSetting("ZLSOFT", mstrReg, "������ɫ_����", vbBlack)
    
    lbl��������_����.Caption = GetSetting("ZLSOFT", mstrReg, "��������_����", "΢���ź�")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "������������_����", "False"), ";����", "")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "������������_б��", "False"), ";б��", "")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(lbl��������_����.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "������������_�ֺ�", "12")
    
    cmd��������_����.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "����������������", "0")) = 1)
    cmd������ɫ_����.Enabled = cmd��������_����.Enabled
    
    chk���д��ڵ�������.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "���д��ڵ�������", "0")) = 1, 1, 0)
    shp������ɫ_����.FillColor = GetSetting("ZLSOFT", mstrReg, "������ɫ_����", vbBlack)
    
    lbl��������_����.Caption = GetSetting("ZLSOFT", mstrReg, "��������_����", "΢���ź�")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "���д�������_����", "False"), ";����", "")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "���д�������_б��", "False"), ";б��", "")
    lbl��������_����.Caption = lbl��������_����.Caption & IIf(lbl��������_����.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "���д�������_�ֺ�", "12")
    
    cmd��������_����.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "���д��ڵ�������", "0")) = 1)
    cmd������ɫ_����.Enabled = cmd��������_����.Enabled
    
    '�ָ�����ҩ����
    chk��ʾ����ҩ.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "��ʾ����ҩ", "1")) = 1, 1, 0)
    txt����ҩ_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "����ҩ_��", "0"))
    txt����ҩ_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "����ҩ_��", "0"))
    txt����ҩ_�п�.Text = Val(GetSetting("ZLSOFT", mstrReg, "����ҩ_�п�", "800"))
    txt����ҩ_�и�.Text = Val(GetSetting("ZLSOFT", mstrReg, "����ҩ_�и�", "350"))
    txt����ҩ_����.Text = Val(GetSetting("ZLSOFT", mstrReg, "����ҩ_����", "5"))
    shp����ҩ��ɫ.FillColor = GetSetting("ZLSOFT", mstrReg, "����ҩ��ɫ", vbBlack)
    
    lbl����ҩ����.Caption = GetSetting("ZLSOFT", mstrReg, "����ҩ����", "΢���ź�")
    lbl����ҩ����.Caption = lbl����ҩ����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "����ҩ����", "False"), ";����", "")
    lbl����ҩ����.Caption = lbl����ҩ����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "����ҩб��", "False"), ";б��", "")
    lbl����ҩ����.Caption = lbl����ҩ����.Caption & IIf(lbl����ҩ����.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "����ҩ�ֺ�", "12")
    
    txt����ҩ_��.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "��ʾ����ҩ", "1")) = 1)
    txt����ҩ_��.Enabled = txt����ҩ_��.Enabled
    cmd����ҩ����.Enabled = txt����ҩ_��.Enabled
    cmd����ҩ��ɫ.Enabled = txt����ҩ_��.Enabled
    txt����ҩ_�п�.Enabled = txt����ҩ_��.Enabled
    txt����ҩ_�и�.Enabled = txt����ҩ_��.Enabled
    txt����ҩ_����.Enabled = txt����ҩ_��.Enabled
    
    '�ָ��ѹ�������
    chk��ʾ�ѹ���.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "��ʾ�ѹ���", "1")) = 1, 1, 0)
    txt�ѹ���_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "�ѹ���_��", "0"))
    txt�ѹ���_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "�ѹ���_��", "0"))
    txt�ѹ���_�п�.Text = Val(GetSetting("ZLSOFT", mstrReg, "�ѹ���_�п�", "800"))
    txt�ѹ���_�и�.Text = Val(GetSetting("ZLSOFT", mstrReg, "�ѹ���_�и�", "350"))
    txt�ѹ���_����.Text = Val(GetSetting("ZLSOFT", mstrReg, "�ѹ���_����", "5"))
    shp�ѹ�����ɫ.FillColor = GetSetting("ZLSOFT", mstrReg, "�ѹ�����ɫ", vbBlack)
    
    lbl�ѹ�������.Caption = GetSetting("ZLSOFT", mstrReg, "�ѹ�������", "΢���ź�")
    lbl�ѹ�������.Caption = lbl�ѹ�������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "�ѹ��Ŵ���", "False"), ";����", "")
    lbl�ѹ�������.Caption = lbl�ѹ�������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "�ѹ���б��", "False"), ";б��", "")
    lbl�ѹ�������.Caption = lbl�ѹ�������.Caption & IIf(lbl�ѹ�������.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "�ѹ����ֺ�", "12")
    
    txt�ѹ���_��.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "��ʾ�ѹ���", "1")) = 1)
    txt�ѹ���_��.Enabled = txt�ѹ���_��.Enabled
    cmd�ѹ�������.Enabled = txt�ѹ���_��.Enabled
    cmd�ѹ�����ɫ.Enabled = txt�ѹ���_��.Enabled
    txt�ѹ���_�п�.Enabled = txt�ѹ���_��.Enabled
    txt�ѹ���_�и�.Enabled = txt�ѹ���_��.Enabled
    txt�ѹ���_����.Enabled = txt�ѹ���_��.Enabled
    
    '�ָ���ʾ����
    chk��ʾ��ʾ.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "��ʾ��ʾ", "1")) = 1, 1, 0)
    txt��ʾ_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "��ʾ_��", "0"))
    txt��ʾ_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "��ʾ_��", "0"))
    txt��ʾ_����.Text = GetSetting("ZLSOFT", mstrReg, "��ʾ_����", "")
    shp��ʾ��ɫ.FillColor = GetSetting("ZLSOFT", mstrReg, "��ʾ��ɫ", vbBlack)
    
    lbl��ʾ����.Caption = GetSetting("ZLSOFT", mstrReg, "��ʾ����", "΢���ź�")
    lbl��ʾ����.Caption = lbl��ʾ����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "��ʾ����", "False"), ";����", "")
    lbl��ʾ����.Caption = lbl��ʾ����.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "��ʾб��", "False"), ";б��", "")
    lbl��ʾ����.Caption = lbl��ʾ����.Caption & IIf(lbl��ʾ����.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "��ʾ�ֺ�", "12")
    
    txt��ʾ_��.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "��ʾ��ʾ", "1")) = 1)
    txt��ʾ_��.Enabled = txt��ʾ_��.Enabled
    cmd��ʾ����.Enabled = txt��ʾ_��.Enabled
    cmd��ʾ��ɫ.Enabled = txt��ʾ_��.Enabled
    txt��ʾ_����.Enabled = txt��ʾ_��.Enabled
    
    '�ָ�ʱ������
    chk��ʾʱ��.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "��ʾʱ��", "1")) = 1, 1, 0)
    txtʱ��_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "ʱ��_��", "0"))
    txtʱ��_��.Text = Val(GetSetting("ZLSOFT", mstrReg, "ʱ��_��", "0"))
    shpʱ����ɫ.FillColor = GetSetting("ZLSOFT", mstrReg, "ʱ����ɫ", vbBlack)
    
    lblʱ������.Caption = GetSetting("ZLSOFT", mstrReg, "ʱ������", "΢���ź�")
    lblʱ������.Caption = lblʱ������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "ʱ�����", "False"), ";����", "")
    lblʱ������.Caption = lblʱ������.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "ʱ��б��", "False"), ";б��", "")
    lblʱ������.Caption = lblʱ������.Caption & IIf(lblʱ������.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "ʱ���ֺ�", "12")
    
    txtʱ��_��.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "��ʾʱ��", "1")) = 1)
    txtʱ��_��.Enabled = txt��ʾ_��.Enabled
    cmdʱ������.Enabled = txt��ʾ_��.Enabled
    cmdʱ����ɫ.Enabled = txt��ʾ_��.Enabled
    
End Sub

Private Sub txt����ҩ_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����ҩ_��
End Sub

Private Sub txt����ҩ_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����ҩ_�и�_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����ҩ_�и�
End Sub

Private Sub txt����ҩ_�и�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����ҩ_����_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����ҩ_����
End Sub

Private Sub txt����ҩ_����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����ҩ_�п�_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����ҩ_�п�
End Sub

Private Sub txt����ҩ_�п�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����ҩ_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����ҩ_��
End Sub

Private Sub txt����ҩ_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����_��
End Sub

Private Sub txt����_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt����_��
End Sub

Private Sub txt����_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt������ʾʱ��_Change()
    If Val(txt������ʾʱ��.Text) < 1 Then
        txt������ʾʱ��.Text = 1
    ElseIf Val(txt������ʾʱ��.Text) > 60 Then
        txt������ʾʱ��.Text = 60
    End If
End Sub

Private Sub txt������ʾʱ��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt������ʾʱ��
End Sub

Private Sub txt������ʾʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtʱ��_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtʱ��_��
End Sub

Private Sub txtʱ��_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtʱ��_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtʱ��_��
End Sub

Private Sub txtʱ��_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt������ѯʱ��_Change()
    If Val(txt������ѯʱ��.Text) < 1 Then
        txt������ѯʱ��.Text = 1
    ElseIf Val(txt������ѯʱ��.Text) > 60 Then
        txt������ѯʱ��.Text = 60
    End If
End Sub

Private Sub txt������ѯʱ��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt������ѯʱ��
End Sub

Private Sub txt������ѯʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ʾ_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt��ʾ_��
End Sub

Private Sub txt��ʾ_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ʾ_����_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt��ʾ_����
End Sub

Private Sub txt��ʾ_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt��ʾ_��
End Sub

Private Sub txt��ʾ_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҩ��_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҩ��_��
End Sub

Private Sub txtҩ��_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҩ��_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҩ��_��
End Sub

Private Sub txtҩ��_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҺ����_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҺ����_��
End Sub

Private Sub txtҺ����_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҺ����_�߶�_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҺ����_�߶�
End Sub

Private Sub txtҺ����_�߶�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҺ����_���_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҺ����_���
End Sub

Private Sub txtҺ����_���_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҺ����_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtҺ����_��
End Sub

Private Sub txtҺ����_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�ѹ���_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt�ѹ���_��
End Sub

Private Sub txt�ѹ���_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�ѹ���_�и�_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt�ѹ���_�и�
End Sub

Private Sub txt�ѹ���_�и�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�ѹ���_����_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt�ѹ���_����
End Sub

Private Sub txt�ѹ���_����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�ѹ���_�п�_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt�ѹ���_�п�
End Sub

Private Sub txt�ѹ���_�п�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�ѹ���_��_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt�ѹ���_��
End Sub

Private Sub txt�ѹ���_��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

