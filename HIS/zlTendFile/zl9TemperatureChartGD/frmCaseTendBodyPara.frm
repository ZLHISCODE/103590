VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   7275
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8955
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame FraSplit 
      Height          =   30
      Left            =   120
      TabIndex        =   71
      Top             =   6720
      Width           =   8775
   End
   Begin VB.Frame fra 
      Caption         =   "��¼��"
      Height          =   1185
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8745
      Begin VB.ComboBox cboSinger 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":000C
         Left            =   5265
         List            =   "frmCaseTendBodyPara.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2670
      End
      Begin VB.CheckBox chk 
         Caption         =   "��¼��Ԥ������ӡʱǩ������ʾǩ��ͼƬ"
         Height          =   180
         Index           =   10
         Left            =   195
         TabIndex        =   7
         Top             =   885
         Width           =   3645
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����ļ�ҳ�밴�ļ���ʽ˳����"
         Height          =   180
         Index           =   11
         Left            =   3960
         TabIndex        =   8
         Top             =   885
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺ����ͬһʱ����Ҫ��¼��ݻ����ļ�"
         Height          =   180
         Index           =   9
         Left            =   195
         TabIndex        =   5
         Top             =   600
         Width           =   3645
      End
      Begin VB.CheckBox chk 
         Caption         =   "ֻ�ڵ�ǰҳ����ʾ��ҳ���ݣ�������ҳ����ʾ��"
         Height          =   180
         Index           =   13
         Left            =   3960
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox cboNodule 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0010
         Left            =   1380
         List            =   "frmCaseTendBodyPara.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¼����ǩģʽ"
         Height          =   180
         Index           =   12
         Left            =   3945
         TabIndex        =   3
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С��ȱʡ��ʽ"
         Height          =   180
         Index           =   5
         Left            =   195
         TabIndex        =   1
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6285
      TabIndex        =   72
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   73
      Top             =   6855
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "���µ�"
      Height          =   5220
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   8745
      Begin VB.CheckBox chk 
         Caption         =   "���������(����/����)��ʽ¼��"
         Height          =   180
         Index           =   18
         Left            =   5280
         TabIndex        =   79
         Top             =   4920
         Width           =   3465
      End
      Begin VB.ComboBox cboStyle 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0014
         Left            =   1680
         List            =   "frmCaseTendBodyPara.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2800
         Width           =   1635
      End
      Begin VB.ComboBox cboNote 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0018
         Left            =   1680
         List            =   "frmCaseTendBodyPara.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2475
         Width           =   1635
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ����ʾƤ�Խ��"
         Height          =   180
         Index           =   8
         Left            =   195
         TabIndex        =   75
         Top             =   3795
         Width           =   2790
      End
      Begin VB.CheckBox chk 
         Caption         =   "�೦�����Է��ӷ�ĸ��ʾ"
         Height          =   180
         Index           =   15
         Left            =   3300
         TabIndex        =   74
         Top             =   3795
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "ȫ���������¼�롢��ʾ����ʱ��(h)"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   48
         Top             =   4935
         Width           =   3330
      End
      Begin VB.PictureBox PicValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1710
         Left            =   2115
         ScaleHeight     =   1710
         ScaleWidth      =   2175
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   495
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdUnVisible 
            Height          =   315
            Left            =   1695
            Picture         =   "frmCaseTendBodyPara.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "ȡ��"
            Top             =   1350
            Width           =   450
         End
         Begin zl9TemperatureChartGD.ColorPicker usrColor 
            Height          =   2190
            Left            =   0
            TabIndex        =   14
            Top             =   -495
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
            Color           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "���µ��ļ��Ŀ�ʼʱ��"
         Height          =   750
         Left            =   5880
         TabIndex        =   68
         Top             =   3840
         Width           =   2790
         Begin VB.OptionButton opt���µ���ʼʱ�� 
            Caption         =   "��Ժʱ��"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   69
            Top             =   330
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt���µ���ʼʱ�� 
            Caption         =   "���ʱ��"
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   70
            Top             =   330
            Width           =   1125
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "��������Ժ�ɫֱ�����(����Ϊб��)"
         Height          =   180
         Index           =   17
         Left            =   5280
         TabIndex        =   49
         Top             =   4650
         Width           =   3465
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ܡ�������Ŀ��ʾ�������ݣ�������ʾ���죩"
         Height          =   180
         Index           =   6
         Left            =   195
         TabIndex        =   47
         Top             =   4650
         Width           =   4215
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����������ݴ�ӡ���ʱ������ʾ�������ݼ̳�)"
         Height          =   180
         Index           =   4
         Left            =   195
         TabIndex        =   46
         Top             =   4365
         Width           =   4400
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ��ӡҽԺ����"
         Height          =   180
         Index           =   1
         Left            =   3300
         TabIndex        =   44
         Top             =   3510
         Width           =   2550
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ӥ�����µ�����������0��ʼ"
         Height          =   180
         Index           =   16
         Left            =   3300
         TabIndex        =   43
         Top             =   3225
         Width           =   2595
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ӥ�����µ���ʾ��Ժ��Ϣ"
         Height          =   180
         Index           =   5
         Left            =   3300
         TabIndex        =   41
         Top             =   2925
         Width           =   2535
      End
      Begin VB.Frame fra 
         Caption         =   "�����Զ���־"
         Height          =   3660
         Index           =   15
         Left            =   5910
         TabIndex        =   50
         Top             =   165
         Width           =   2775
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   5
            ItemData        =   "frmCaseTendBodyPara.frx":05A6
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05A8
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   2100
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   4
            ItemData        =   "frmCaseTendBodyPara.frx":05AA
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05AC
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1734
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   3
            ItemData        =   "frmCaseTendBodyPara.frx":05AE
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05B0
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1368
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   2
            ItemData        =   "frmCaseTendBodyPara.frx":05B2
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05B4
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1002
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   1
            ItemData        =   "frmCaseTendBodyPara.frx":05B6
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05B8
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   636
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   0
            ItemData        =   "frmCaseTendBodyPara.frx":05BA
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05BC
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   270
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   6
            ItemData        =   "frmCaseTendBodyPara.frx":05BE
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05C0
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   2466
            Width           =   2100
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   7
            ItemData        =   "frmCaseTendBodyPara.frx":05C2
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":05C4
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   2835
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Զ���־��˳���ڵ�������"
            Height          =   180
            Index           =   12
            Left            =   150
            TabIndex        =   67
            Top             =   3285
            Width           =   2505
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ"
            Height          =   180
            Index           =   50
            Left            =   135
            TabIndex        =   61
            Top             =   2160
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   49
            Left            =   135
            TabIndex        =   59
            Top             =   1794
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   48
            Left            =   135
            TabIndex        =   57
            Top             =   1428
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            Height          =   180
            Index           =   46
            Left            =   135
            TabIndex        =   55
            Top             =   1062
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            Height          =   180
            Index           =   45
            Left            =   135
            TabIndex        =   53
            Top             =   675
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ"
            Height          =   180
            Index           =   44
            Left            =   135
            TabIndex        =   51
            Top             =   315
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   63
            Top             =   2526
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   65
            Top             =   2895
            Width           =   360
         End
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   12
         Top             =   240
         Width           =   250
      End
      Begin VB.ComboBox cboOper 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":05C6
         Left            =   4095
         List            =   "frmCaseTendBodyPara.frx":05C8
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   225
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   2565
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "0"
         Top             =   2175
         Width           =   390
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ���ӡʱ�������������(�������ʵ���ʹ����Ч)"
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   45
         Top             =   4080
         Width           =   4770
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   2900
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "18"
         Top             =   885
         Width           =   350
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   4250
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "6"
         Top             =   885
         Width           =   350
      End
      Begin VB.ComboBox cboSplit 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":05CA
         Left            =   2760
         List            =   "frmCaseTendBodyPara.frx":05CC
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1830
         Width           =   900
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ��Ե�����ʾ(����Ϊ˫��)"
         Height          =   180
         Index           =   7
         Left            =   195
         TabIndex        =   42
         Top             =   3510
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ�����ʾ���˵������Ϣ"
         Height          =   180
         Index           =   3
         Left            =   195
         TabIndex        =   40
         Top             =   3225
         Width           =   2790
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   32
         Text            =   "1"
         Top             =   1515
         Width           =   420
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "0"
         Top             =   1185
         Width           =   390
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "0"
         Top             =   885
         Width           =   370
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   11
         Text            =   "14"
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������ע�������ٴ�����ʱ,ֹͣǰһ��������ע"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   18
         Top             =   525
         Width           =   4500
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   6
         Left            =   2050
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(6)"
         BuddyDispid     =   196627
         BuddyIndex      =   6
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   1
         Left            =   2970
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196627
         BuddyIndex      =   1
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2100
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1515
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196627
         BuddyIndex      =   2
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   3
         Left            =   4580
         TabIndex        =   27
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   8
         BuddyControl    =   "txt(4)"
         BuddyDispid     =   196627
         BuddyIndex      =   4
         OrigLeft        =   4580
         OrigTop         =   885
         OrigRight       =   4835
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   2
         Left            =   3230
         TabIndex        =   24
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   18
         BuddyControl    =   "txt(3)"
         BuddyDispid     =   196627
         BuddyIndex      =   3
         OrigLeft        =   3230
         OrigTop         =   885
         OrigRight       =   3485
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   4
         Left            =   2955
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2175
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         BuddyControl    =   "txt(5)"
         BuddyDispid     =   196627
         BuddyIndex      =   5
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���²�����ʾ��ʽ"
         Height          =   180
         Index           =   13
         Left            =   195
         TabIndex        =   77
         Top             =   2865
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��˵����ʾλ��"
         Height          =   180
         Index           =   11
         Left            =   195
         TabIndex        =   39
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ȱʡ��ʽ"
         Height          =   180
         Index           =   10
         Left            =   2580
         TabIndex        =   16
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���±����ʱ�����߹̶����        ��"
         Height          =   180
         Index           =   9
         Left            =   210
         TabIndex        =   36
         Top             =   2205
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҹ���       ����"
         Height          =   180
         Index           =   7
         Left            =   2350
         TabIndex        =   22
         Top             =   945
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����       ��"
         Height          =   180
         Index           =   8
         Left            =   3885
         TabIndex        =   25
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Զ���־��ʱ��֮�����ӷ�    "
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   34
         Top             =   1875
         Width           =   2880
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����¼�볬����ǰ        ��Ļ����¼����"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   31
         Top             =   1560
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���±����ʱ��������ݹ̶�        ��"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   28
         Top             =   1230
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���¿�ʼ��¼ʱ��"
         Height          =   180
         Index           =   31
         Left            =   210
         TabIndex        =   19
         Top             =   945
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   1410
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�������ע    ��,��ɫ"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   10
         Top             =   270
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '��ʼ���µ����
    '------------------------------------------------------------------------------------------------------------------
    cboBody(0).Clear
    cboBody(0).AddItem "0-����ʾ"
    cboBody(0).AddItem "1-��ʾ˵��"
    cboBody(0).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(1).Clear
    cboBody(1).AddItem "0-����ʾ"
    cboBody(1).AddItem "1-��ʾ˵��"
    cboBody(1).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(2).Clear
    cboBody(2).AddItem "0-����ʾ"
    cboBody(2).AddItem "1-��ʾ˵��"
    cboBody(2).AddItem "2-��ʾ˵����ʱ��"
    cboBody(2).AddItem "3-��ʾ˵���Ϳ���"
    cboBody(2).AddItem "4-��ʾ˵��,����,ʱ��"
    
    cboBody(3).Clear
    cboBody(3).AddItem "0-����ʾ"
    cboBody(3).AddItem "1-��ʾ˵��"
    cboBody(3).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(4).Clear
    cboBody(4).AddItem "0-����ʾ"
    cboBody(4).AddItem "1-��ʾ˵��"
    cboBody(4).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(5).Clear
    cboBody(5).AddItem "0-����ʾ"
    cboBody(5).AddItem "1-��ʾ˵��"
    cboBody(5).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(6).Clear
    cboBody(6).AddItem "0-����ʾ"
    cboBody(6).AddItem "1-��ʾ˵��"
    cboBody(6).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(7).Clear
    cboBody(7).AddItem "0-����ʾ"
    cboBody(7).AddItem "1-��ʾ˵��"
    cboBody(7).AddItem "2-��ʾ˵����ʱ��"
    
    cboNodule.Clear
    cboNodule.AddItem "0-������"
    cboNodule.AddItem "1-���º���"
    cboNodule.AddItem "2-��˫����"
    cboNodule.AddItem "3-�Ϻ���"
    
    cboSplit.Clear
    cboSplit.AddItem "����"
    cboSplit.AddItem "��"
    cboSplit.AddItem "����ʾ"
    
    '51338,������,2012-07-06
    cboOper.Clear
    cboOper.AddItem "0-����ʾ"
    cboOper.AddItem "1-��ʾ0"
    cboOper.AddItem "2-��ʾ��������"
    
    '51512,������,2012-07-11
    cboNote.Clear
    cboNote.AddItem "0-��ʾ������"
    cboNote.AddItem "1-��ʾ������"
    cboNote.AddItem "2-����ʾ"
    
    '43588,������,2012-09-13,��Ӽ�¼����ǩģʽ
    cboSinger.Clear
    cboSinger.AddItem "0-Ƹ��ְ��+��ǩȨ��"
    cboSinger.AddItem "1-��ǩȨ��"
    
    cboStyle.Clear
    cboStyle.AddItem "0-��ͷ"
    cboStyle.AddItem "1-����"
    cboStyle.AddItem "2-����+��ͷ"
    cboStyle.AddItem "3-����+����"
    
    intStart = zlDatabase.GetPara("���µ��ļ���ʼʱ��", glngSys, 1255, 1, Array(opt���µ���ʼʱ��(0), opt���µ���ʼʱ��(1)), InStr(mstrPrivs, "����ѡ������") > 0)
    opt���µ���ʼʱ��(intStart).Value = True
    txt(6).Text = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4, Array(txt(6), ud(6), lbl(31)), InStr(mstrPrivs, "����ѡ������") > 0)
    txt(1).Text = zlDatabase.GetPara("���±������", glngSys, 1255, 8, Array(txt(1), ud(1), lbl(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    txt(5).Text = zlDatabase.GetPara("�������߹̶��������", glngSys, 1255, 0, Array(txt(5), ud(4), lbl(9)), InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = zlDatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1:1", Array(cboBody(0), cboBody(1), cboBody(2), cboBody(3), cboBody(4), cboBody(5), cboBody(6), cboBody(7)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    For intLoop = 0 To 7
        If UBound(Split(strTmp, ";")) >= intLoop Then
            cboBody(intLoop).ListIndex = Val(Split(strTmp, ";")(intLoop))
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    strTmp = zlDatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255, "0", Array(cboNodule, lbl(5)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboNodule.ListIndex = Val(strTmp)
    Else
        cboNodule.ListIndex = 0
    End If
    
    strTmp = zlDatabase.GetPara("���±�־�ָ���", glngSys, 1255, "0", Array(cboSplit, lbl(6)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboSplit.ListIndex = Val(strTmp)
    Else
        cboSplit.ListIndex = 0
    End If
    
    '����ҹ���־
    strTmp = zlDatabase.GetPara("����ʱ��ҹ���־", glngSys, 1255, "18;6", Array(lbl(7), txt(3), ud(2), lbl(8), txt(4), ud(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    If UBound(Split(strTmp, ";")) >= 1 Then
        txt(3).Text = Abs(Val(Split(strTmp, ";")(0)))
        txt(4).Text = Abs(Val(Split(strTmp, ";")(1)))
    Else
         txt(3).Text = Abs(Val(strTmp))
    End If
    
    '51338,������,2012-07-06
    '��������ȱʡ��ʽ
    strTmp = zlDatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2", Array(cboOper, lbl(10)), InStr(mstrPrivs, "����ѡ������") > 0)
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboOper.ListIndex = Val(strTmp)
    Else
        cboOper.ListIndex = 0
    End If
    
    '51512,������,2012-07-11
    'δ��˵����ʾλ��
    strTmp = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0", Array(cboNote, lbl(9)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboNote.ListIndex = CInt(Val(strTmp))
    Else
        cboNote.ListIndex = 0
    End If
    
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    strTmp = Val(zlDatabase.GetPara("��¼����ǩģʽ", glngSys, 1255, "0", Array(cboSinger, lbl(12)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
        cboSinger.ListIndex = CInt(Val(strTmp))
    Else
        cboSinger.ListIndex = 0
    End If
    
    '���²�����ʾ��ʽ
    strTmp = Val(zlDatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0", Array(cboStyle, lbl(13)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        cboStyle.ListIndex = CInt(Val(strTmp))
    Else
        cboStyle.ListIndex = 0
    End If
    
    txt(0).Text = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10", Array(txt(0), lbl(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("��ӡҽԺ����", glngSys, 1255, "1", Array(chk(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    '51282,������,2012-08-06,DYEYҪ�������Ŀ����¼�롢��ʾ����Сʱ
    chk(2).Value = Val(zlDatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Value = Val(zlDatabase.GetPara("���µ���ʾ���", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    txt(2).Text = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1", Array(txt(2), lbl(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(4).Value = Val(zlDatabase.GetPara("����������", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(5).Value = Val(zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, "1", Array(chk(5)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(6).Value = Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, "1", Array(chk(6)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(7).Value = Val(zlDatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, "0", Array(chk(7)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(8).Value = Val(zlDatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0", Array(chk(8))))
    chk(9).Value = Val(zlDatabase.GetPara("��Ӧ��ݻ����ļ�", glngSys, 1255, "0", Array(chk(9)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(10).Value = Val(zlDatabase.GetPara("��¼��ǩ������ʾ��ʽ", glngSys, 1255, "0", Array(chk(10)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(11).Value = Val(zlDatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, "0", Array(chk(11)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(12).Value = Val(zlDatabase.GetPara("���±�־��˳��������", glngSys, 1255, "0", Array(chk(12)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(13).Value = Val(zlDatabase.GetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", glngSys, 1255, "0", Array(chk(13)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(14).Value = Val(zlDatabase.GetPara("���µ�����ӡ������", glngSys, 1255, "0", Array(chk(14)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(15).Value = Val(zlDatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, "0", Array(chk(15)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(16).Value = Val(zlDatabase.GetPara("Ӥ�����µ�����������ʾ0", glngSys, 1255, "0", Array(chk(16)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(17).Value = Val(zlDatabase.GetPara("���������䷽ʽ", glngSys, 1255, "0", Array(chk(17)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(18).Value = Val(zlDatabase.GetPara("���������(����/����)��ʽ¼��", glngSys, 1255, "0", Array(chk(18)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(11).Enabled = Not CheckPrintDate
    '51283,������,2012-07-11
    picColor.BackColor = Val(zlDatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "0", , InStr(mstrPrivs, "����ѡ������") > 0))

    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Function CheckPrintDate() As Boolean
'---------------------------------------------------------
'����:'��鲡���Ƿ���ڴ�ӡ����,������ھͲ��������û����ļ�ҳ�����
'---------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    CheckPrintDate = False
    
    strSQL = "Select 1 From  ���˻����ӡ  Where ��ӡҳ�� is not null and   Rownum<2"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "����Ƿ���ڴ�ӡ����")
    If rsTemp.RecordCount > 0 Then
        CheckPrintDate = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intStart As Integer
    Dim strTmp As String
    Dim lngColor As Long
    
    If opt���µ���ʼʱ��(0).Value Then
        intStart = 0
    Else
        intStart = 1
    End If
    
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex & ";" & cboBody(2).ListIndex & ";" & cboBody(3).ListIndex & ";" & cboBody(4).ListIndex & ";" & cboBody(5).ListIndex & ";" & cboBody(6).ListIndex & ";" & cboBody(7).ListIndex
    Call zlDatabase.SetPara("���¿�ʼʱ��", Val(txt(6).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���±������", Val(txt(1).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ����", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�������ע����", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����¼�뻤����������", Val(txt(2).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�ٴ�����ֹͣǰ�α�ע", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("δ��˵����ʾλ��", Val(cboNote.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾ���", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��ӡҽԺ����", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '51282,������,2012-08-06,DYEYҪ�������Ŀ����¼�롢��ʾ����Сʱ
    Call zlDatabase.SetPara("ȫ�������ʾ¼��ʱ��", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����������", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", chk(5).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���ܲ�����ʾ��������", chk(6).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾ��ʽ", chk(7).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾƤ�Խ��", chk(8).Value, glngSys, 1255)
    Call zlDatabase.SetPara("��Ӧ��ݻ����ļ�", chk(9).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��¼��ǩ������ʾ��ʽ", chk(10).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�����ļ�ҳ�����", chk(11).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("С��ȱʡ��ʽ", Val(cboNodule.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���±�־�ָ���", Val(cboSplit.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���±�־��˳��������", chk(12).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ��ļ���ʼʱ��", intStart, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����ʱ��ҹ���־", txt(3).Text & ";" & txt(4).Text, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", chk(13).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ�����ӡ������", chk(14).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�೦������ʾ��ʽ", chk(15).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("Ӥ�����µ�����������ʾ0", chk(16).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�������߹̶��������", Val(txt(5).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���������䷽ʽ", chk(17).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���²�����ʾ��ʽ", Val(cboStyle.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���������(����/����)��ʽ¼��", chk(18).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '51338,������,2012-07-06
    Call zlDatabase.SetPara("��������ȱʡ��ʽ", Val(cboOper.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '51283,������,2012-07-11
    lngColor = picColor.BackColor
    Call zlDatabase.SetPara("����������ʾ��ɫ", lngColor, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    Call zlDatabase.SetPara("��¼����ǩģʽ", Val(cboSinger.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdUnVisible_Click()
    picValue.Visible = False
End Sub



Private Sub picColor_Click()
    picValue.Visible = True
    picValue.ZOrder 0
    picValue.SetFocus
    usrColor.Color = picColor.BackColor
End Sub

Private Sub PicValue_GotFocus()
    usrColor.SetFocus
End Sub

Private Sub usrColor_LostFocus()
   If Not Me.ActiveControl Is usrColor _
        And Not Me.ActiveControl Is picValue _
        And Not Me.ActiveControl Is cmdUnVisible _
    Then picValue.Visible = False
End Sub

Private Sub usrColor_pOK()
    picColor.BackColor = usrColor.Color
    picValue.Visible = False
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

