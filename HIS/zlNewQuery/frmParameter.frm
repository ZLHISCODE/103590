VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParameter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѯ��������"
   ClientHeight    =   7095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6540
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   6570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   76
      Top             =   6570
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   77
      Top             =   6570
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   6360
      Left            =   75
      TabIndex        =   79
      Top             =   90
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1.����"
      TabPicture(0)   =   "frmParameter.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk(4)"
      Tab(0).Control(1)=   "pic(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2.����"
      TabPicture(1)   =   "frmParameter.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pic(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3.�Һ�"
      TabPicture(2)   =   "frmParameter.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pic(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4.�۸�"
      TabPicture(3)   =   "frmParameter.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5.���׹Һ�"
      TabPicture(4)   =   "frmParameter.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "pic(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5895
         Index           =   4
         Left            =   240
         ScaleHeight     =   5895
         ScaleWidth      =   6135
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   360
         Width           =   6135
         Begin VB.Frame Fra����ɫ 
            Caption         =   "����ɫ"
            Height          =   735
            Left            =   120
            TabIndex        =   108
            Top             =   480
            Width           =   5775
            Begin VB.PictureBox picBgColor 
               Height          =   350
               Index           =   1
               Left            =   3720
               ScaleHeight     =   285
               ScaleWidth      =   1275
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.PictureBox picBgColor 
               Height          =   350
               Index           =   0
               Left            =   960
               ScaleHeight     =   285
               ScaleWidth      =   1275
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�±���"
               Height          =   180
               Index           =   11
               Left            =   3000
               TabIndex        =   112
               Top             =   330
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ϱ���"
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   111
               Top             =   330
               Width           =   540
            End
         End
         Begin VB.CommandButton cmdSelFont 
            Caption         =   "��������"
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   120
            Width           =   1100
         End
         Begin VB.TextBox txt���׹Һźű� 
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "���ü��׹Һźű� �˺ű����"
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelReg 
            Caption         =   "��"
            Height          =   300
            Left            =   2280
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   120
            Width           =   300
         End
         Begin VB.CommandButton cmdSelFont 
            Caption         =   "��ʾ��Ϣ����"
            Height          =   300
            Index           =   0
            Left            =   4320
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   120
            Width           =   1300
         End
         Begin VB.Frame Fra 
            Caption         =   "�±���"
            Height          =   1455
            Index           =   18
            Left            =   120
            TabIndex        =   100
            Top             =   4440
            Width           =   5775
            Begin VB.Frame Fra 
               Height          =   1095
               Index           =   19
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   5535
               Begin VB.TextBox txt�±��� 
                  BackColor       =   &H00FFFFFF&
                  Height          =   705
                  Left            =   120
                  MaxLength       =   1500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   103
                  Top             =   240
                  Width           =   5295
               End
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "�ϱ���"
            Height          =   1455
            Index           =   16
            Left            =   120
            TabIndex        =   97
            Top             =   1320
            Width           =   5775
            Begin VB.Frame Fra 
               Height          =   1095
               Index           =   17
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   5535
               Begin VB.TextBox txt�ϱ��� 
                  BackColor       =   &H00FFFFFF&
                  Height          =   705
                  Left            =   120
                  MaxLength       =   1500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   99
                  Top             =   240
                  Width           =   5295
               End
            End
         End
         Begin VB.Frame Fra 
            Height          =   1095
            Index           =   15
            Left            =   240
            TabIndex        =   96
            Top             =   3120
            Width           =   5535
            Begin VB.TextBox txt�Һ���ʾ 
               BackColor       =   &H00FFFFFF&
               Height          =   705
               Left            =   120
               MaxLength       =   1500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   101
               Text            =   "frmParameter.frx":0098
               Top             =   240
               Width           =   5295
            End
         End
         Begin MSComDlg.CommonDialog dlgThis 
            Left            =   6000
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame FraFont 
            Caption         =   "��ʾ��Ϣ"
            Height          =   1455
            Left            =   120
            TabIndex        =   113
            Top             =   2880
            Width           =   5775
         End
         Begin VB.Label lbl���׹Һ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Һ���Ŀ"
            Height          =   180
            Left            =   120
            TabIndex        =   114
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ر���ҳ�ϵ�ҽԺ��Ϣ��ʾ(&G)"
         Height          =   180
         Index           =   4
         Left            =   -74400
         TabIndex        =   12
         Top             =   2520
         Width           =   2730
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5910
         Index           =   3
         Left            =   -74925
         ScaleHeight     =   5910
         ScaleWidth      =   6240
         TabIndex        =   88
         Top             =   360
         Width           =   6240
         Begin VB.Frame Fra 
            Caption         =   "����"
            Height          =   1065
            Index           =   9
            Left            =   0
            TabIndex        =   63
            Top             =   4830
            Width           =   4395
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   6
               Left            =   2070
               MaxLength       =   3
               TabIndex        =   65
               Text            =   "30"
               Top             =   300
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   7
               Left            =   2070
               MaxLength       =   2
               TabIndex        =   68
               Text            =   "10"
               Top             =   630
               Width           =   690
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   300
               Left            =   2775
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   630
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(7)"
               BuddyDispid     =   196626
               BuddyIndex      =   7
               OrigLeft        =   3375
               OrigTop         =   1230
               OrigRight       =   3615
               OrigBottom      =   1530
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDWait 
               Height          =   300
               Left            =   2775
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(6)"
               BuddyDispid     =   196626
               BuddyIndex      =   6
               OrigLeft        =   3375
               OrigTop         =   885
               OrigRight       =   3615
               OrigBottom      =   1185
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�۸��ѯ�������(&4)            ��"
               Height          =   180
               Left            =   300
               TabIndex        =   67
               Top             =   690
               Width           =   2970
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�۸��ѯͣ��ʱ��(&3)            ��"
               Height          =   180
               Left            =   300
               TabIndex        =   64
               Top             =   360
               Width           =   2970
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "��ʾ�շ����"
            Height          =   2445
            Index           =   8
            Left            =   4395
            TabIndex        =   70
            Top             =   90
            Width           =   1830
            Begin VB.ListBox lstShow 
               Height          =   2160
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   71
               Top             =   210
               Width           =   1650
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "��ʾ�շ���Ŀ����"
            Height          =   4710
            Index           =   7
            Left            =   0
            TabIndex        =   59
            Top             =   90
            Width           =   4395
            Begin VB.CommandButton cmdClsAll 
               Caption         =   "ȫ��(&D)"
               Height          =   350
               Left            =   1230
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   4275
               Width           =   1100
            End
            Begin VB.CommandButton cmdSelAll 
               Caption         =   "ȫѡ(&A)"
               Height          =   350
               Left            =   90
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   4275
               Width           =   1100
            End
            Begin MSComctlLib.TreeView tvw 
               Height          =   4035
               Left            =   90
               TabIndex        =   60
               Top             =   210
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   7117
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   494
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Checkboxes      =   -1  'True
               ImageList       =   "ils16"
               Appearance      =   1
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "ȱʡѡ�����"
            Height          =   1635
            Index           =   1
            Left            =   4395
            TabIndex        =   74
            Top             =   4260
            Width           =   1830
            Begin VB.ListBox lstClass 
               Height          =   1320
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   75
               Top             =   225
               Width           =   1665
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "��ʾ������Ŀ"
            Height          =   1680
            Index           =   0
            Left            =   4395
            TabIndex        =   72
            Top             =   2550
            Width           =   1815
            Begin VB.ListBox lstPrice 
               Height          =   1320
               Left            =   75
               Style           =   1  'Checkbox
               TabIndex        =   73
               Top             =   255
               Width           =   1650
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5940
         Index           =   0
         Left            =   -74940
         ScaleHeight     =   5940
         ScaleWidth      =   6255
         TabIndex        =   80
         Top             =   315
         Width           =   6255
         Begin VB.CheckBox chkUnload 
            Caption         =   "�رղ�ѯ�������¼����(&F)"
            Height          =   180
            Left            =   540
            TabIndex        =   92
            Top             =   2490
            Width           =   2595
         End
         Begin VB.CheckBox chkShowWorkTime 
            Caption         =   "���վ���ɲ�ѯ�����ϰ�ʱ��(&W)"
            Height          =   180
            Left            =   540
            TabIndex        =   13
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "�豸����(&S)"
            Height          =   350
            Left            =   4875
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   315
            Width           =   1245
         End
         Begin VB.CommandButton cmdYiBao 
            Caption         =   "ҽ������(&B)"
            Height          =   350
            Left            =   4875
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "0"
            Top             =   390
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   3
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "5"
            Top             =   780
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   9
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "30"
            Top             =   1140
            Width           =   570
         End
         Begin VB.TextBox txt 
            Height          =   1185
            Index           =   2
            Left            =   540
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   4725
            Width           =   5595
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   4
            Left            =   870
            TabIndex        =   90
            Top             =   4005
            Width           =   5295
         End
         Begin VB.CheckBox chkusewww 
            Caption         =   "����ҽԺ��վ(�򿪵���ҳֻ����CTRL+w��ALT+F4�ر�)"
            Height          =   255
            Left            =   525
            TabIndex        =   9
            Top             =   1515
            Width           =   4935
         End
         Begin VB.TextBox txturl 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   11
            Text            =   "www.zlsoft.com"
            Top             =   1830
            Width           =   4845
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   16
            Text            =   "5"
            Top             =   3465
            Width           =   540
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   3
            Left            =   870
            TabIndex        =   81
            Top             =   105
            Width           =   5295
         End
         Begin VB.Frame Fra 
            Height          =   120
            Index           =   2
            Left            =   750
            TabIndex        =   82
            Top             =   3135
            Width           =   5295
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   1
            Left            =   1920
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3450
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txt(1)"
            BuddyDispid     =   196626
            BuddyIndex      =   1
            OrigLeft        =   2340
            OrigTop         =   165
            OrigRight       =   2580
            OrigBottom      =   465
            Max             =   300
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   0
            Left            =   2880
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(0)"
            BuddyDispid     =   196626
            BuddyIndex      =   0
            OrigLeft        =   3300
            OrigTop         =   195
            OrigRight       =   3540
            OrigBottom      =   495
            Max             =   300
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   2
            Left            =   2880
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(3)"
            BuddyDispid     =   196626
            BuddyIndex      =   3
            OrigLeft        =   3285
            OrigTop         =   570
            OrigRight       =   3525
            OrigBottom      =   870
            Max             =   600
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   4
            Left            =   2865
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1140
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txt(9)"
            BuddyDispid     =   196626
            BuddyIndex      =   9
            OrigLeft        =   3285
            OrigTop         =   570
            OrigRight       =   3525
            OrigBottom      =   870
            Max             =   600
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�Զ�������ҳ���(&R)           ��"
            Height          =   180
            Index           =   0
            Left            =   510
            TabIndex        =   0
            Top             =   450
            Width           =   2880
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���վ���ˢ�¼��(&T)           ��"
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Width           =   2880
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�Զ�������Ӽ��(&E)          ����"
            Height          =   180
            Index           =   2
            Left            =   495
            TabIndex        =   6
            Top             =   1200
            Width           =   2970
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������Ϣ"
            Height          =   180
            Index           =   8
            Left            =   135
            TabIndex        =   19
            Top             =   4020
            Width           =   720
         End
         Begin VB.Label lbl 
            Caption         =   $"frmParameter.frx":00AF
            Height          =   390
            Index           =   9
            Left            =   510
            TabIndex        =   20
            Top             =   4290
            Width           =   5610
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ַ:"
            Enabled         =   0   'False
            Height          =   180
            Index           =   3
            Left            =   810
            TabIndex        =   10
            Top             =   1875
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ϣ��ѯ"
            Height          =   180
            Index           =   5
            Left            =   135
            TabIndex        =   89
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���ż��"
            Height          =   180
            Index           =   4
            Left            =   15
            TabIndex        =   14
            Top             =   3150
            Width           =   720
         End
         Begin VB.Label lbl 
            Caption         =   "(ע:��Flash,ʵ�����������ŵ����ʱ��)"
            Height          =   180
            Index           =   7
            Left            =   2520
            TabIndex        =   18
            Top             =   3510
            Width           =   3705
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���ż��(&5)          ��"
            Height          =   180
            Index           =   6
            Left            =   390
            TabIndex        =   15
            Top             =   3510
            Width           =   2070
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5925
         Index           =   1
         Left            =   -74895
         ScaleHeight     =   5925
         ScaleWidth      =   6150
         TabIndex        =   85
         Top             =   390
         Width           =   6150
         Begin VB.Frame Fra 
            Caption         =   "����ʾ��ϸ"
            Height          =   2940
            Index           =   10
            Left            =   4020
            TabIndex        =   26
            Top             =   45
            Width           =   2130
            Begin VB.ListBox lst 
               Height          =   2580
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   27
               Top             =   240
               Width           =   1950
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "����"
            Height          =   2835
            Index           =   6
            Left            =   30
            TabIndex        =   28
            Top             =   3015
            Width           =   6105
            Begin VB.OptionButton opt 
               Caption         =   "���Ǽ�ʱ���ѯ���˷���"
               Height          =   180
               Index           =   1
               Left            =   300
               TabIndex        =   43
               Top             =   2460
               Width           =   2310
            End
            Begin VB.OptionButton opt 
               Caption         =   "������ʱ���ѯ���˷���"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   42
               Top             =   2160
               Value           =   -1  'True
               Width           =   2310
            End
            Begin VB.CheckBox chkExit 
               Caption         =   "�����ڷ��ò�ѯ������ָ���˳���ѯ(&E)"
               Height          =   225
               Left            =   345
               TabIndex        =   41
               ToolTipText     =   "�˳�ָ��""AdminExitQuery""(�����ִ�Сд)"
               Top             =   1755
               Width           =   3585
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   5
               Left            =   2070
               MaxLength       =   5
               TabIndex        =   33
               Text            =   "10"
               Top             =   660
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   4
               Left            =   2070
               MaxLength       =   5
               TabIndex        =   30
               Text            =   "30"
               Top             =   300
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   10
               Left            =   2070
               MaxLength       =   3
               TabIndex        =   36
               Text            =   "0"
               Top             =   990
               Width           =   690
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   8
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   39
               Text            =   "0"
               Top             =   1350
               Width           =   360
            End
            Begin MSComCtl2.UpDown udn 
               Height          =   300
               Index           =   3
               Left            =   1425
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1350
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   30
               BuddyControl    =   "txt(8)"
               BuddyDispid     =   196626
               BuddyIndex      =   8
               OrigLeft        =   3285
               OrigTop         =   570
               OrigRight       =   3525
               OrigBottom      =   870
               Max             =   365
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown udn 
               Height          =   300
               Index           =   5
               Left            =   2775
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   990
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   30
               BuddyControl    =   "txt(10)"
               BuddyDispid     =   196626
               BuddyIndex      =   10
               OrigLeft        =   3300
               OrigTop         =   195
               OrigRight       =   3540
               OrigBottom      =   495
               Max             =   300
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   300
               Left            =   2775
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(4)"
               BuddyDispid     =   196626
               BuddyIndex      =   4
               OrigLeft        =   3375
               OrigTop         =   165
               OrigRight       =   3615
               OrigBottom      =   465
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDown3 
               Height          =   300
               Left            =   2775
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   660
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Value           =   60
               BuddyControl    =   "txt(5)"
               BuddyDispid     =   196626
               BuddyIndex      =   5
               OrigLeft        =   3375
               OrigTop         =   525
               OrigRight       =   3615
               OrigBottom      =   825
               Max             =   600
               Min             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ò�ѯͣ��ʱ��(&1)            ��"
               Height          =   180
               Left            =   300
               TabIndex        =   29
               Top             =   360
               Width           =   2970
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ò�ѯ�������(&2)            ��"
               Height          =   180
               Left            =   300
               TabIndex        =   32
               Top             =   720
               Width           =   2970
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "�Զ����ط��ü��(&U)            ��(ע:�ڷ�����ҳΪ0ʱ��Ч)"
               Height          =   180
               Left            =   300
               TabIndex        =   35
               Top             =   1065
               Width           =   5130
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "�����ǰ       ����������(0������)"
               Height          =   180
               Left            =   315
               TabIndex        =   38
               Top             =   1395
               Width           =   3240
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "�����֤"
            Height          =   2940
            Index           =   5
            Left            =   15
            TabIndex        =   24
            Top             =   45
            Width           =   4005
            Begin VB.ListBox lstID 
               Height          =   2580
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   25
               Top             =   255
               Width           =   3870
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "�������뷨"
            Height          =   1530
            Left            =   -10500
            TabIndex        =   86
            Top             =   90
            Visible         =   0   'False
            Width           =   3915
            Begin VB.ComboBox cmbIME 
               Height          =   300
               Left            =   0
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   0
               Visible         =   0   'False
               Width           =   2730
            End
            Begin VB.Label Label10 
               Caption         =   "     ��ѡ��һ������ϲ�������뷨��ΪĬ�����뷨�������ڿɽ��к���¼���λ���Զ��򿪣�Ȼ�����뿪ʱ�Զ��رա�"
               Height          =   750
               Left            =   525
               TabIndex        =   87
               Top             =   270
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Image Image6 
               Height          =   240
               Left            =   120
               Picture         =   "frmParameter.frx":011C
               Top             =   300
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   5805
         Index           =   2
         Left            =   -74925
         ScaleHeight     =   5805
         ScaleWidth      =   6210
         TabIndex        =   84
         Top             =   420
         Width           =   6210
         Begin VB.Frame Fra 
            Caption         =   "�Һ����"
            Height          =   1635
            Index           =   12
            Left            =   2775
            TabIndex        =   55
            Top             =   2205
            Width           =   3420
            Begin MSComctlLib.ListView LvwClass 
               Height          =   1335
               Left            =   90
               TabIndex        =   56
               Top             =   225
               Width           =   3240
               _ExtentX        =   5715
               _ExtentY        =   2355
               View            =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "�Һŷ�ʽ"
            Height          =   1635
            Index           =   11
            Left            =   45
            TabIndex        =   53
            Top             =   2205
            Width           =   2730
            Begin VB.ListBox lstGh 
               Height          =   1320
               Left            =   105
               Style           =   1  'Checkbox
               TabIndex        =   54
               Top             =   225
               Width           =   2505
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "��ʾ����"
            Height          =   1935
            Index           =   14
            Left            =   45
            TabIndex        =   57
            Top             =   3870
            Width           =   6150
            Begin VB.TextBox TxtDisp 
               Height          =   1620
               Left            =   135
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               Text            =   "frmParameter.frx":29FE
               Top             =   240
               Width           =   5910
            End
         End
         Begin VB.Frame Fra 
            Height          =   2220
            Index           =   13
            Left            =   45
            TabIndex        =   91
            Top             =   -45
            Width           =   6150
            Begin VB.CheckBox chk 
               Caption         =   "������ʾ�����Һŷ��ذ�ť"
               Height          =   255
               Index           =   1
               Left            =   3105
               TabIndex        =   94
               Top             =   1800
               Width           =   2895
            End
            Begin VB.CheckBox chk��� 
               Caption         =   "�����Һ�ʱ����ʾ��Ѻű�"
               Height          =   255
               Left            =   3105
               TabIndex        =   93
               Top             =   1388
               Width           =   2895
            End
            Begin VB.CommandButton cmdSetup 
               Caption         =   "Ʊ�ݴ�ӡ����(&S)"
               Height          =   350
               Left            =   210
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1605
            End
            Begin VB.CheckBox ChkPWDDisp 
               Caption         =   "���￨����������ʾ"
               Height          =   210
               Left            =   255
               TabIndex        =   50
               Top             =   1125
               Width           =   1980
            End
            Begin VB.VScrollBar VSstay 
               Height          =   300
               Left            =   2955
               Min             =   1
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   630
               Value           =   10
               Width           =   255
            End
            Begin VB.VScrollBar VSFresh 
               Height          =   300
               Left            =   2955
               Min             =   1
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   285
               Value           =   10
               Width           =   255
            End
            Begin VB.TextBox TxtFreshTime 
               Height          =   300
               Left            =   1905
               MaxLength       =   4
               TabIndex        =   45
               Text            =   "600"
               Top             =   285
               Width           =   1020
            End
            Begin VB.TextBox TXTPwdDelay 
               Height          =   300
               Left            =   1905
               MaxLength       =   4
               TabIndex        =   48
               Text            =   "60"
               Top             =   630
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һ�ʱ���ɻ��۵�"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   51
               Top             =   1425
               Width           =   1800
            End
            Begin VB.Label LblReshTIme 
               AutoSize        =   -1  'True
               Caption         =   "�ҺŰ���ˢ������"
               Height          =   180
               Left            =   255
               TabIndex        =   44
               Top             =   360
               Width           =   1440
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "�ҺŴ���ɿ���ʱ��"
               Height          =   180
               Left            =   255
               TabIndex        =   47
               Top             =   690
               Width           =   1620
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2340
      Top             =   6555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":2A25
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":2FBF
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":3359
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarFirst As Boolean
Private mstrClass As String
Private mstrPrivs As String

Public Function ShowDialog(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Function

Private Function IsPrivs(ByVal strPriv As String) As Boolean
    IsPrivs = (InStr(";" & mstrPrivs & ";", ";" & strPriv & ";") > 0)
End Function

Private Function GetDownAllKey(objNode As Node, ByRef blnCheck As Boolean) As Boolean
    Dim objChild As Node

    On Error GoTo errHand

    objNode.Checked = blnCheck
    If objNode.Children > 0 Then

        Set objChild = objNode.Child
        Do While Not (objChild Is Nothing)

            If GetDownAllKey(objChild, blnCheck) = False Then GoTo errHand

            Set objChild = objChild.Next
        Loop

    End If

    GetDownAllKey = True

    Exit Function

errHand:

End Function

Private Function SetParentCheck(objNode As Node, ByRef blnCheck As Boolean) As Boolean
    Dim objParent As Node

    On Error GoTo errHand

    If blnCheck = False Then Exit Function

    Set objParent = objNode.Parent


    If Not (objParent Is Nothing) Then

        objParent.Checked = blnCheck

        If SetParentCheck(objParent, blnCheck) = False Then GoTo errHand

    End If

    SetParentCheck = True

    Exit Function

errHand:

End Function

Private Sub Load����()
    Dim strTmp As String
    Dim objNode As Node
    Dim lngLoop As Long



    '��ʾ�շ���Ŀ����,ҩƷ�����ķ���
    gstrSQL = "Select -1 As ID,'����ҩ' As ����,Null+0 As �ϼ�id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,����,Decode(�ϼ�id,Null,-1,�ϼ�id) As �ϼ�id,'K' As PrimaryKey  From ҩƷ��;����  Where ����='����ҩ' Start with �ϼ�id Is Null Connect By Prior ID=�ϼ�id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -2 As ID,'�г�ҩ' As ����,Null+0 As �ϼ�id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,����,Decode(�ϼ�id,Null,-2,�ϼ�id) As �ϼ�id,'K' As PrimaryKey  From ҩƷ��;����  Where ����='�г�ҩ' Start with �ϼ�id Is Null Connect By Prior ID=�ϼ�id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -3 As ID,'�в�ҩ' As ����,Null+0 As �ϼ�id,'K' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,����,Decode(�ϼ�id,Null,-3,�ϼ�id) As �ϼ�id,'K' As PrimaryKey From ҩƷ��;����  Where ����='�в�ҩ' Start with �ϼ�id Is Null Connect By Prior ID=�ϼ�id" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select -4 As ID,'��ҩ��' As ����,Null+0 As �ϼ�id,'P' As PrimaryKey From Dual" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select ID,����,Decode(�ϼ�id,Null,-4,�ϼ�id) As �ϼ�id,'P' As PrimaryKey From �շѷ���Ŀ¼  Start with �ϼ�id Is Null Connect By Prior ID=�ϼ�id"

    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF

            If zlCommFun.Nvl(gRs("�ϼ�id").Value, 0) = 0 Then
                Set objNode = tvw.Nodes.Add(, , gRs("PrimaryKey").Value & gRs("ID").Value, gRs("����").Value, 3, 3)
            Else
                Set objNode = tvw.Nodes.Add(gRs("PrimaryKey").Value & gRs("�ϼ�id").Value, tvwChild, gRs("PrimaryKey").Value & gRs("ID").Value, gRs("����").Value, 1, 1)
            End If
            
            gRs.MoveNext
        Wend
    End If

    tbs.Tab = 3
    DoEvents
    
    Dim blnUnSelect As Boolean
    
    strTmp = zlDatabase.GetPara("������ʾ���շѷ���", glngSys, 1536, "", Array(tvw), IsPrivs("��������"))
    If strTmp <> "" Then
        If Left(strTmp, 1) = "-" Then blnUnSelect = True
        strTmp = "," & strTmp & ","
    End If
    
    For lngLoop = 1 To tvw.Nodes.Count
        Set objNode = tvw.Nodes(lngLoop)
        If strTmp = "" Then
            tvw.Nodes(lngLoop).Checked = True
        Else
            If blnUnSelect Then
                If InStr(strTmp, ",-" & objNode.Key & ",") = 0 Then objNode.Checked = True
                
            Else
                If InStr(strTmp, "," & objNode.Key & ",") > 0 Then objNode.Checked = True
                
            End If
        End If
    Next
    tbs.Tab = 0
End Sub

Private Sub CboClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Check1_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Index = 5 Then

            tbs.Tab = 1
            If cmbIME.Enabled And cmbIME.Visible Then
                cmbIME.SetFocus
            End If
            Exit Sub

        End If
        
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub chkExit_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkExit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub ChkPWDDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkShowWorkTime_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkShowWorkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkUnload_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub chkUnload_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'zyk add 200410
Private Sub chkusewww_Click()
    lbl(3).Enabled = Not lbl(3).Enabled
    txturl.Enabled = Not txturl.Enabled
    cmdOK.Tag = "1"
End Sub

Private Sub chkusewww_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbIME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClsAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To tvw.Nodes.Count
        tvw.Nodes(lngLoop).Checked = False
    Next
    cmdOK.Tag = "1"
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1536)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim v_Class As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    If DecideReg = False Then Exit Sub
        
        
    '�����ʱ�������֤
    '------------------------------------------------------------------------------------------------------------------

    strTmp = ""
    
    With lstID

        For lngLoop = 0 To .ListCount - 1
            
            If .Selected(lngLoop) = True Then
                strTmp = strTmp & "1"
            Else
                strTmp = strTmp & "0"
            End If
            

        Next
    End With
    

    Call SetPara("��ѯ���÷�ʽ", strTmp, IsPrivs("��������"))
    
    Call SetPara("����ʱ������", IIf(opt(0).Value, 0, 1), IsPrivs("��������"))
    Call SetPara("�ҺŲ���ʾ��Ѻű�", IIf(chk���.Value = 1, 1, 0), IsPrivs("��������"))
    '------------------------------------------------------------------------------------------------------------------
    v_Class = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            v_Class = v_Class & ",'" & Left(lst.List(i), 1) & "'"
        End If
    Next
    v_Class = IIf(v_Class <> "", Mid(v_Class, 2), "")
    
    Call SetPara("���ò�����ϸ", v_Class, IsPrivs("��������"))
    
    strTmp = ""
    For lngLoop = 0 To lstPrice.ListCount - 1
        If lstPrice.Selected(lngLoop) Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call SetPara("�۸���ʾ��Ϣ", strTmp, IsPrivs("��������"))
    
    strTmp = ""
    For lngLoop = 0 To lstClass.ListCount - 1
        If lstClass.Selected(lngLoop) Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call SetPara("�۸���ʾ���", strTmp, IsPrivs("��������"))
    
    '-----------------------------------------------------------------------------------------------------------------
    strTmp = ""
    For i = 0 To lstShow.ListCount - 1
        If lstShow.Selected(i) Then
            strTmp = strTmp & ",'" & Left(lstShow.List(i), 1) & "'"
        End If
    Next
    strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")

    Call SetPara("������ʾ���շ����", strTmp, IsPrivs("��������"))
    
    '------------------------------------------------------------------------------------------------------------------
    Dim blnUnSelect As Boolean
    
    strTmp = ""
    For i = 1 To tvw.Nodes.Count
        If tvw.Nodes(i).Checked Then
            strTmp = strTmp & "," & tvw.Nodes(i).Key
        Else
            blnUnSelect = True
        End If
    Next
    If blnUnSelect = False Then
        strTmp = ""
    Else
        strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")
    End If
    
    Dim lngMaxLength As Long
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "Select ����ֵ From zlparameters Where 1=2"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    lngMaxLength = rs.Fields(0).DefinedSize
    
    If Len(strTmp) > lngMaxLength Then
        '������lngMaxLength����ȡû��ѡ�еģ������е�ֵǰ����һ������
        
        strTmp = ""
        For i = 1 To tvw.Nodes.Count
            If tvw.Nodes(i).Checked = False Then
                strTmp = strTmp & ",-" & tvw.Nodes(i).Key
            End If
        Next
        strTmp = IIf(strTmp <> "", Mid(strTmp, 2), "")
        
        If Len(strTmp) > lngMaxLength Then
            MsgBox "ѡ��ķ����������࣬������ѡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
    End If
    Call SetPara("������ʾ���շѷ���", strTmp, IsPrivs("��������"))
    Call SetPara("�Һ�ʱ���ɻ��۵�", chk(0).Value, IsPrivs("��������"))
    
    If chkusewww.Value = False Then
        Call SetPara("ҽԺ��ҳ", "", IsPrivs("��������"))
    Else
        Call SetPara("ҽԺ��ҳ", txturl.Text, IsPrivs("��������"))
    End If
    
    If cmdOK.Tag = "1" Then
        On Error GoTo errHand
        
        cmdOK.Tag = ""
        
        Call SetPara("��沥�ż��", Val(txt(1).Text), IsPrivs("��������"))
        Call SetPara("������ҳ���", Val(txt(0).Text), IsPrivs("��������"))
        Call SetPara("���վ���ˢ�¼��", Val(txt(3).Text), IsPrivs("��������"))
        Call SetPara("������Ϣ", txt(2).Text, IsPrivs("��������"))
        Call SetPara("���ò�ѯͣ��ʱ��", Val(txt(4).Text), IsPrivs("��������"))
        Call SetPara("���ò�ѯ�������", Val(txt(5).Text), IsPrivs("��������"))
        Call SetPara("�۸��ѯͣ��ʱ��", Val(txt(6).Text), IsPrivs("��������"))
        Call SetPara("�۸��ѯ�������", Val(txt(7).Text), IsPrivs("��������"))
        Call SetPara("�������ǰ���������", Val(txt(8).Text), IsPrivs("��������"))
        Call SetPara("����������Ӽ��ʱ��", Val(txt(9).Text), IsPrivs("��������"))
        
        Call SetPara("������ʾ�����Һŷ��ذ�ť", chk(1).Value, IsPrivs("��������"))
        
        If txt(10).Enabled Then
            Call SetPara("���ط��ü��", Val(txt(10).Text), IsPrivs("��������"))
        Else
            Call SetPara("���ط��ü��", 0, IsPrivs("��������"))
        End If
        
        Call SetPara("�ر���ҳ�ϵ�ҽԺ��Ϣ��ʾ", chk(4).Value, IsPrivs("��������"))

        Call SetPara("���վ���ɲ�ѯ�����ϰ�ʱ��", Val(chkShowWorkTime.Value), IsPrivs("��������"))
        
        Call gfrmMain.FrameDefault.RefreshPage
        Call gfrmMain.RefreshParamer(Val(txt(0).Text), Val(txt(9).Text))
    End If
    '-------------------------------------------------------------------
    '���ù���Ѻŵ� Id
    '------------------------------------------------------------------
    Call SetPara("�򵥹Һźű�", txt���׹Һźű�.Tag)
    Call SaveFreeRegist
    Unload Me
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
End Sub

Private Sub cmdSelAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To tvw.Nodes.Count
        tvw.Nodes(lngLoop).Checked = True
    Next
    cmdOK.Tag = "1"
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1111", Me)
End Sub

Private Sub cmdYiBao_Click()
    gclsInsure.InsureSupport
End Sub



Private Sub Form_Activate()
    If mvarFirst = False Then Exit Sub
    mvarFirst = False
    
    Call Load����
    
    If txt(4).Enabled And txt(4).Visible Then
        txt(4).SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim lngLoop As Long
    
'    txt(1).Text = GetInterval
    txt(1).Text = Val(zlDatabase.GetPara("��沥�ż��", glngSys, 1536, "5", Array(txt(1), udn(1)), IsPrivs("��������")))
    
    TxtDisp.Text = zlDatabase.GetPara("��ʾ����ʾ��Ϣ", glngSys, 1536, "", Array(TxtDisp), IsPrivs("��������"))
    chkExit.Value = Val(zlDatabase.GetPara("����ָ���˳���ѯ", glngSys, 1536, "0", Array(chkExit), IsPrivs("��������")))
    chkUnload.Value = Val(zlDatabase.GetPara("�رղ�ѯ�������¼����", glngSys, 1536, "0", Array(chkUnload), IsPrivs("��������")))
    
    On Error Resume Next
    opt(CLng(zlDatabase.GetPara("����ʱ������", glngSys, 1536, "0", Array(opt(0), opt(1)), IsPrivs("��������")))).Value = True
    On Error GoTo 0
    chk(0).Value = Val(zlDatabase.GetPara("�Һ�ʱ���ɻ��۵�", glngSys, 1536, "1", Array(chk(0)), IsPrivs("��������")))
    chk���.Value = Val(zlDatabase.GetPara("�ҺŲ���ʾ��Ѻű�", glngSys, 1536, "0", Array(chk���), IsPrivs("��������")))
        
    txt(0).Text = Val(zlDatabase.GetPara("������ҳ���", glngSys, 1536, 0, Array(txt(0), udn(0)), IsPrivs("��������")))
    txt(3).Text = Val(zlDatabase.GetPara("���վ���ˢ�¼��", glngSys, 1536, 5, Array(txt(3), udn(2)), IsPrivs("��������")))
    txt(2).Text = zlDatabase.GetPara("������Ϣ", glngSys, 1536, "", Array(txt(2)), IsPrivs("��������"))
    txt(4).Text = Val(zlDatabase.GetPara("���ò�ѯͣ��ʱ��", glngSys, 1536, 30, Array(txt(4), UpDown1), IsPrivs("��������")))
    txt(5).Text = Val(zlDatabase.GetPara("���ò�ѯ�������", glngSys, 1536, 10, Array(txt(5), UpDown3), IsPrivs("��������")))
    txt(6).Text = Val(zlDatabase.GetPara("�۸��ѯͣ��ʱ��", glngSys, 1536, 30, Array(txt(6), UDWait), IsPrivs("��������")))
    txt(7).Text = Val(zlDatabase.GetPara("�۸��ѯ�������", glngSys, 1536, 10, Array(txt(7), UpDown2), IsPrivs("��������")))
    
    txt(8).Text = Val(zlDatabase.GetPara("�������ǰ���������", glngSys, 1536, 0, Array(txt(8), udn(3)), IsPrivs("��������")))
    txt(9).Text = Val(zlDatabase.GetPara("����������Ӽ��ʱ��", glngSys, 1536, 30, Array(txt(9), udn(4)), IsPrivs("��������")))
    txt(10).Text = Val(zlDatabase.GetPara("���ط��ü��", glngSys, 1536, 0, Array(txt(10), udn(5)), IsPrivs("��������")))
        
    chk(4).Value = Val(zlDatabase.GetPara("�ر���ҳ�ϵ�ҽԺ��Ϣ��ʾ", glngSys, 1536, 0, Array(chk(4)), IsPrivs("��������")))
    chk(1).Value = Val(zlDatabase.GetPara("������ʾ�����Һŷ��ذ�ť", glngSys, 1536, 0, Array(chk(1)), IsPrivs("��������")))
    
    
    chkShowWorkTime.Value = Val(zlDatabase.GetPara("���վ���ɲ�ѯ�����ϰ�ʱ��", glngSys, 1536, 0, Array(chkShowWorkTime), IsPrivs("��������")))
    
    Dim v_Class As String
    Dim strTmp As String

    strTmp = zlDatabase.GetPara("������ʾ���շ����", glngSys, 1536, "", Array(lstShow), IsPrivs("��������"))
    If strTmp <> "" Then strTmp = "," & strTmp & ","
    
    v_Class = zlDatabase.GetPara("���ò�����ϸ", glngSys, 1536, "", Array(lst), IsPrivs("��������"))
    v_Class = "," & v_Class & ","
    
    Set gRs = zlDatabase.OpenSQLRecord("select ����,����||'-'||���� as ��� from �շ���Ŀ���", Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            lst.AddItem IIf(IsNull(gRs!���), "", gRs!���)
            If InStr(v_Class, ",'" & gRs!���� & "',") > 0 Then lst.Selected(lst.NewIndex) = True
            
            lstShow.AddItem IIf(IsNull(gRs!���), "", gRs!���)

            If strTmp = "" Then
                lstShow.Selected(lstShow.NewIndex) = True
            Else
                If InStr(strTmp, ",'" & gRs!���� & "',") > 0 Then lstShow.Selected(lstShow.NewIndex) = True
            End If
            
            gRs.MoveNext
        Wend
    End If
    
    
    '�����ʱ�������֤,000000000
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = Trim(zlDatabase.GetPara("��ѯ���÷�ʽ", glngSys, 1536, "100000000", Array(lstID), IsPrivs("��������")))
    If strTmp = "000000000" Then strTmp = "100000000"
    
    With lstID
        .Clear
        .AddItem "����ͨ�����￨�����"
        .AddItem "����ͨ������Ų����"
        .AddItem "����ͨ��סԺ�Ų����"
        .AddItem "����ͨ������ID�Ų����"
        .AddItem "����ͨ��ҽ���������"
        .AddItem "����ͨ�����֤�����"
        .AddItem "����ͨ��Ʊ�ݺŲ����"
        .AddItem "����ͨ�����ݺŲ����"
        .AddItem "����ͨ���ɣÿ������"
        
        For lngLoop = 0 To .ListCount - 1
'            If Val(zlDatabase.GetPara(.List(lngLoop), glngSys, 1536, IIf(lngLoop = 0, 1, 0), Array(lstID), IsPrivs("��������"))) = 1 Then
'                .Selected(lngLoop) = True
'            End If
            
            If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
                .Selected(lngLoop) = True
            End If
            
        Next
    End With
    '
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = Trim(zlDatabase.GetPara("�۸���ʾ��Ϣ", glngSys, 1536, "0000011", Array(lstPrice), IsPrivs("��������")))
    If Len(strTmp) = 6 Then strTmp = strTmp & "1"
    
    lstPrice.Clear
    lstPrice.AddItem "��������"
    lstPrice.AddItem "����"
    lstPrice.AddItem "����"
    lstPrice.AddItem "��ʶ����"
    lstPrice.AddItem "��ʶ����"
    lstPrice.AddItem "ָ���ۼ�"
    lstPrice.AddItem "����"
    
    For lngLoop = 0 To lstPrice.ListCount - 1
        
        If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
            lstPrice.Selected(lngLoop) = True
        End If
        
    Next
    
    Dim blnHave As Boolean
    
    '�Һŷ�ʽ��ʼ
    '------------------------------------------------------------------------------------------------------------------
    
    With lstGh
        .Clear
        .AddItem "���￨"
        .AddItem "ҽ����"
        .AddItem "���֤"
        .AddItem "�ɣÿ�"
        
        '�Һ����
        strTmp = "," & zlDatabase.GetPara("�Һ����", glngSys, 1536, "", Array(lstGh), IsPrivs("��������")) & ","
        If strTmp = ",���߶�����," Then strTmp = ",���￨,ҽ����,"
        
        blnHave = False
        For lngLoop = 0 To .ListCount - 1
        
            If InStr(strTmp, "," & .List(lngLoop) & ",") > 0 Then
                .Selected(lngLoop) = True
                blnHave = True
            End If

        Next
        
        If blnHave = False Then
            For lngLoop = 0 To .ListCount - 1
                .Selected(lngLoop) = True
            Next
        End If
    End With
    
    
    strTmp = Trim(zlDatabase.GetPara("�۸���ʾ���", glngSys, 1536, "000000", Array(lstClass), IsPrivs("��������")))
    
    lstClass.Clear
    lstClass.AddItem "ҩ��"
    lstClass.AddItem "����"
    lstClass.AddItem "���"
    lstClass.AddItem "����"
    lstClass.AddItem "����"
    lstClass.AddItem "��������"
    
    blnHave = False
    For lngLoop = 0 To lstClass.ListCount - 1
'        If Val(zlDatabase.GetPara(lstClass.List(lngLoop), glngSys, 1536, 0, Array(lstClass), IsPrivs("��������"))) = 1 Then
'            lstClass.Selected(lngLoop) = True
'            blnHave = True
'        End If

        If Val(Mid(strTmp, lngLoop + 1, 1)) = 1 Then
            lstClass.Selected(lngLoop) = True
            blnHave = True
        End If
        
    Next
    If blnHave = False Then
        For lngLoop = 0 To lstClass.ListCount - 1
            lstClass.Selected(lngLoop) = True
        Next
    End If
    
    cmdOK.Tag = ""
    
    mvarFirst = True
    '�������Һ���Ϣ���г�ʼ��
    LoadRegSelef
    
    Dim wwwurl As String
    
    wwwurl = zlDatabase.GetPara("ҽԺ��ҳ", glngSys, 1536, "", Array(chkusewww, txturl), IsPrivs("��������"))
    If wwwurl <> "" Then
        chkusewww.Value = 1
        lbl(3).Enabled = True
        txturl.Enabled = True
        txturl.Text = wwwurl
    End If
     Call LoadFreeRegist
     Call InitFreeRegist
    
    
    
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstClass_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub lstClass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
    
End Sub

Private Sub lstGh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstPrice_ItemCheck(Item As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub lstPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lstShow_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub lstShow_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub LvwClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        tbs.Tab = 2
        TxtFreshTime.SetFocus
    End If
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    Dim i As Long
    
    tbs.ZOrder 0
    For i = 0 To pic.UBound
        pic(i).Enabled = False
    Next
    pic(tbs.Tab).Enabled = True
    
    Select Case tbs.Tab
        Case 0
            If txt(0).Enabled Then txt(0).SetFocus
        Case 1
            If lstID.Enabled Then lstID.SetFocus
        Case 2
            If TxtFreshTime.Enabled Then TxtFreshTime.SetFocus
        Case 3
            If tvw.Enabled Then tvw.SetFocus
    End Select
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)

    Dim lngLoop As Long
    Dim blnCheck As Boolean

    blnCheck = Node.Checked
    '�¼�
    Call GetDownAllKey(Node, Node.Checked)

    '����
    Call SetParentCheck(Node, Node.Checked)
    cmdOK.Tag = "1"
End Sub


Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "1"
    
    If Index = 0 Then
        udn(5).Enabled = (Val(txt(Index).Text) = 0)
        txt(10).Enabled = (Val(txt(Index).Text) = 0)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
        
   If KeyAscii = 13 Then
        KeyAscii = 0
        
        Select Case Index
        Case 2
            tbs.Tab = 1
            'tbs.Tabs(1).Select = True
            
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    
    Else
        Select Case Index
        Case 2
        Case Else
            If CheckIsInclude(UCase(Chr(KeyAscii)), "������") = True Then KeyAscii = 0
        End Select
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Cancel = False Then
        Select Case Index
        Case 1
            If Val(txt(Index).Text) < 5 Then
                MsgBox "��沥�ŵ�ʱ��������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 300 Then
                MsgBox "��沥�ŵ�ʱ��������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 0
            If Val(txt(Index).Text) > 300 Then
                MsgBox "������ҳ��ʱ��������300���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 3
            If Val(txt(Index).Text) < 5 Then
                MsgBox "������ҳ��ʱ��������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "������ҳ��ʱ��������600���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 4
            If Val(txt(Index).Text) < 5 Then
                MsgBox "���ò�ѯͣ��ʱ������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "���ò�ѯͣ��ʱ������10���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 5
            If Val(txt(Index).Text) < 5 Then
                MsgBox "���ò�ѯ�����������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "���ò�ѯ�����������10���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 6
            If Val(txt(Index).Text) < 5 Then
                MsgBox "�۸��ѯͣ��ʱ������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "�۸��ѯͣ��ʱ������10���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 7
            If Val(txt(Index).Text) < 5 Then
                MsgBox "�۸��ѯ�����������5���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 600 Then
                MsgBox "�۸��ѯ�����������10���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 8
            If Val(txt(Index).Text) < 0 Then
                MsgBox "��������Ϊ������", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 365 Then
                MsgBox "��ѯ������ò��ܳ���365�죡", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        Case 9
            If Val(txt(Index).Text) < 1 Then
                MsgBox "������Ӽ������С��1���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
            
            If Val(txt(Index).Text) > 365 Then
                MsgBox "������Ӽ�����ܴ���600(��10Сʱ)���ӣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                SelAll txt(Index)
                Exit Sub
            End If
        End Select
        
    End If
End Sub


Private Sub TxtDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tbs.Tab = 3
        tvw.SetFocus
        
    End If
End Sub

Private Sub TxtFreshTime_GotFocus()
    SelAll TxtFreshTime
End Sub

Private Sub LoadRegSelef()
    Dim rsTmp As New ADODB.Recordset
    Dim Itmx As ListItem
    Dim i As Integer
    
    
    On Error GoTo ErrHandle
    
    Call ReadRegest                             '��ע���֮�ж�ȡ��ʼ������
    
    '����ϵͳ������ĺ�����ʾ
    gstrSQL = "select ����,���� from ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    i = 1
    LvwClass.ListItems.Clear
    Do While Not rsTmp.EOF
        Set Itmx = LvwClass.ListItems.Add(, "K" + CStr(i), CStr(rsTmp("����")))
        If InStr(mstrClass, CStr(rsTmp("����"))) > 0 Then Itmx.Checked = True
    rsTmp.MoveNext
    i = i + 1
    Loop
    rsTmp.Close
    '��ϵͳ����ĺ�����ʾ
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReadRegest()
    '������������ע���֮�ж�ȡ������ʾ�ڽ���
    If zlDatabase.GetPara("������ˢ������", glngSys, 1536, "", Array(TxtFreshTime, VSFresh), IsPrivs("��������")) = "" Then
        TxtFreshTime.Text = 600
        TXTPwdDelay.Text = 60
'        CboClass.Text = "���￨�Һ�"
        ChkPWDDisp.Value = 0
    Else
        TxtFreshTime.Text = zlDatabase.GetPara("������ˢ������", glngSys, 1536, "", Array(TxtFreshTime, VSFresh), IsPrivs("��������"))
        TXTPwdDelay.Text = zlDatabase.GetPara("������֤����ͣ��ʱ��", glngSys, 1536, "", Array(TXTPwdDelay, VSstay), IsPrivs("��������"))

        mstrClass = zlDatabase.GetPara("�Һŵĺ���", glngSys, 1536, "", Array(LvwClass), IsPrivs("��������"))
        ChkPWDDisp.Value = Val(zlDatabase.GetPara("������ʾ����", glngSys, 1536, 0, Array(ChkPWDDisp), IsPrivs("��������")))
        
    End If
End Sub

Private Sub WriteRegedit()
    Dim i As Long
    Dim strTmp As String
    
    '�������ӽ��泭д����ע���
    mstrClass = ""
    For i = 1 To CLng(LvwClass.ListItems.Count)
        If LvwClass.ListItems("K" + CStr(i)).Checked = True Then
          mstrClass = mstrClass + "'" + LvwClass.ListItems("K" + CStr(i)).Text + "',"
        End If
    Next
    If Trim(mstrClass) <> "" Then mstrClass = Mid(mstrClass, 1, Len(mstrClass) - 1)
    
    strTmp = ""
    With lstGh
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = strTmp & "," & .List(i)
            End If
        Next
    End With
    
    Call SetPara("������ˢ������", Val(TxtFreshTime.Text), IsPrivs("��������"))
    Call SetPara("������֤����ͣ��ʱ��", Val(TXTPwdDelay.Text), IsPrivs("��������"))
    Call SetPara("�Һ����", strTmp, IsPrivs("��������"))
    Call SetPara("�Һŵĺ���", mstrClass, IsPrivs("��������"))
    Call SetPara("������ʾ����", ChkPWDDisp.Value, IsPrivs("��������"))
    Call SetPara("��ʾ����ʾ��Ϣ", TxtDisp.Text, IsPrivs("��������"))
    Call SetPara("����ָ���˳���ѯ", chkExit.Value, IsPrivs("��������"))
    Call SetPara("�رղ�ѯ�������¼����", chkUnload.Value, IsPrivs("��������"))

End Sub

Private Sub TxtFreshTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    If CheckIsInclude(UCase(Chr(KeyAscii)), "������") = True Then KeyAscii = 0
End Sub

Private Sub TxtFreshTime_LostFocus()
    If Not IsNumeric(TxtFreshTime.Text) Then
        MsgBox "�뽫ˢ����������Ϊ������Ϣ", vbInformation, gstrSysName
        tbs.Tab = 2
        TxtFreshTime.SetFocus
    End If
End Sub

Private Sub TXTPwdDelay_GotFocus()
    SelAll TXTPwdDelay
End Sub

Private Sub TXTPwdDelay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    If CheckIsInclude(UCase(Chr(KeyAscii)), "������") = True Then KeyAscii = 0
End Sub

Private Sub TXTPwdDelay_LostFocus()
    If Not IsNumeric(TXTPwdDelay.Text) Then
        MsgBox "�뽫�ҺŴ���ɿ���ʱ������Ϊ������Ϣ", vbInformation, gstrSysName
        tbs.Tab = 2
        TXTPwdDelay.SetFocus
    End If
End Sub
'zyk add 200410
Private Sub txturl_Change()
        cmdOK.Tag = "1"
End Sub

Private Sub VSFresh_Change()
    If Not IsNumeric(TxtFreshTime.Text) Then
        MsgBox "�뽫ˢ��ʱ������Ϊ������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TxtFreshTime.Text < 1 Then TxtFreshTime.Text = 1
    TxtFreshTime.Text = TxtFreshTime.Text + 10 - VSFresh.Value
    VSFresh.Value = 10
End Sub

Private Sub VSstay_Change()
    If Not IsNumeric(TXTPwdDelay.Text) Then
        MsgBox "�뽫ˢ��ʱ������Ϊ������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TXTPwdDelay.Text < 1 Then TXTPwdDelay.Text = 1
    TXTPwdDelay.Text = TXTPwdDelay.Text + 10 - VSstay.Value
    VSstay.Value = 10
End Sub

Private Function DecideReg() As Boolean
    Dim i As Integer

    '�жϹҺŵĺ�����Ϣ
    mstrClass = ""
    For i = 1 To CLng(LvwClass.ListItems.Count)
        If LvwClass.ListItems("K" + CStr(i)).Checked = True Then
          mstrClass = mstrClass + "'" + LvwClass.ListItems("K" + CStr(i)).Text + "',"
        End If
    Next
'    If mstrClass = "" Then
'       MsgBox "������ѡ��һ��Һ���Ŀ", vbInformation, gstrSysName
'       DecideReg = False: Exit Function
'    End If
     '�ж�ˢ��ʱ���
    If Not IsNumeric(TxtFreshTime.Text) Then
         MsgBox "�뽫ˢ��ʱ������Ϊ������Ϣ", vbInformation, gstrSysName
         If TxtFreshTime.Enabled And TxtFreshTime.Visible Then TxtFreshTime.SetFocus
         DecideReg = False: Exit Function
    End If

    If TxtFreshTime.Text < 0 Or TxtFreshTime.Text > 9999 Then
         MsgBox "�뽫ˢ��ʱ������Ϊ0��9999��������Ϣ", vbInformation, gstrSysName
         If TxtFreshTime.Enabled And TxtFreshTime.Visible Then TxtFreshTime.SetFocus
         DecideReg = False: Exit Function
    End If
   '�ж������ӳٴ���
    If Not IsNumeric(TXTPwdDelay.Text) Then
         MsgBox "�뽫������֤������ӳ�ʱ������Ϊ1��9999��������Ϣ", vbInformation, gstrSysName
         If TXTPwdDelay.Enabled And TXTPwdDelay.Visible Then TXTPwdDelay.SetFocus
         DecideReg = False: Exit Function
    End If
    If (TXTPwdDelay.Text > 9999) Or (TXTPwdDelay.Text < 0) Then
         MsgBox "�뽫������֤������ӳ�ʱ������Ϊ0��9999��������Ϣ", vbInformation, gstrSysName
         If TXTPwdDelay.Enabled And TXTPwdDelay.Visible Then TXTPwdDelay.SetFocus
         DecideReg = False: Exit Function
    End If

    On Error GoTo ErrHandle
    
    Call WriteRegedit
    
    DecideReg = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadFreeRegist()
'---------------------------
'�����Ѿ����õ� ���׹Һ����
'---------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "select a.���� as �ű�,'['|| a.����||']'||b.���� as ���� from �ҺŰ��� a,�շ���ĿĿ¼ b  where a.��Ŀid =b.id and  a.����=[1]"
    On Error GoTo hErr
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, GetPara("�򵥹Һźű�", -1, True))
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Me.txt���׹Һźű�.Text = rsTmp!����
    Me.txt���׹Һźű�.Tag = Nvl(rsTmp!�ű�, -1)
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub selFreeRegist()
'------------------------------------------------
'����:���ص�ǰ�ɹҵĹҺ�
'------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    Dim strTime As String
    Dim i As Long
   vRect = GetControlRect(Me.txt���׹Һźű�.hwnd)
   On Error GoTo ErrHandle
            '�����ǰʱ�����ھ����ʱ���
            strTime = _
                  "Select ʱ��� From ʱ��� Where" & _
                  " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                  " Between" & _
                  " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & _
                  " And" & _
                  " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
                  " Or" & _
                  " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                  " Between" & _
                  " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & _
                  " And" & _
                  " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
   
   
            strSQL = "" & _
            "   Select  distinct M.ID as ID,M.���� as �ű�,M.����ID as ����ID,M.���� as ����,M.��ĿID as ��ĿID,C.���� as ����, " & _
            "             N.���� as ����,Nvl(M.ҽ������, ' ') as ҽ������,M.ҽ��id,Decode(To_Char(SysDate,'D'),'1',M.����," & _
            "             '2',M.��һ,'3',M.�ܶ�,'4',M.����,'5',M.����,'6',M.����,'7',M.����)  as ʱ��" & _
            "   From �ҺŰ��� M,�շ���ĿĿ¼ N,���ű� C " & _
            "   Where M.ID not in (  Select  A.ID from �ҺŰ��� A,���˹ҺŻ��� B,�ҺŰ������� C " & _
            "                                   Where  a.����ID = B.����ID And a.��ĿID = B.��ĿID And a.id=c.����ID(+) And " & _
            "                                          Decode(To_Char(Sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) =C.������Ŀ(+) " & vbNewLine & _
            "                                          And Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0)  And   a.����ID = B.����ID And a.��ĿID = B.��ĿID   And " & GetNodeCheckSQL("N.վ��") & " And " & GetNodeCheckSQL("C.վ��") & " And " & _
            "                                               B.����=Trunc(Sysdate)  and c.�޺���<= B.�ѹ��� and C.�޺���<>0 ) " & _
            "               And  Decode(To_Char(SysDate,'D'),'1',M.����,'2',M.��һ,'3',M.�ܶ�,'4', M.����,'5',M.����,'6',M.����,'7',M.����) in (" + strTime + ")  " & _
            "               And M.��ĿID=N.ID  and M.����ID=C.ID   " & _
            "               And M.ͣ������ is NULL And (M.ҽ��id Is Null Or Exists (Select 1 From ��Ա�� y Where y.ID=M.ҽ��id And " & GetNodeCheckSQL("y.վ��") & _
            "               And (y.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.����ʱ�� Is Null)) ) " & _
            "               And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=M.ID and Sysdate between ��ʼֹͣʱ�� and ����ֹͣʱ�� )"
            
            strSQL = "" & _
            "  Select �ű� as ID ,'['||�ű�||']'||���� as ����,����,ҽ������,����,ʱ��,sum(nvl(�۸�,0)) as �۸� " & _
            "  From ( With A1 as (" & strSQL & ") " & _
            "           Select  A1.*,D.�ּ� as �۸�  From A1,�շѼ�Ŀ D " & _
            "           Where A1.��ĿID=D.�շ�ϸĿID And     D.ִ������<=sysdate and (D.��ֹ����> sysdate or D.��ֹ���� is null)  " & _
            "           Union all " & _
            "           Select  A1.*,D.�ּ� as �۸�  From A1,�շѴ�����Ŀ A,�շѼ�Ŀ D " & _
            "           Where A1.��ĿID=A.����ID and A.����ID=D.�շ�ϸĿID  And  D.ִ������<=sysdate and (D.��ֹ����> sysdate or D.��ֹ���� is null)  " & _
            "       )" & _
            " Group by ID,����,�ű�,����ID,��ĿID,����,����,ҽ������,ҽ��id,ʱ��   Having sum(nvl(�۸�,0))=0" & _
              vbNewLine & "  union all  " & vbNewLine & _
             " select '-1' as id ,'[�����ü��׹Һ���Ŀ]' as ����, null as ����,null as ҽ������,null as ����,null as ʱ��,null as �۸� from Dual " & _
             "   Order by ����,�۸�"
            '����ID,��ĿID,ҽ��id,
            Set rsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�Һ���Ŀѡ��", False, _
            Nvl(Me.txt���׹Һźű�.Tag, -1), "", False, False, True, vRect.Left, vRect.Top, Me.txt���׹Һźű�.Height, blnCancel, True, True)
            If blnCancel Then Exit Sub
            If rsInfo Is Nothing Then
                MsgBox "��ǰû�п��õĹҺ���Ŀ���뵽�ҺŰ��������ã�", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
            Me.txt���׹Һźű�.Text = IIf(Nvl(rsInfo!ID, -1) = -1, "", Nvl(rsInfo!����))
            Me.txt���׹Һźű�.Tag = Nvl(rsInfo!ID, -1)
            
            Exit Sub
ErrHandle:
            If ErrCenter() = 1 Then Resume
            SaveErrLog
End Sub


Private Sub InitFreeRegist()
    Dim strFontName As String, strMsg As String, dblColor As Double, dblSize As Double
    Dim dblUpBgColor As Double, dblDownBgColor As Double
    Dim blnBold As Boolean, blnItalic As Boolean
    '��ʾ��Ϣ
    If GetRegistParaFont("�򵥹Һ���ʾ��Ϣ", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With txt�Һ���ʾ
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
'            .Font.Size = dblSize
            .ForeColor = dblColor
            If dblColor = vbWhite Then
              .BackColor = &HE0E0E0
            Else
               .BackColor = vbWhite
            End If
        End With
    End If
    If GetRegistParaFont("�򵥹Һ��ϱ���", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.txt�ϱ���
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
'            .SelStart=0:.SelLength=1:.setfo
            .ForeColor = dblColor
            If dblColor = vbWhite Then
                .BackColor = &HE0E0E0
             Else
               .BackColor = vbWhite
            End If
        End With
    End If
    If GetRegistParaFont("�򵥹Һ��±���", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
      With Me.txt�±���
            .Text = strMsg
            .Font.Name = strFontName
            .Tag = dblSize
            .FontItalic = blnItalic
            .FontBold = blnBold
            '          .Font.Size = dblSize
            .ForeColor = dblColor
            If dblColor = vbWhite Then
              .BackColor = &HE0E0E0
            Else
               .BackColor = vbWhite
            End If
      End With
    End If
    If GetFreeRegistBGColor(dblUpBgColor, dblDownBgColor) Then
       Me.picBgColor(0).BackColor = dblUpBgColor
       Me.picBgColor(1).BackColor = dblDownBgColor
    End If
    
End Sub

Private Function SaveFreeRegist()
    With Me.txt�Һ���ʾ
        SetRegistParaFont "�򵥹Һ���ʾ��Ϣ", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                          .FontBold, .FontItalic
        
    End With
    With Me.txt�ϱ���
       SetRegistParaFont "�򵥹Һ��ϱ���", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                          .FontBold, .FontItalic
    End With
    With Me.txt�±���
       SetRegistParaFont "�򵥹Һ��±���", .Text, .Font.Name, CDbl(Val(.Tag)), CDbl(.ForeColor), _
                         .FontBold, .FontItalic
    End With
    SetFreeRegistBGColor CDbl(picBgColor(0).BackColor), CDbl(Me.picBgColor(1).BackColor)
End Function

Private Sub cmdSelFont_Click(Index As Integer)
    '---------------------
    '���ü��׹Һ� ���������ɫ��
    '----------------------
    With Me.dlgThis
        Select Case Index
    
        Case 0
            .DialogTitle = "���ü��׹Һ���ʾ��Ϣ����"
            .flags = &H2 + &H1 + &H400 + &H800 + &H100  '&H100000 +
            .Color = Me.txt�Һ���ʾ.ForeColor
            .FontBold = txt�Һ���ʾ.Font.Bold
            .FontItalic = txt�Һ���ʾ.FontItalic
            .FontName = txt�Һ���ʾ.Font.Name
            .FontSize = IIf(Val(txt�Һ���ʾ.Tag) > 0, Val(txt�Һ���ʾ.Tag), txt�Һ���ʾ.Font.Size)
             Err.Clear: On Error Resume Next:
             .CancelError = True
            .ShowFont
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
             txt�Һ���ʾ.Tag = Val(.FontSize)
             txt�Һ���ʾ.Font.Name = .FontName
             txt�Һ���ʾ.Font.Bold = .FontBold
             txt�Һ���ʾ.Font.Italic = .FontItalic
             txt�Һ���ʾ.ForeColor = .Color
             If .Color = vbWhite Then
                 txt�Һ���ʾ.BackColor = &HE0E0E0
             Else
                txt�Һ���ʾ.BackColor = vbWhite
             End If
             
        Case 1
             .DialogTitle = "���ü��׹Һű�������"
            .flags = &H2 + &H1 + &H400 + &H800 + &H100
            .Color = Me.txt�ϱ���.ForeColor
            .FontBold = txt�ϱ���.Font.Bold
            .FontItalic = txt�ϱ���.FontItalic
            .FontName = txt�ϱ���.Font.Name
            .FontSize = IIf(Val(txt�ϱ���.Tag) > 0, Val(txt�ϱ���.Tag), txt�ϱ���.Font.Size)
             Err.Clear: On Error Resume Next:
             .CancelError = True
            .ShowFont
            If Err.Number <> 0 Then Err.Clear:  On Error GoTo 0: Exit Sub
             txt�ϱ���.Tag = Val(.FontSize)
             txt�ϱ���.Font.Name = .FontName
             txt�ϱ���.Font.Bold = .FontBold
             txt�ϱ���.Font.Italic = .FontItalic
             txt�ϱ���.ForeColor = .Color
             txt�±���.Tag = Val(.FontSize)
             txt�±���.Font.Name = .FontName
             txt�±���.Font.Bold = .FontBold
             txt�±���.Font.Italic = .FontItalic
             txt�±���.ForeColor = .Color
             If .Color = vbWhite Then
                 txt�ϱ���.BackColor = &HE0E0E0
                 txt�±���.BackColor = &HE0E0E0
             Else
                txt�ϱ���.BackColor = vbWhite
                txt�±���.BackColor = vbWhite
             End If
        End Select
    End With
  
     
End Sub

Private Sub cmdSelReg_Click()
    Call selFreeRegist
End Sub


Private Sub picBgColor_Click(Index As Integer)
        With Me.dlgThis
        
        .DialogTitle = "���ü��׹Һ��ϱ�����ɫ"
        .flags = &H2 + &H1
        .Color = Me.picBgColor(Index).BackColor
        Err.Clear: On Error Resume Next:
        .CancelError = True
        .ShowColor
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
        Me.picBgColor(Index).BackColor = .Color: picBgColor(Index).Tag = 1
'        If Index = 0 Then
'            Me.txt�ϱ���.BackColor = .Color
'        Else
'            Me.txt�±���.BackColor = .Color
'        End If
    End With
End Sub
