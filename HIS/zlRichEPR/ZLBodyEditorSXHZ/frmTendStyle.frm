VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmTendFileStyle 
   Appearance      =   0  'Flat
   Caption         =   "�����ļ���ʽ"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12225
   Icon            =   "frmTendStyle.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12225
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picTable 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   810
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   0
      Top             =   3540
      Width           =   6885
      Begin VB.OptionButton optTabTiers 
         Caption         =   "��(&3)"
         Height          =   180
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Top             =   750
         Width           =   780
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   4380
         TabIndex        =   107
         Top             =   1905
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   4380
         TabIndex        =   105
         Top             =   255
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   1020
         TabIndex        =   102
         Top             =   2355
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   1020
         TabIndex        =   101
         Top             =   225
         Width           =   2235
      End
      Begin VB.TextBox txtHeadText 
         Height          =   885
         Left            =   4980
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2070
         Width           =   1710
      End
      Begin MSComCtl2.UpDown udHeadCol 
         Height          =   300
         Left            =   4335
         TabIndex        =   39
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadCol"
         BuddyDispid     =   196614
         OrigLeft        =   5985
         OrigTop         =   2085
         OrigRight       =   6225
         OrigBottom      =   2370
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeadCol 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "1"
         Top             =   2430
         Width           =   330
      End
      Begin VB.TextBox txtHeadRow 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4035
         MaxLength       =   1
         TabIndex        =   35
         Text            =   "2"
         Top             =   2070
         Width           =   330
      End
      Begin MSComCtl2.UpDown udRecordTo 
         Height          =   285
         Left            =   5730
         TabIndex        =   28
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   8
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordTo"
         BuddyDispid     =   196621
         OrigLeft        =   5985
         OrigTop         =   405
         OrigRight       =   6225
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRecordFrom 
         Height          =   285
         Left            =   4365
         TabIndex        =   25
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   18
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordFrom"
         BuddyDispid     =   196622
         OrigLeft        =   4440
         OrigTop         =   405
         OrigRight       =   4680
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabCols 
         Height          =   285
         Left            =   1530
         TabIndex        =   6
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   5
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTabCols"
         BuddyDispid     =   196629
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1770
         OrigBottom      =   360
         Max             =   30
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdRecordFont 
         Caption         =   "��������(&N)"
         Height          =   300
         Left            =   3810
         TabIndex        =   30
         Top             =   840
         Width           =   1185
      End
      Begin VB.CommandButton cmdRecordColor 
         Caption         =   "������ɫ(&L)"
         Height          =   300
         Left            =   3810
         TabIndex        =   29
         Top             =   1200
         Width           =   1185
      End
      Begin VB.TextBox txtRecordTo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5385
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "8"
         Top             =   465
         Width           =   345
      End
      Begin VB.TextBox txtRecordFrom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "18"
         Top             =   465
         Width           =   585
      End
      Begin VB.CommandButton cmdTitleFont 
         Caption         =   "��������(&T)"
         Height          =   300
         Left            =   435
         TabIndex        =   20
         Top             =   2910
         Width           =   1185
      End
      Begin VB.TextBox txtTitleText 
         Height          =   300
         Left            =   435
         TabIndex        =   19
         Text            =   "�ر����¼��"
         Top             =   2535
         Width           =   2790
      End
      Begin VB.CommandButton cmdTabGridColor 
         Caption         =   "�����ɫ(&G)"
         Height          =   300
         Left            =   435
         TabIndex        =   17
         Top             =   1785
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabTextColor 
         Caption         =   "�ı���ɫ(&R)"
         Height          =   300
         Left            =   435
         TabIndex        =   15
         Top             =   1425
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabFont 
         Caption         =   "�ı�����(&F)"
         Height          =   300
         Left            =   435
         TabIndex        =   13
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtTabRowHeight 
         Height          =   300
         Left            =   2730
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "300"
         Top             =   345
         Width           =   510
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "˫(&2)"
         Height          =   180
         Index           =   1
         Left            =   1995
         TabIndex        =   11
         Top             =   750
         Width           =   780
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "��(&1)"
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   750
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txtTabCols 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "5"
         Top             =   345
         Width           =   300
      End
      Begin MSComCtl2.UpDown udHeadRow 
         Height          =   285
         Left            =   4335
         TabIndex        =   36
         Top             =   2070
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadRow"
         BuddyDispid     =   196615
         OrigLeft        =   4920
         OrigTop         =   2085
         OrigRight       =   5160
         OrigBottom      =   2385
         Max             =   3
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "��ͬ���ݵ�ͬ�����ڵ�Ԫ�Զ��ϲ���"
         Height          =   180
         Index           =   1
         Left            =   3810
         TabIndex        =   106
         Top             =   3030
         Width           =   2880
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "��¼����ʱ������ʱʹ�øø�ʽ��"
         Height          =   180
         Index           =   0
         Left            =   3810
         TabIndex        =   104
         Top             =   1605
         Width           =   2700
      End
      Begin VB.Label lblBasic 
         AutoSize        =   -1  'True
         Caption         =   "������̬"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   103
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblHeadText 
         AutoSize        =   -1  'True
         Caption         =   "�ı�"
         Height          =   180
         Left            =   4620
         TabIndex        =   40
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lblHeadCol 
         AutoSize        =   -1  'True
         Caption         =   "�к�"
         Height          =   180
         Left            =   3630
         TabIndex        =   37
         Top             =   2490
         Width           =   360
      End
      Begin VB.Label lblHeadRow 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   180
         Left            =   3630
         TabIndex        =   34
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label lblHeadSet 
         AutoSize        =   -1  'True
         Caption         =   "��ͷ��Ԫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   33
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label lblRecordFont 
         Caption         =   "����,9"
         Height          =   180
         Left            =   5010
         TabIndex        =   32
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblRecordColor 
         Caption         =   "������ɫ"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   5010
         TabIndex        =   31
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblRecordTo 
         AutoSize        =   -1  'True
         Caption         =   "����       ��"
         Height          =   180
         Left            =   4995
         TabIndex        =   26
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordFrom 
         AutoSize        =   -1  'True
         Caption         =   "��       ����"
         Height          =   180
         Left            =   3810
         TabIndex        =   23
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordStyle 
         AutoSize        =   -1  'True
         Caption         =   "������ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblTitleFont 
         Caption         =   "����,20"
         Height          =   180
         Left            =   1635
         TabIndex        =   21
         Top             =   3015
         Width           =   1605
      End
      Begin VB.Label lblTitleText 
         AutoSize        =   -1  'True
         Caption         =   "�����ı�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   2265
         Width           =   780
      End
      Begin VB.Shape shpTabGridColor 
         Height          =   180
         Left            =   1635
         Top             =   1890
         Width           =   1605
      End
      Begin VB.Label lblTabTextColor 
         Caption         =   "�ı���ɫ"
         Height          =   180
         Left            =   1635
         TabIndex        =   16
         Top             =   1530
         Width           =   1605
      End
      Begin VB.Label lblTabFont 
         Caption         =   "����,9"
         Height          =   180
         Left            =   1635
         TabIndex        =   14
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label lblTabRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "��С�и�"
         Height          =   180
         Left            =   1935
         TabIndex        =   7
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lblTabTiers 
         AutoSize        =   -1  'True
         Caption         =   "��ͷ����"
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblTabCols 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   435
         TabIndex        =   4
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   6630
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   330
      Width           =   6885
      Begin VB.CheckBox chk 
         Caption         =   "��Ҫ����"
         Height          =   210
         Left            =   4020
         TabIndex        =   64
         Top             =   2985
         Width           =   1110
      End
      Begin VB.TextBox txtColumnPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4725
         TabIndex        =   61
         Top             =   2175
         Width           =   1965
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "Ӧ��(&Y)"
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   97
         Top             =   1635
         Width           =   1100
      End
      Begin VB.TextBox txtColumnPostfix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4725
         TabIndex        =   63
         Top             =   2535
         Width           =   1965
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   1680
         Left            =   3990
         TabIndex        =   59
         Top             =   450
         Width           =   2700
      End
      Begin MSComCtl2.UpDown udColumnNo 
         Height          =   300
         Left            =   4605
         TabIndex        =   58
         Top             =   120
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtColumnNo"
         BuddyDispid     =   196657
         OrigLeft        =   5400
         OrigTop         =   75
         OrigRight       =   5640
         OrigBottom      =   375
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtColumnNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4215
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "1"
         Top             =   120
         Width           =   405
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   55
         Top             =   1185
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2760
         TabIndex        =   54
         Top             =   885
         Width           =   1100
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   53
         Top             =   465
         Width           =   2370
      End
      Begin VB.Label lblColumnPostfix 
         AutoSize        =   -1  'True
         Caption         =   "��׺����"
         Height          =   180
         Left            =   3990
         TabIndex        =   62
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label lblColumnPrefix 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺����"
         Height          =   180
         Left            =   3990
         TabIndex        =   60
         Top             =   2235
         Width           =   720
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         Caption         =   "��        ��������Ŀ:"
         Height          =   180
         Left            =   4005
         TabIndex        =   56
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ�����¼��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.PictureBox picPaper 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   2880
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   3
      Top             =   3195
      Width           =   6885
      Begin VB.ComboBox cboFoot 
         Height          =   300
         Left            =   1050
         TabIndex        =   111
         Top             =   2970
         Width           =   5490
      End
      Begin VB.Frame fraPaper 
         Height          =   30
         Index           =   0
         Left            =   210
         TabIndex        =   98
         Top             =   2475
         Width           =   6390
      End
      Begin VB.TextBox txtPageHead 
         Height          =   300
         Left            =   1050
         TabIndex        =   95
         Top             =   2610
         Width           =   5490
      End
      Begin VB.ComboBox cboPaperKind 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   435
         Width           =   5355
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   72
         Text            =   "297.08"
         Top             =   825
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   68
         Text            =   "210.05"
         Top             =   825
         Width           =   945
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   3
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   89
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   2
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   85
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   1
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   81
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   0
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   77
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   5745
         TabIndex        =   92
         Top             =   1275
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   5745
         TabIndex        =   93
         Top             =   1680
         Width           =   750
      End
      Begin MSComCtl2.UpDown udHeight 
         Height          =   300
         Left            =   4575
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196668
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4425
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWidth 
         Height          =   300
         Left            =   2235
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196669
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   0
         Left            =   1905
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(0)"
         BuddyDispid     =   196670
         BuddyIndex      =   0
         OrigLeft        =   1785
         OrigTop         =   1410
         OrigRight       =   2025
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   1
         Left            =   3615
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(1)"
         BuddyDispid     =   196670
         BuddyIndex      =   1
         OrigLeft        =   3780
         OrigTop         =   1410
         OrigRight       =   4020
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   2
         Left            =   1905
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(2)"
         BuddyDispid     =   196670
         BuddyIndex      =   2
         OrigLeft        =   1785
         OrigTop         =   1815
         OrigRight       =   2025
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   3
         Left            =   3615
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(3)"
         BuddyDispid     =   196670
         BuddyIndex      =   3
         OrigLeft        =   3780
         OrigTop         =   1815
         OrigRight       =   4020
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "��ӡ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   270
         TabIndex        =   99
         Top             =   135
         Width           =   690
      End
      Begin VB.Label lblOrient 
         AutoSize        =   -1  'True
         Caption         =   "ֽ�ŷ���:"
         Height          =   180
         Left            =   4830
         TabIndex        =   108
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label lblPaperHint 
         AutoSize        =   -1  'True
         Caption         =   "ע��:  ���ʵ�ʴ�ӡ���͵�ǰ��ӡ�����������ܵ���ֽ������ʧЧ��"
         Height          =   180
         Left            =   270
         TabIndex        =   100
         Top             =   2115
         Width           =   5490
      End
      Begin VB.Label lblPageHead 
         AutoSize        =   -1  'True
         Caption         =   "ҳü�ı�"
         Height          =   180
         Left            =   255
         TabIndex        =   94
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lblPageFoot 
         AutoSize        =   -1  'True
         Caption         =   "ҳ���ı�"
         Height          =   180
         Left            =   255
         TabIndex        =   96
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         Caption         =   "ֽ������"
         Height          =   180
         Left            =   270
         TabIndex        =   65
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   2535
         TabIndex        =   70
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3375
         TabIndex        =   71
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   1065
         TabIndex        =   67
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   4890
         TabIndex        =   74
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         Caption         =   "ҳ�߾�:"
         Height          =   180
         Left            =   270
         TabIndex        =   75
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   2790
         TabIndex        =   88
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   2
         Left            =   1065
         TabIndex        =   84
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   2790
         TabIndex        =   80
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   1065
         TabIndex        =   76
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   2190
         TabIndex        =   79
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   3
         Left            =   3915
         TabIndex        =   83
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   2190
         TabIndex        =   87
         Top             =   1725
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   5
         Left            =   3915
         TabIndex        =   91
         Top             =   1725
         Width           =   360
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   1380
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   1
      Top             =   3330
      Width           =   6885
      Begin VB.CheckBox chkLabelCrLf 
         Caption         =   "����"
         Height          =   195
         Left            =   3870
         TabIndex        =   48
         Top             =   3000
         Width           =   780
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Ӧ��(&A)"
         Height          =   300
         Index           =   2
         Left            =   2700
         TabIndex        =   51
         Top             =   1635
         Width           =   1100
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2700
         TabIndex        =   45
         Top             =   1185
         Width           =   1100
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "ѡ��(&U)"
         Height          =   300
         Index           =   0
         Left            =   2700
         TabIndex        =   44
         Top             =   885
         Width           =   1100
      End
      Begin VB.TextBox txtLabelPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5670
         TabIndex        =   50
         Top             =   2940
         Width           =   975
      End
      Begin VB.ListBox lstLabelUsed 
         Height          =   2400
         Left            =   3870
         TabIndex        =   47
         Top             =   465
         Width           =   2775
      End
      Begin VB.ListBox lstLabelItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   43
         Top             =   465
         Width           =   2370
      End
      Begin VB.Label lblLabelPrefix 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺�ı�"
         Height          =   180
         Left            =   4905
         TabIndex        =   49
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label lblLabelUsed 
         AutoSize        =   -1  'True
         Caption         =   "�����õı�ǩ:"
         Height          =   180
         Left            =   3885
         TabIndex        =   46
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label lblLabelItems 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��ǩ��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1665
      Index           =   1
      Left            =   615
      ScaleHeight     =   1665
      ScaleWidth      =   7860
      TabIndex        =   113
      Top             =   2775
      Width           =   7860
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1425
         Left            =   0
         TabIndex        =   114
         Top             =   0
         Width           =   3810
         _cx             =   6720
         _cy             =   2514
         Appearance      =   2
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   12632256
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483644
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTendStyle.frx":058A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4005
      Index           =   0
      Left            =   1365
      ScaleHeight     =   4005
      ScaleWidth      =   6645
      TabIndex        =   112
      Top             =   -75
      Width           =   6645
      Begin XtremeSuiteControls.TabControl tbcStyle 
         Height          =   3930
         Left            =   15
         TabIndex        =   115
         Top             =   0
         Width           =   5460
         _Version        =   589884
         _ExtentX        =   9631
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picSum 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   9240
      ScaleHeight     =   1875
      ScaleWidth      =   6885
      TabIndex        =   109
      Top             =   3915
      Width           =   6885
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1020
         Left            =   435
         TabIndex        =   110
         Top             =   390
         Width           =   1995
         _cx             =   3519
         _cy             =   1799
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Editable        =   0
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
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   8265
      Top             =   6345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   116
      Top             =   7230
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendStyle.frx":0672
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17674
            Text            =   "���Ը���ҽԺʵ����������õ��������¼�Ĳ鿴�������ʽ��"
            TextSave        =   "���Ը���ҽԺʵ����������õ��������¼�Ĳ鿴�������ʽ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmTendStyle.frx":0F06
      Left            =   75
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendFileStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conRowHeight = 300        '��׼�и߶�

'��ӡֽ�ų���(256=�Զ���)
Const PageSize1 = "�ż㣬 8 1/2��11 Ӣ��"
Const PageSize2 = "+A611 С���ż㣬 8 1/2��11 Ӣ��"
Const PageSize3 = "С�ͱ��� 11��17 Ӣ��"
Const PageSize4 = "�����ʣ� 17��11 Ӣ��"
Const PageSize5 = "�����ļ��� 8 1/2��14 Ӣ��"
Const PageSize6 = "�����飬5 1/2��8 1/2 Ӣ��"
Const PageSize7 = "�����ļ���7 1/2��10 1/2 Ӣ��"
Const PageSize8 = "A3, 297��420 ����"
Const PageSize9 = "A4, 210��297 ����"
Const PageSize10 = "A4С�ţ� 210��297 ����"
Const PageSize11 = "A5, 148��210 ����"
Const PageSize12 = "B4, 250��354 ����"
Const PageSize13 = "B5, 182��257 ����"
Const PageSize14 = "�Կ����� 8 1/2��13 Ӣ��"
Const PageSize15 = "�Ŀ����� 215��275 ����"
Const PageSize16 = "10��14 Ӣ��"
Const PageSize17 = "11��17 Ӣ��"
Const PageSize18 = "������8 1/2��11 Ӣ��"
Const PageSize19 = "#9 �ŷ⣬ 3 7/8��8 7/8 Ӣ��"
Const PageSize20 = "#10 �ŷ⣬ 4 1/8��9 1/2 Ӣ��"
Const PageSize21 = "#11 �ŷ⣬ 4 1/2��10 3/8 Ӣ��"
Const PageSize22 = "#12 �ŷ⣬ 4 1/2��11 Ӣ��"
Const PageSize23 = "#14 �ŷ⣬ 5��11 1/2 Ӣ��"
Const PageSize24 = "C �ߴ繤����"
Const PageSize25 = "D �ߴ繤����"
Const PageSize26 = "E �ߴ繤����"
Const PageSize27 = "DL ���ŷ⣬ 110��220 ����"
Const PageSize28 = "C5 ���ŷ⣬ 162��229 ����"
Const PageSize29 = "C3 ���ŷ⣬ 324��458 ����"
Const PageSize30 = "C4 ���ŷ⣬ 229��324 ����"
Const PageSize31 = "C6 ���ŷ⣬ 114��162 ����"
Const PageSize32 = "C65 ���ŷ⣬114��229 ����"
Const PageSize33 = "B4 ���ŷ⣬ 250��353 ����"
Const PageSize34 = "B5 ���ŷ⣬176��250 ����"
Const PageSize35 = "B6 ���ŷ⣬ 176��125 ����"
Const PageSize36 = "�ŷ⣬ 110��230 ����"
Const PageSize37 = "�ŷ������ 3 7/8��7 1/2 Ӣ��"
Const PageSize38 = "�ŷ⣬ 3 5/8��6 1/2 Ӣ��"
Const PageSize39 = "U.S. ��׼��д���� 14 7/8��11 Ӣ��"
Const PageSize40 = "�¹���׼��д���� 8 1/2��12 Ӣ��"
Const PageSize41 = "�¹����ɸ�д���� 8 1/2��13 Ӣ��"

Private mlngFileId As Long       '���༭�ļ�¼ID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private mblnOK As Boolean        '�Ƿ���ɱ༭�˳�

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'��ʱ����
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private strTemp As String
Private lngCurColor As Long
Private strCurFont As String
Private objFont As StdFont
Private mblnChanged As Boolean

Private Property Let DataChanged(vData As Boolean)
    
    mblnChanged = vData
        
    If mblnChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
    
End Property

Private Property Get DataChanged() As Boolean
    
    DataChanged = mblnChanged

End Property

Private Function GetPaperName(ByVal intSize As Integer) As String
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '���أ� ֽ������
    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

Private Sub LoadPaper()
    '---------------------------------------------------
    '���ܣ�װ�뵱ǰ��ӡ�����õ�ֽ��
    '---------------------------------------------------
    Dim intCurPaper As Integer
    Dim strDevice As String
        
    With Me.cboPaperKind
        .AddItem GetPaperName(256)
        .ItemData(.NewIndex) = 256
        If Not ExistsPrinter Then
            .ListIndex = 0
            .Enabled = False
            Exit Sub
        End If
        
        strDevice = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
        For lngCount = 0 To Printers.Count - 1
            If Printers(lngCount).DeviceName = strDevice Then
                Set Printer = Printers(lngCount)
                Exit For
            End If
        Next
        Me.lblPrinter.Caption = "��ǰ��ӡ��: " & Printer.DeviceName
        
        intCurPaper = Printer.PaperSize
        .Enabled = True
        For lngCount = 1 To 41
            Err = 0: On Error Resume Next
            Printer.PaperSize = lngCount
            Err = 0: On Error GoTo 0
            If Printer.PaperSize = lngCount Then
                .AddItem GetPaperName(lngCount)
                .ItemData(.NewIndex) = lngCount
                If lngCount = intCurPaper Then .ListIndex = .NewIndex
            End If
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub

Public Function ShowMe(ByVal frmParent As Object, Optional ByVal lngFileId As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    Dim strTitle As String
    
    mlngFileId = lngFileId

    Err = 0: On Error GoTo errHand
    
    If RefreshData = False Then
        DataChanged = False
        Unload Me
        Exit Function
    End If
    
    DataChanged = False
    
    '---------------------------------------------------
    '������ʾ
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
'    DataChanged = False
    
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Function RefreshData() As Boolean


    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    Dim strTitle As String
    
    
    '
    With vfgThis
        .Cols = 6
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    '---------------------------------------------------
    '���ݵ�ǰ��ӡ����װ���ѡֽ��
    '---------------------------------------------------
    Call LoadPaper
    '---------------------------------------------------
    '�������ݻ�ȡ
    '---------------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select l.���, l.����, l.˵�� From �����ļ��б� l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "�����¼��ʽ - " & rsTemp!����
    strTitle = rsTemp!����
    
    gstrSQL = "Select i.������" & _
            " From ����������Ŀ i, ������������ k" & _
            " Where k.Id = i.����id And k.���� In ('02', '05', '06') And i.�滻�� = 1" & _
            " Order By i.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstLabelItems.Clear
        Do While Not .EOF
            Me.lstLabelItems.AddItem "" & !������
            .MoveNext
        Loop
        If Me.lstLabelItems.ListCount > 0 Then Me.lstLabelItems.ListIndex = 0
    End With
    
    gstrSQL = "Select ��Ŀ���� From �����¼��Ŀ Order By ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstColumnItems.Clear
        Me.lstColumnItems.AddItem "����"
        Me.lstColumnItems.AddItem "ʱ��"
        Do While Not .EOF
            Me.lstColumnItems.AddItem "" & !��Ŀ����
            .MoveNext
        Loop
        Me.lstColumnItems.AddItem "��ʿ"
        Me.lstColumnItems.AddItem "ǩ����"
        Me.lstColumnItems.AddItem "ǩ������"
        Me.lstColumnItems.AddItem "ǩ��ʱ��"
        Me.lstColumnItems.ListIndex = 0
    End With
    
    '---------------------------------------------------
    '������ʽ��ȡ
    '---------------------------------------------------
    '�ձ��ʱδ���������ͷ���,����ȱʡ��ͷ�����ǵ���,����ͷ������������СֵΪ2,�ڵ���ñ�ͷʱ����
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����"
                If Val("" & !�����ı�) = 1 Then
                    Me.optTabTiers(0).Value = True
                    Call optTabTiers_Click(0)
                ElseIf Val("" & !�����ı�) = 2 Then
                    Me.optTabTiers(1).Value = True
                    Call optTabTiers_Click(1)
                Else
                    Me.optTabTiers(2).Value = True
                    Call optTabTiers_Click(2)
                End If
            Case "������":  Me.udTabCols.Value = Val("" & !�����ı�)
            Case "��С�и�"
                Me.txtTabRowHeight.Text = Val("" & !�����ı�)
                Call txtTabRowHeight_Change
            Case "�ı�����"
                Me.lblTabFont.Caption = "" & !�����ı�
                strCurFont = Me.lblTabFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.vfgThis.Font = objFont
            Case "�ı���ɫ"
                Me.lblTabTextColor.ForeColor = Val("" & !�����ı�)
                Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
            Case "�����ɫ"
                Me.shpTabGridColor.BorderColor = Val("" & !�����ı�)
                With Me.vfgThis
                    .GridColor = Me.shpTabGridColor.BorderColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            
            Case "�����ı�"
                Me.txtTitleText.Text = "" & !�����ı�
                Call txtTitleText_Change
            Case "��������"
                Me.lblTitleFont.Caption = "" & !�����ı�
                strCurFont = Me.lblTitleFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
                    .ROWHEIGHT(1) = objFont.Size * 20 + 150
                End With
            
            Case "��ʼʱ��": Me.udRecordFrom.Value = Val("" & !�����ı�)
            Case "��ֹʱ��": Me.udRecordTo.Value = Val("" & !�����ı�)
            Case "��������"
                Me.lblRecordFont.Caption = "" & !�����ı�
                strCurFont = Me.lblRecordFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
                End With
            Case "������ɫ"
                Me.lblRecordColor.ForeColor = Val("" & !�����ı�)
                With Me.vfgThis
                    .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
                End With
            End Select
            .MoveNext
        Loop
    End With
    
    Dim strPaper As String
    gstrSQL = "Select ��ʽ, ҳü, ҳ�� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    If Not rsTemp.EOF Then strPaper = "" & rsTemp!��ʽ: Me.txtPageHead.Text = "" & rsTemp!ҳü: Me.cboFoot.Text = "" & rsTemp!ҳ��
    
    With vfgThis
        .Cell(flexcpText, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = cboFoot.Text
    End With
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        For lngCount = 0 To Me.cboPaperKind.ListCount - 1
            If Me.cboPaperKind.ItemData(lngCount) = Val(Split(strPaper, ";")(0)) Then Me.cboPaperKind.ListIndex = lngCount: Exit For
        Next
        If Me.cboPaperKind.ListIndex = 0 Then
            If UBound(Split(strPaper, ";")) >= 2 Then Me.txtHeight.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(2)), vbTwips, vbMillimeters), 2)
            If UBound(Split(strPaper, ";")) >= 3 Then Me.txtWidth.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(3)), vbTwips, vbMillimeters), 2)
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 1 Then
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.optOrient(1).Value = True
        Else
            Me.optOrient(0).Value = True
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 4 Then Me.txtMarjin(2).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then Me.txtMarjin(3).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then Me.txtMarjin(0).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then Me.txtMarjin(1).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        Me.lstLabelUsed.Clear
        Do While Not .EOF
            Me.lstLabelUsed.AddItem !�����ı� & "{" & !Ҫ������ & "}"
            Me.lstLabelUsed.ItemData(Me.lstLabelUsed.NewIndex) = !�Ƿ���
            .MoveNext
        Loop
        If Me.lstLabelUsed.ListCount > 0 Then
            Me.lstLabelUsed.ListIndex = 0
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Else
            Me.cmdLabel(1).Enabled = False
            Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
            Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
        End If
        Call cmdLabel_Click(2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        Do While Not .EOF
            Me.vfgThis.TextMatrix(!�����д� + 2, !�������) = "" & !�����ı�
            .MoveNext
        Loop
        Call udHeadCol_Change
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.vfgThis.ColWidth(!�������) = Val("" & !��������)
            If Me.udColumnNo.Value <> !������� Then Me.udColumnNo.Value = !�������
            Me.lstColumnUsed.AddItem !�����ı� & "{" & !Ҫ������ & "}" & !Ҫ�ص�λ
            Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = zlcommfun.NVL(!Ҫ�ر�ʾ, 0)
            .MoveNext
            If .EOF Then
                Call cmdColumn_Click(2)
            ElseIf Me.udColumnNo.Value <> !������� Then
                Call cmdColumn_Click(2)
            End If
        Loop
        Me.udColumnNo.Value = Me.vfgThis.Col
    End With
    
    '����ʱ��
    '------------------------------------------------------------------------------------------------------------------
    Dim aryTmp As Variant
    
    
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı� " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '����ʱ��'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With vsf
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                strTemp = zlcommfun.NVL(rsTemp!�����ı�)
                If strTemp <> "" Then
                    
                    aryTmp = Split(strTemp, ",")
                    
                    If UBound(aryTmp) >= 2 Then
                        If .TextMatrix(.Rows - 1, 1) <> "" And .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
                        
                        .TextMatrix(.Rows - 1, 1) = Trim(aryTmp(0))
                        .TextMatrix(.Rows - 1, 2) = Trim(aryTmp(1))
                        .TextMatrix(.Rows - 1, 3) = Trim(aryTmp(2))
                    End If
                End If
                rsTemp.MoveNext
            Loop
            mclsVsf.AppendRows = True
        End If
    End With
    
    '�ٰ��кϲ�
    For lngCount = 1 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    RefreshData = True
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

    
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim RS As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
        With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '�����
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "���沢�˳�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): cbrControl.ToolTipText = "�����Ѹ��ĵ�����(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "�ָ�"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�ָ����ϴα���ʱ������״̬"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.ToolTipText = "�˳���ǰ����ƴ���(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlFontSet(strTitle As String, strFont As String)
    With Me.dlgThis
        .flags = &H3 Or &H400 Or &H200 Or &H10000
        .DialogTitle = strTitle
        .FontName = Split(strFont, ",")(0)
        .FontSize = Val(Split(strFont, ",")(1))
        If InStr(1, strFont, "��") > 0 Then
            .FontBold = True
        Else
            .FontBold = False
        End If
        If InStr(1, strFont, "б") > 0 Then
            .FontItalic = True
        Else
            .FontItalic = False
        End If
        Err = 0: On Error Resume Next
        .ShowFont
        .flags = 0
        If Err.Number <> 0 Then Exit Sub
        strFont = .FontName & "," & .FontSize
        If .FontBold Or .FontItalic Then
            strFont = strFont & "," & IIf(.FontBold, "��", "") & IIf(.FontItalic, "б", "")
        End If
    End With
End Sub

Private Sub zlColorSet(strTitle As String, lngColor As Long)
    With Me.dlgThis
        .DialogTitle = strTitle
        .COLOR = lngColor
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        lngColor = .COLOR
    End With
End Sub

Private Sub cboFoot_Change()
    With vfgThis
        .Cell(flexcpText, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = cboFoot.Text
    End With
    DataChanged = True
End Sub

Private Sub cboFoot_Click()
    With vfgThis
        .Cell(flexcpText, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = cboFoot.Text
    End With
    DataChanged = True
End Sub

Private Sub cboFoot_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("@&'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cboPaperKind_Click()
    If Me.cboPaperKind.ListIndex <= 0 Then
        Me.txtWidth.Enabled = True: Me.udWidth.Enabled = True
        Me.txtHeight.Enabled = True: Me.udHeight.Enabled = True
        Me.optOrient(0).Value = True
        Me.optOrient(0).Enabled = False: Me.optOrient(1).Enabled = False
    Else
        Me.txtWidth.Enabled = False: Me.udWidth.Enabled = False
        Me.txtHeight.Enabled = False: Me.udHeight.Enabled = False
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        Err = 0: On Error Resume Next
        Printer.PaperSize = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
        Me.txtWidth.Text = Me.ScaleX(Printer.Width, vbTwips, vbMillimeters)
        Me.txtHeight.Text = Me.ScaleY(Printer.Height, vbTwips, vbMillimeters)
        If Printer.Orientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    End If
    DataChanged = True
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
            DataChanged = False
        End If
        
    Case conMenu_Edit_Transf_Cancle
                
        Call RefreshData
        DataChanged = False
        
    Case conMenu_File_Exit
        
        mblnOK = False
        Unload Me
        
    Case conMenu_Help_Help
        
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
        
    End Select
End Sub

Private Sub chk_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = chk.Value
    End With
End Sub

Private Sub chkLabelCrLf_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = Me.chkLabelCrLf.Value
    End With
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim strTmp As String
    
    With Me.lstColumnUsed
        Select Case Index
        Case 0
            If Me.lstColumnItems.ListIndex = -1 Then Exit Sub
            .AddItem "{" & Me.lstColumnItems.List(Me.lstColumnItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdColumn(1).Enabled = True
            Me.txtColumnPrefix.Enabled = True
            Me.txtColumnPostfix.Enabled = True
            chk.Enabled = (lstColumnItems.ListCount = 1)
            If chk.Enabled = False Then chk.Value = 0
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdColumn(1).Enabled = False
                Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
                Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
            End If
            
            chk.Enabled = (lstColumnItems.ListCount = 1)
        Case 2
            strTemp = ""
            strTmp = ""
            For lngCount = 0 To .ListCount - 1
                strTemp = strTemp & Space(1) & .List(lngCount) & "`" & .ItemData(lngCount)
                strTmp = strTmp & Space(1) & .List(lngCount)
            Next
            strTemp = Trim(strTemp)
            strTmp = Trim(strTmp)
            
            With vfgThis
                .TextMatrix(6, Me.udColumnNo.Value) = strTmp
                .TextMatrix(7, Me.udColumnNo.Value) = strTmp & " "
                
                .Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = strTemp
                .Cell(flexcpData, 7, udColumnNo.Value, 7, udColumnNo.Value) = strTemp & " "
                
            End With
            DataChanged = True
        End Select
    End With
End Sub

Private Sub cmdLabel_Click(Index As Integer)
    With Me.lstLabelUsed
        Select Case Index
        Case 0
            If Me.lstLabelItems.ListIndex = -1 Then Exit Sub
            .AddItem Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "��{" & Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdLabel(1).Enabled = False
                Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
                Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
            End If
        Case 2
            Dim intCrLf As Integer
            intCrLf = 0
            strTemp = ""
            For lngCount = 0 To .ListCount - 1
                If .ItemData(lngCount) <> 0 Then intCrLf = intCrLf + 1
                strTemp = strTemp & Space(1) & IIf(.ItemData(lngCount) = 0, "", vbCrLf) & .List(lngCount)
            Next
            strTemp = Trim(strTemp)
            For lngCount = Me.vfgThis.FixedCols To Me.vfgThis.Cols - 1
                Me.vfgThis.TextMatrix(2, lngCount) = strTemp
            Next
            Me.vfgThis.ROWHEIGHT(2) = Me.vfgThis.FontSize * 20 * (intCrLf + 1) + 150
            DataChanged = True
        End Select
    End With
End Sub

Private Function CheckData() As Boolean
    Dim intRow As Integer
    
    With vsf
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) <> "" And .TextMatrix(intRow, 2) <> "" And .TextMatrix(intRow, 3) <> "" Then
                
                If CheckTime(.TextMatrix(intRow, 2)) = False Then Exit Function
                If CheckTime(.TextMatrix(intRow, 3)) = False Then Exit Function
                
            End If
        Next
    End With
    
    CheckData = True
End Function

Private Function CheckTime(ByVal strText As String) As Boolean

    Dim intRow As Integer
    Dim intPos As Integer
    Dim strTmp As String
        
    intPos = InStr(strText, ":")
    
    If intPos > 0 Then
        
        strTmp = Mid(strText, 1, intPos - 1)
        If Val(strTmp) < 0 Or Val(strTmp) > 23 Then
            MsgBox "Сʱֻ����0-23֮�䣡", vbInformation, gstrSysName
            Exit Function
        End If
        
        strText = Mid(strText, intPos + 1)
        intPos = InStr(strText, ":")
        If intPos > 0 Then
            strTmp = Mid(strText, 1, intPos - 1)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr(Mid(strText, intPos + 1), ":") > 0 Then
                
                MsgBox "ʱ���ʽ����ȷ��", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Mid(strText, intPos + 1)
                If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                    MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            strTmp = strText
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    Else
        If Val(strText) < 0 Or Val(strText) > 23 Then
            MsgBox "Сʱֻ����0-23֮�䣡", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    CheckTime = True
    
End Function

Private Function SaveData() As Boolean
    
    If CheckData = False Then Exit Function
        
    If Me.optOrient(0).Value = True Then
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    End If
    
    If Me.optTabTiers(0).Value Then
        gstrSQL = mlngFileId & ",1," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    ElseIf Me.optTabTiers(1).Value Then
        gstrSQL = mlngFileId & ",2," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    Else
        gstrSQL = mlngFileId & ",3," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    End If

    gstrSQL = gstrSQL & ",'" & Me.lblTabFont.Caption & "'," & Me.lblTabTextColor.ForeColor & "," & Me.shpTabGridColor.BorderColor
    gstrSQL = gstrSQL & ",'" & Trim(Me.txtTitleText.Text) & "','" & Me.lblTitleFont.Caption & "'"
    gstrSQL = gstrSQL & "," & Me.udRecordFrom.Value & "," & Me.udRecordTo.Value & ",'" & Me.lblRecordFont.Caption & "'," & Me.lblRecordColor.ForeColor
    
    gstrSQL = gstrSQL & ",'" & Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
    gstrSQL = gstrSQL & ";" & IIf(Me.optOrient(0).Value, 1, 2)
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "'"
    gstrSQL = gstrSQL & ",'" & Me.txtPageHead.Text & "','" & Me.cboFoot.Text & "'"
    
    With Me.vfgThis
        gstrSQL = gstrSQL & ",'" & Replace(Replace(.TextMatrix(2, .FixedCols), " ", "|"), vbCrLf, "'||Chr(13)||Chr(10)||'") & "'"
        strTemp = ""
        For lngCount = .FixedCols To .Cols - 1
            If .RowHidden(3) = False Then strTemp = strTemp & "|" & lngCount & ",1," & Trim(.TextMatrix(3, lngCount))
            If .RowHidden(4) = False Then strTemp = strTemp & "|" & lngCount & ",2," & Trim(.TextMatrix(4, lngCount))
            If .RowHidden(5) = False Then strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
'            strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strTemp, 2) & "'"
        strTemp = ""
        For lngCount = .FixedCols To .Cols - 1
            strTemp = strTemp & "|" & lngCount & "," & .ColWidth(lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strTemp, 2) & "'"
    End With
    
    '��д����ʱ��
    '------------------------------------------------------------------------------------------------------------------
    With vsf
        strTemp = ""
        For lngCount = 1 To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" And .TextMatrix(lngCount, 2) <> "" And .TextMatrix(lngCount, 3) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(lngCount, 1)) & "," & Trim(.TextMatrix(lngCount, 2)) & "," & Trim(.TextMatrix(lngCount, 3))
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        gstrSQL = gstrSQL & ",'" & strTemp & "'"
    End With
    
    gstrSQL = "Zl_�����ļ���ʽ_Update(" & gstrSQL & ")"
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveData = True
    
    mblnOK = True
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdRecordColor_Click()
    lngCurColor = Me.lblRecordColor.ForeColor
    Call zlColorSet("������ʽ��ɫ", lngCurColor)
    If lngCurColor = Me.lblRecordColor.ForeColor Then Exit Sub
    Me.lblRecordColor.ForeColor = lngCurColor
    With Me.vfgThis
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
    DataChanged = True
End Sub

Private Sub cmdRecordFont_Click()
    strCurFont = Me.lblRecordFont.Caption
    Call zlFontSet("������ʽ����", strCurFont)
    If strCurFont = Me.lblRecordFont.Caption Then Exit Sub
    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
    End With
    DataChanged = True
End Sub

Private Sub cmdTabFont_Click()
    strCurFont = Me.lblTabFont.Caption
    Call zlFontSet("�ı�����", strCurFont)
    If strCurFont = Me.lblTabFont.Caption Then Exit Sub
    Me.lblTabFont.Caption = strCurFont
'    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
         .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    Set Me.vfgThis.Font = objFont
    DataChanged = True
End Sub

Private Sub cmdTabGridColor_Click()
    lngCurColor = Me.shpTabGridColor.BorderColor
    Call zlColorSet("�����ɫ", lngCurColor)
    If lngCurColor = Me.shpTabGridColor.BorderColor Then Exit Sub
    Me.shpTabGridColor.BorderColor = lngCurColor
    With Me.vfgThis
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
    End With
    DataChanged = True
End Sub

Private Sub cmdTabTextColor_Click()
    lngCurColor = Me.lblTabTextColor.ForeColor
    Call zlColorSet("�ı���ɫ", lngCurColor)
    If lngCurColor = Me.lblTabTextColor.ForeColor Then Exit Sub
    Me.lblTabTextColor.ForeColor = lngCurColor
    Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
    DataChanged = True
End Sub

Private Sub cmdTitleFont_Click()
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("��������", strCurFont)
    If strCurFont = Me.lblTitleFont.Caption Then Exit Sub
    Me.lblTitleFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
    End With
    DataChanged = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).Hwnd
    Case 2
        Item.Handle = picPane(1).Hwnd
    End Select
End Sub

Private Sub Form_Load()
    Me.picTable.BackColor = Me.BackColor
    Me.picLabel.BackColor = Me.BackColor
    Me.picCloumn.BackColor = Me.BackColor
    Me.picPaper.BackColor = Me.BackColor

    '---------------------------------------------------
    '����ҳ����֯
    With Me.tbcStyle
        .Left = Me.ScaleLeft: .Top = Me.ScaleTop
        With .PaintManager
'            .Appearance = xtpTabAppearancePropertyPage
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "��������", Me.picTable.Hwnd, 0
        .InsertItem 1, "���ϱ�ǩ", Me.picLabel.Hwnd, 0
        .InsertItem 2, "��������", Me.picCloumn.Hwnd, 0
        .InsertItem 3, "ҳ���ʽ", Me.picPaper.Hwnd, 0
        .InsertItem 4, "����ʱ��", Me.picSum.Hwnd, 0
        .Item(0).Selected = True
    End With
    With Me.vfgThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.tbcStyle.Top + Me.tbcStyle.Height + 45: .Height = Me.ScaleHeight - .Top
    End With
'    With Me.picVBar
'        .BackColor = Me.BackColor
'        .Left = Me.ScaleWidth - .Width
'        .Top = 0: .Height = Me.Height
'    End With
    
    '---------------------------------------------------
    'Ĭ����ʽ����
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        .Rows = 9
        .TextMatrix(1, 0) = "�����ı�"
        .TextMatrix(2, 0) = "���ϱ�ǩ"
        
        .TextMatrix(3, 0) = "��ͷ��Ԫ"
        .TextMatrix(4, 0) = "��ͷ��Ԫ"
        .TextMatrix(5, 0) = "��ͷ��Ԫ"
        
        .TextMatrix(6, 0) = "��������"
        .TextMatrix(7, 0) = "������ʽ"
        .TextMatrix(8, 0) = "���±�ǩ"
        
        .ColWidth(0) = 1200
        .MergeCol(0) = True
        .RowHidden(3) = True
        .RowHidden(4) = True
        
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
'            .TextMatrix(.Rows - 1, lngCount) = "��ӡʱ�䣺[��ӡʱ��]                            ��[ҳ��]ҳ"
        Next
        .MergeRow(1) = True
        .MergeRow(2) = True: .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(.Rows - 1) = True: .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(3) = True
        .MergeRow(4) = True
        .MergeRow(5) = True
        
        .Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        
        Call udTabCols_Change
        Call txtTabRowHeight_Change
        strCurFont = Me.lblTabFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        
        .ForeColor = Me.lblTabTextColor.ForeColor
        
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
            
        Call txtTitleText_Change
        strCurFont = Me.lblTitleFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
            
        strCurFont = Me.lblRecordFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
        
        .RowHidden(3) = True
        .RowHidden(4) = True
        
    End With
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    With cboFoot
        .Clear
        .AddItem ""
        .AddItem "��ӡʱ�䣺{��ӡʱ��}                          ��ӡ�ˣ�{��ӡ��}"
        .AddItem "��ӡ�ˣ�{��ӡ��}"
        .AddItem "��ӡʱ�䣺{��ӡʱ��}"
    End With
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        
        Call .Initialize(Me.Controls, vsf, True, True)
        Call .ClearColumn

        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        
        Call .AppendColumn("ʱ������", 2100, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��ʼʱ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����ʱ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
        
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(mclsVsf.ColIndex("ʱ������"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("��ʼʱ��"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("����ʱ��"), True, vbVsfEditText)
                
        .AppendRows = True
    End With
    
    vsf.ColHidden(0) = True
        
    Dim objPane As Pane
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitMenuBar
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "ʾ��": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "���": objPane.Options = PaneNoCaption
                
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMan, 1, 15, 255, Me.ScaleWidth, 255)
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If DataChanged Then
        Cancel = (MsgBox("���ĺ����Ʊ��뱣������Ч���Ƿ�������棿", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        Me.txtColumnPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        Me.txtColumnPostfix.Text = Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "}") + 1)
        chk.Value = .ItemData(.ListIndex)
    End With
End Sub

Private Sub lstLabelItems_DblClick()
    Call cmdLabel_Click(0)
End Sub

Private Sub lstLabelUsed_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        Me.chkLabelCrLf.Value = IIf(.ItemData(.ListIndex) = 0, vbUnchecked, vbChecked)
        Me.txtLabelPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
    End With
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        Cancel = (.TextMatrix(Row, 1) = "" Or .TextMatrix(Row, 2) = "" Or .TextMatrix(Row, 3) = "")
    End With
End Sub

Private Sub optOrient_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub optTabTiers_Click(Index As Integer)
    
    With vfgThis
        If optTabTiers(0).Value Then
            If .Row = 5 Or .Row = 4 Then .Row = 3
            .RowHidden(3) = False
            .RowHidden(4) = True
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 1
        ElseIf optTabTiers(1).Value Then
            If .Row = 5 Then .Row = 4
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 2
        Else
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = False
            udHeadRow.Min = 1
            udHeadRow.Max = 3
        End If

    End With
    DataChanged = True
End Sub

Private Sub optTabTiers_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub picPane_Resize(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbcStyle.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
'        cmdOK.Left = tbcStyle.Left + tbcStyle.Width + 90
'        cmdCancel.Left = tbcStyle.Left + tbcStyle.Width + 90
'        imgNote.Left = tbcStyle.Left + tbcStyle.Width + 90
'        lblNote.Left = tbcStyle.Left + tbcStyle.Width + 90
    Case 1
        vfgThis.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub picSum_Resize()
    On Error Resume Next
    
    vsf.Move 15, 15, picSum.Width - 30, picSum.Height - 30
    mclsVsf.AppendRows = True
    
End Sub

Private Sub txtColumnPostfix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "}")) & Me.txtColumnPostfix.Text
    End With
End Sub

Private Sub txtColumnPostfix_GotFocus()
    Me.txtColumnPostfix.SelStart = 0: Me.txtColumnPostfix.SelLength = 4000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txtColumnPostfix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtColumnPrefix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtColumnPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtColumnPrefix_GotFocus()
    Me.txtColumnPrefix.SelStart = 0: Me.txtColumnPrefix.SelLength = 4000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txtColumnPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeadRow_Change()
    DataChanged = True
End Sub

Private Sub txtHeadText_Change()
    Dim strInput As String
    Dim lngStart As Long, lngRow As Long, lngCol As Long
    Dim blnExist As Boolean
    
    strInput = Trim(Me.txtHeadText.Text)
    '���,��������ڵ��ĸ���Ԫ���ֵ��ͬ,����������(�п�����Ҫ������,����,����,���½��м��)
    lngRow = udHeadRow.Value + 2
    lngCol = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
    
    If lngRow <= 4 Then
        If (vfgThis.TextMatrix(3, lngCol) = vfgThis.TextMatrix(4, lngCol) And vfgThis.TextMatrix(3, lngCol) <> "") Then
            If lngCol > 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
    If lngRow >= 4 Then
        If (vfgThis.TextMatrix(4, lngCol) = vfgThis.TextMatrix(5, lngCol) And vfgThis.TextMatrix(4, lngCol) <> "") Then
            If lngCol > 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
Limit:
    If blnExist Then strInput = strInput & "_1"

WriteIt:
    Me.txtHeadText.Text = strInput
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    DataChanged = True
End Sub

Private Sub txtHeadText_GotFocus()
    Me.txtHeadText.SelStart = 0: Me.txtHeadText.SelLength = 4000
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txtHeadText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeight_Change()
    DataChanged = True
End Sub

Private Sub txtLabelPrefix_Change()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtLabelPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtLabelPrefix_GotFocus()
    Me.txtLabelPrefix.SelStart = 0: Me.txtLabelPrefix.SelLength = 4000
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txtLabelPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txtPageHead_Change()
    DataChanged = True
End Sub

Private Sub txtPageHead_GotFocus()
    Me.txtPageHead.SelStart = 0: Me.txtPageHead.SelLength = 4000
    Call zlcommfun.OpenIme(True)
    
End Sub

Private Sub txtPageHead_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("@&'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtRecordFrom_Change()
    DataChanged = True
End Sub

Private Sub txtRecordTo_Change()
    DataChanged = True
End Sub

Private Sub txtTabCols_Change()
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_Change()
    Me.vfgThis.RowHeightMin = Val(Me.txtTabRowHeight.Text)
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_GotFocus()
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txtTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtTitleText_Change()
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(1, lngCount) = Trim(Me.txtTitleText.Text)
        Next
    End With
    
    DataChanged = True
End Sub

Private Sub txtTitleText_GotFocus()
    Me.txtTitleText.SelStart = 0: Me.txtTitleText.SelLength = 4000
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txtTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    DataChanged = True
End Sub

Private Sub udColumnNo_Change()
    Me.lstColumnUsed.Clear
'    strTemp = Trim(Me.vfgThis.TextMatrix(5, Me.udColumnNo.Value))
'
    strTemp = vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value)
    
    If strTemp = "" Then
        Me.lstColumnUsed.ListIndex = -1
        Me.cmdColumn(1).Enabled = False
        Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
        Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
    Else
        Dim aryCol() As String
        aryCol = Split(strTemp, Space(1))
        For lngCount = 0 To UBound(aryCol)
            If InStr(aryCol(lngCount), "`") > 0 Then
                Me.lstColumnUsed.AddItem Mid(aryCol(lngCount), 1, InStr(aryCol(lngCount), "`") - 1)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Mid(aryCol(lngCount), InStr(aryCol(lngCount), "`") + 1))
            Else
                Me.lstColumnUsed.AddItem aryCol(lngCount)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = 0
            End If
        Next
        Me.lstColumnUsed.ListIndex = 0
        Me.cmdColumn(1).Enabled = True
        Me.txtColumnPrefix.Enabled = True
        Me.txtColumnPostfix.Enabled = True
        
        chk.Enabled = (lstColumnUsed.ListCount = 1)
        
    End If
End Sub

Private Sub udHeadCol_Change()
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    txtHeadText.Text = vfgThis.TextMatrix(udHeadRow.Value + 2, udHeadCol.Value)

    DataChanged = blnSvrChanged
End Sub

Private Sub udHeadRow_Change()

    Call udHeadCol_Change
    
End Sub

Private Sub udRecordFrom_Change()
    If Me.udRecordFrom.Value > Me.udRecordTo.Value Then
        Me.lblRecordTo.Caption = "����" & Space(7) & "��"
    Else
        Me.lblRecordTo.Caption = "����" & Space(7) & "��"
    End If
End Sub

Private Sub udRecordTo_Change()
    Call udRecordFrom_Change
End Sub

Private Sub udTabCols_Change()
    Me.vfgThis.Cols = Me.udTabCols.Value + 1
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
    Me.udHeadCol.Max = Me.udTabCols.Value
    If Val(Me.txtHeadCol.Text) > Me.udHeadCol.Max Then Me.txtHeadCol.Text = Me.udHeadCol.Max
    Me.udColumnNo.Max = Me.udTabCols.Value
    If Val(Me.txtColumnNo.Text) > Me.udColumnNo.Max Then Me.txtColumnNo.Text = Me.udColumnNo.Max
    
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .TextMatrix(1, lngCount) = .TextMatrix(1, .FixedCols)
            .TextMatrix(2, lngCount) = .TextMatrix(2, .FixedCols)
            .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, .FixedCols)
        Next
        .Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, Me.shpTabGridColor.BorderColor, 1, 1, 1, 1, 1, 1
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    Me.udHeadCol.Value = NewCol
    Me.udColumnNo.Value = NewCol
    
    If NewRow >= 3 And NewRow <= 5 Then udHeadRow.Value = NewRow - 2
    
     DataChanged = blnSvrChanged
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    DataChanged = True
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '�༭����
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)

End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick()
    '�༭����
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    '�༭����
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    
    '�༭����,������
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("��ʼʱ��"), vsf.ColIndex("����ʱ��")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select
        
        
    End If
    Call mclsVsf.KeyPress(KeyAscii)

End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("��ʼʱ��"), vsf.ColIndex("����ʱ��")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select

    End If
    Call mclsVsf.KeyPressEdit(KeyAscii)
    
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub
