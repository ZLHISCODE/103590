VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1935
      Left            =   840
      TabIndex        =   74
      Top             =   5520
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   31
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6240
      TabIndex        =   32
      Top             =   960
      Width           =   1100
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   5175
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   34
         Top             =   600
         Width           =   5505
         Begin MSComctlLib.ListView lvw���� 
            Height          =   2835
            Left            =   1200
            TabIndex        =   46
            Top             =   3960
            Visible         =   0   'False
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   5001
            View            =   1
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "imgsDrug"
            SmallIcons      =   "imgsDrug"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TreeView tvw��� 
            Height          =   4245
            Left            =   -240
            TabIndex        =   38
            Top             =   3960
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   7488
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imgsDrug"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.ComboBox cbo���Ʒ��� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   750
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.ComboBox Cbo�ƻ����� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.CheckBox chk���Ʒ��� 
            Caption         =   "���Ʒ���"
            Height          =   300
            Left            =   600
            TabIndex        =   41
            Top             =   720
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CheckBox Chk�ƻ����� 
            Caption         =   "�ƻ�����"
            Height          =   300
            Left            =   600
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox txt������ 
            Height          =   300
            Left            =   1515
            MaxLength       =   8
            TabIndex        =   73
            Top             =   3690
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.ComboBox Cbo�ⷿ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   51
            Text            =   "Cbo�ⷿ"
            Top             =   1530
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1530
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CheckBox Chk����ⷿ 
            Caption         =   "����ⷿ"
            Height          =   300
            Left            =   600
            TabIndex        =   50
            Top             =   1530
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   44
            Top             =   750
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox Txt������Ʊ�� 
            Height          =   300
            Left            =   3780
            TabIndex        =   71
            Top             =   3330
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Txt��ʼ��Ʊ�� 
            Height          =   300
            Left            =   1530
            TabIndex        =   69
            Top             =   3330
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   67
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   65
            Top             =   2940
            Width           =   1365
         End
         Begin VB.CheckBox chkClass 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   600
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   45
            Top             =   750
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��"
            Height          =   300
            Left            =   600
            TabIndex        =   52
            Top             =   1530
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   54
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Chk������ 
            Caption         =   "������"
            Height          =   300
            Left            =   600
            TabIndex        =   55
            Top             =   1920
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txt������ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   56
            Top             =   1920
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CommandButton Cmd������ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   57
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   49
            Top             =   1140
            Width           =   255
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   48
            Top             =   1140
            Width           =   3255
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ"
            Height          =   300
            Left            =   600
            TabIndex        =   47
            Top             =   1140
            Width           =   990
         End
         Begin VB.CheckBox chk��Ʊ���� 
            Caption         =   "��Ʊ�������"
            Height          =   405
            Left            =   600
            TabIndex        =   60
            Top             =   2340
            Visible         =   0   'False
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpStart��Ʊ 
            Height          =   315
            Left            =   1650
            TabIndex        =   61
            Top             =   2340
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   141819907
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpEnd��Ʊ 
            Height          =   315
            Left            =   3600
            TabIndex        =   63
            Top             =   2340
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   129499139
            CurrentDate     =   36263
         End
         Begin VB.ComboBox Cbo��� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1920
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.CheckBox Chk��� 
            Caption         =   "���"
            Height          =   300
            Left            =   600
            TabIndex        =   58
            Top             =   1920
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   930
            TabIndex        =   72
            Top             =   3750
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   3240
            TabIndex        =   70
            Top             =   3390
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Lbl��Ʊ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊ��"
            Height          =   180
            Left            =   975
            TabIndex        =   68
            Top             =   3390
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   3120
            TabIndex        =   66
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   975
            TabIndex        =   64
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   3360
            TabIndex        =   62
            Top             =   2400
            Visible         =   0   'False
            Width           =   180
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   4170
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkAccStrike 
            Caption         =   "�Ѳ������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   20
            Top             =   2760
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chk�ѱ�� 
            Caption         =   "����������"
            Height          =   255
            Left            =   720
            TabIndex        =   25
            Top             =   3113
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkδ��� 
            Caption         =   "δ��������"
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   3113
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk�з�Ʊ 
            Caption         =   "�з�Ʊ"
            Height          =   255
            Left            =   720
            TabIndex        =   27
            Top             =   3487
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chk�޷�Ʊ 
            Caption         =   "�޷�Ʊ"
            Height          =   255
            Left            =   2400
            TabIndex        =   28
            Top             =   3487
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkAcc 
            Caption         =   "δ�������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   19
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "������˿�"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2400
            TabIndex        =   30
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "δ����˿�"
            Height          =   180
            Left            =   720
            TabIndex        =   29
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   17
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��˵���"
            Height          =   420
            Left            =   480
            TabIndex        =   4
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txt����NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   3
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt��ʼNo 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   2
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "δ��˳���"
            Height          =   300
            Left            =   720
            TabIndex        =   10
            Top             =   1400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "����˳���"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   16
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   9
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   312
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   312
            Index           =   1
            Left            =   3588
            TabIndex        =   15
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   2
            Left            =   3600
            TabIndex        =   24
            Top             =   2835
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   22
            Top             =   2835
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Ѹ��˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   18
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chk�Ѵ�ӡ 
            Caption         =   "�Ѵ�ӡ����"
            Height          =   255
            Left            =   720
            TabIndex        =   75
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkδ��ӡ 
            Caption         =   "δ��ӡ����"
            Height          =   255
            Left            =   2400
            TabIndex        =   76
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   7
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   5
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   2640
            TabIndex        =   8
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   12
            Top             =   2028
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   14
            Top             =   2034
            Width           =   180
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   1
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3360
            TabIndex        =   23
            Top             =   2895
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   2
            Left            =   900
            TabIndex        =   21
            Top             =   2895
            Visible         =   0   'False
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   6480
      Top             =   1560
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
            Picture         =   "frmSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long                 '��������
Private mfrmMain As Form                 '������
Private mint�������� As Integer          '0-����Ҫ����;1-��Ҫ����
Private mblnAdvance As Boolean           '�Ƿ�չ��������������
Private mstrMatch As String              'ƥ�䷽ʽ 0-˫��ƥ�� 1-�������ҵ���ƥ��
Private mstrSelectTag As String          '��ǰѡ��Ķ���
Private mblnStock As Boolean             '��ǰ����Ա�Ƿ���ҩ����Ա���������õ�����Ч
Private mint������� As Integer
Private mblnCancel As Boolean            '���ȡ��

Private Const mint���� As Integer = 0
Private Const mint��� As Integer = 1
Private Const mint���� As Integer = 2
Private Const mintNo As Integer = 3
Private Const mint��Ʊ���� As Integer = 4
Private Const mint��Ʊ�� As Integer = 5

Private Const mintTab��Χ As Integer = 0

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    lngҩƷ As Long
    lng�ⷿ As Long
    str������ As String
    str����� As String
    str������ As String
    lng�ƻ����� As Long
    lng���Ʒ��� As Long
    lng��Ӧ�� As Long
    lng������ As Long
    str���� As String
    lng������ As Long
    str��Ʊ�ſ�ʼ As String
    str��Ʊ�Ž��� As String
    int�������һ����ѯ As Integer
    intδ��� As Integer
    int�ѱ�� As Integer
    int�з�Ʊ As Integer
    int�޷�Ʊ As Integer
    lngҩƷ���� As Long
    str���� As String
    date��Ʊ������ڿ�ʼ As Date
    date��Ʊ������ڽ��� As Date
    intδ��ӡ As Integer
    int�Ѵ�ӡ As Integer
End Type

Private SQLCondition As Type_SQLCondition

Private Type Type_TemporaryInquiries
    intδ��˵��� As Integer 'intδ������ As Integer
    intδ��˳��� As Integer
    int����˵��� As Integer 'int�Ѵ����� As Integer
    int����˳��� As Integer
    int�Ѹ��˵��� As Integer
    int�������� As Integer
    
    intδ������� As Integer
    int�Ѳ������ As Integer
    
    intδ����˿� As Integer
    int������˿� As Integer
End Type

Private TemporaryInquiries As Type_TemporaryInquiries   '��ʱ��ѯ���������ڻָ��ϴ����õĹ�����������������رպ�������������ã�

'ͨ�ù��˴����õ��ļ���Keyֵ����
Private Const mstrNO��ʼKey As String = "NO��ʼ"
Private Const mstrNO����Key As String = "NO����"
Private Const mstr����ʱ�俪ʼKey As String = "����ʱ�俪ʼ"
Private Const mstr����ʱ�����Key As String = "����ʱ�����"
Private Const mstr���ʱ�俪ʼKey As String = "���ʱ�俪ʼ"
Private Const mstr���ʱ�����Key As String = "���ʱ�����"
Private Const mstr����ʱ�俪ʼKey As String = "����ʱ�俪ʼ"
Private Const mstr����ʱ�����Key As String = "����ʱ�����"
Private Const mstrҩƷIDKey As String = "ҩƷID"
Private Const mstr��Ӧ��Key As String = "��Ӧ��"
Private Const mstr������Key As String = "������"
Private Const mstr�����Key As String = "�����"
Private Const mstr������Key As String = "������"
Private Const mstr�ƻ�����Key As String = "�ƻ�����"
Private Const mstr���Ʒ���Key As String = "���Ʒ���"
Private Const mstr�ⷿIDKey As String = "�ⷿID"
Private Const mstrδ��˵���Key As String = "δ��˵���" 'δ������
Private Const mstr����˵���Key As String = "����˵���" '�Ѵ�����
Private Const mstrδ��˳���Key As String = "δ��˳���"
Private Const mstr����˳���Key As String = "����˳���"
Private Const mstr�Ѹ��˵���Key As String = "�Ѹ��˵���"
Private Const mstr����Key As String = "����"             '������
Private Const mstr������Key As String = "������"
Private Const mstr��������Key As String = "��������"
Private Const mstr��Ʊ�ſ�ʼKey As String = "��Ʊ�ſ�ʼ"
Private Const mstr��Ʊ�Ž���Key As String = "��Ʊ�Ž���"
Private Const mstrҩƷ����Key As String = "ҩƷ����"
Private Const mstr����Key As String = "����"
Private Const mstr��Ʊ������ڿ�ʼKey As String = "��Ʊ������ڿ�ʼ"
Private Const mstr��Ʊ������ڽ���Key As String = "��Ʊ������ڽ���"
Private Const mstr�ޱ��Key As String = "�ޱ��"
Private Const mstr�б��Key As String = "�б��"
Private Const mstr�޷�ƱKey As String = "�޷�Ʊ"
Private Const mstr�з�ƱKey As String = "�з�Ʊ"
Private Const mstr�������һ����ѯKey As String = "�������һ����ѯ"
Private Const mstrδ�������Key As String = "δ�������"
Private Const mstr�Ѳ������Key As String = "�Ѳ������"
Private Const mstrδ����˿�Key As String = "δ����˿�"
Private Const mstr������˿�Key As String = "������˿�"
Private Const mstrδ��ӡKey As String = "δ��ӡ"
Private Const mstr�Ѵ�ӡKey As String = "�Ѵ�ӡ"

Public Property Get getKey_NO��ʼ() As String
    getKey_NO��ʼ = mstrNO��ʼKey
End Property

Public Property Get getKey_NO����() As String
    getKey_NO���� = mstrNO����Key
End Property

Public Property Get getKey_����ʱ�俪ʼ() As String
    getKey_����ʱ�俪ʼ = mstr����ʱ�俪ʼKey
End Property

Public Property Get getKey_����ʱ�����() As String
    getKey_����ʱ����� = mstr����ʱ�����Key
End Property

Public Property Get getKey_���ʱ�俪ʼ() As String
    getKey_���ʱ�俪ʼ = mstr���ʱ�俪ʼKey
End Property

Public Property Get getKey_���ʱ�����() As String
    getKey_���ʱ����� = mstr���ʱ�����Key
End Property

Public Property Get getKey_����ʱ�俪ʼ() As String
    getKey_����ʱ�俪ʼ = mstr����ʱ�俪ʼKey
End Property

Public Property Get getKey_����ʱ�����() As String
    getKey_����ʱ����� = mstr����ʱ�����Key
End Property

Public Property Get getKey_ҩƷID() As String
    getKey_ҩƷID = mstrҩƷIDKey
End Property

Public Property Get getKey_��Ӧ��() As String
    getKey_��Ӧ�� = mstr��Ӧ��Key
End Property

Public Property Get getKey_������() As String
    getKey_������ = mstr������Key
End Property

Public Property Get getKey_�����() As String
    getKey_����� = mstr�����Key
End Property

Public Property Get getKey_������() As String
    getKey_������ = mstr������Key
End Property

Public Property Get getKey_�ƻ�����() As String
    getKey_�ƻ����� = mstr�ƻ�����Key
End Property

Public Property Get getKey_���Ʒ���() As String
    getKey_���Ʒ��� = mstr���Ʒ���Key
End Property

Public Property Get getKey_�ⷿID() As String
    getKey_�ⷿID = mstr�ⷿIDKey
End Property

Public Property Get getKey_δ��˵���() As String
    getKey_δ��˵��� = mstrδ��˵���Key
End Property

Public Property Get getKey_����˵���() As String
    getKey_����˵��� = mstr����˵���Key
End Property

Public Property Get getKey_δ��˳���() As String
    getKey_δ��˳��� = mstrδ��˳���Key
End Property

Public Property Get getKey_����˳���() As String
    getKey_����˳��� = mstr����˳���Key
End Property

Public Property Get getKey_�Ѹ��˵���() As String
    getKey_�Ѹ��˵��� = mstr�Ѹ��˵���Key
End Property

Public Property Get getKey_����() As String
    getKey_���� = mstr����Key
End Property

Public Property Get getKey_������() As String
    getKey_������ = mstr������Key
End Property

Public Property Get getKey_��������() As String
    getKey_�������� = mstr��������Key
End Property

Public Property Get getKey_��Ʊ�ſ�ʼ() As String
    getKey_��Ʊ�ſ�ʼ = mstr��Ʊ�ſ�ʼKey
End Property

Public Property Get getKey_��Ʊ�Ž���() As String
    getKey_��Ʊ�Ž��� = mstr��Ʊ�Ž���Key
End Property

Public Property Get getKey_ҩƷ����() As String
    getKey_ҩƷ���� = mstrҩƷ����Key
End Property

Public Property Get getKey_����() As String
    getKey_���� = mstr����Key
End Property

Public Property Get getKey_��Ʊ������ڿ�ʼ() As String
    getKey_��Ʊ������ڿ�ʼ = mstr��Ʊ������ڿ�ʼKey
End Property

Public Property Get getKey_��Ʊ������ڽ���() As String
    getKey_��Ʊ������ڽ��� = mstr��Ʊ������ڽ���Key
End Property

Public Property Get getKey_�ޱ��() As String
    getKey_�ޱ�� = mstr�ޱ��Key
End Property

Public Property Get getKey_�б��() As String
    getKey_�б�� = mstr�б��Key
End Property

Public Property Get getKey_�޷�Ʊ() As String
    getKey_�޷�Ʊ = mstr�޷�ƱKey
End Property

Public Property Get getKey_�з�Ʊ() As String
    getKey_�з�Ʊ = mstr�з�ƱKey
End Property

Public Property Get getKey_�������һ����ѯ() As String
    getKey_�������һ����ѯ = mstr�������һ����ѯKey
End Property

Public Property Get getKey_δ�������() As String
    getKey_δ������� = mstrδ�������Key
End Property

Public Property Get getKey_�Ѳ������() As String
    getKey_�Ѳ������ = mstr�Ѳ������Key
End Property

Public Property Get getKey_δ����˿�() As String
    getKey_δ����˿� = mstrδ����˿�Key
End Property

Public Property Get getKey_������˿�() As String
    getKey_������˿� = mstr������˿�Key
End Property

Public Property Get getKey_δ��ӡ() As String
    getKey_δ��ӡ = mstrδ��ӡKey
End Property

Public Property Get getKey_�Ѵ�ӡ() As String
    getKey_�Ѵ�ӡ = mstr�Ѵ�ӡKey
End Property

Public Property Get In_�������() As Integer
    In_������� = mint�������
End Property

Public Property Let In_�������(ByVal vNewValue As Integer)
    mint������� = vNewValue
End Property

Private Sub cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    '��ȡ�ɲ����Ŀⷿ
    Select Case mlngMode
        Case ģ���.ҩƷ�ƿ�
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.ҩƷ����
            str�������� = "O"
        Case ģ���.��������
            Exit Sub
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo�ⷿ.ListCount = 0 Then Exit Sub
    
    If cbo�ⷿ.ListIndex >= 0 Then
        If Val(cbo�ⷿ.Tag) = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cbo�ⷿ, Trim(cbo�ⷿ.Text), str��������) = False Then
        Exit Sub
    End If
    If cbo�ⷿ.ListIndex >= 0 Then
        cbo�ⷿ.Tag = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    End If
End Sub

Private Sub chkδ��ӡ_Click()
    If chkδ��ӡ.Value = 1 Then chk�Ѵ�ӡ.Value = 0
End Sub

Private Sub chk�Ѵ�ӡ_Click()
    If chk�Ѵ�ӡ.Value = 1 Then chkδ��ӡ.Value = 0
End Sub

Private Sub chkAcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkAccStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkClass_Click()
    If chkClass.Value = 1 Then
        txtClass.Enabled = True
        cmdClass.Enabled = True
    Else
        txtClass.Enabled = False
        cmdClass.Enabled = False
    End If
End Sub

Private Sub chkClass_GotFocus()
    If sstFilter.Tab = 0 Then sstFilter.Tab = 1
End Sub

Private Sub chkClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkJiXin_Click()
    If chkJiXin.Value = 1 Then
        txtJiXing.Enabled = True
        cmdJiXin.Enabled = True
    Else
        txtJiXing.Enabled = False
        cmdJiXin.Enabled = False
    End If
End Sub

Private Sub chkJiXin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkNoStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkNOVerifyBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkYesStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkYesVerifyBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk���Ʒ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk��Ʊ����_Click()
    If chk��Ʊ����.Value = 1 Then
        dtpStart��Ʊ.Enabled = True
        dtpEnd��Ʊ.Enabled = True
    Else
        dtpStart��Ʊ.Enabled = False
        dtpEnd��Ʊ.Enabled = False
    End If
End Sub

Private Sub chk��Ʊ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chk����.Value = 1 Then
            SendKeys vbTab
        Else
            cmdȷ��.SetFocus
        End If
    End If
End Sub

Private Sub Chk��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk�ƻ�����_Click()
    Cbo�ƻ�����.Enabled = IIf(Chk�ƻ�����.Value = 1, True, False)
End Sub

Private Sub chk���Ʒ���_Click()
    cbo���Ʒ���.Enabled = IIf(chk���Ʒ���.Value = 1, True, False)
End Sub

Private Sub Chk�ƻ�����_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk�ƻ�����.SetFocus
    End If
End Sub

Private Sub Chk�ƻ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk����_GotFocus()
    If sstFilter.Tab = 1 Then sstFilter.Tab = 0
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkδ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk�޷�Ʊ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    CmdҩƷ.Enabled = TxtҩƷ.Enabled
End Sub

Private Sub Chk��Ӧ��_Click()
    txt��Ӧ��.Enabled = IIf(chk��Ӧ��.Value = 1, True, False)
    cmd��Ӧ��.Enabled = txt��Ӧ��.Enabled
End Sub

Private Sub Chk������_Click()
    Me.txt������.Enabled = IIf(Chk������.Value = 1, True, False)
    Cmd������.Enabled = IIf(Chk������.Value = 1, True, False)
End Sub

Private Sub Chk���_Click()
    Cbo���.Enabled = IIf(Chk���.Value = 1, True, False)
End Sub

Private Sub chkStrike_Click()
    chkAccStrike.Enabled = IIf(chkStrike.Value = 1, True, False)
End Sub

Private Sub chk����_Click()
    DTP��ʼʱ��(mint����).Enabled = IIf(chk����.Value = 1, True, False)
    DTP����ʱ��(mint����).Enabled = DTP��ʼʱ��(mint����).Enabled
End Sub

Private Sub chk���_Click()
    DTP��ʼʱ��(mint���).Enabled = IIf(chk���.Value = 1, True, False)
    DTP����ʱ��(mint���).Enabled = IIf(chk���.Value = 1, True, False)
    
    Select Case mlngMode
        Case ģ���.�⹺���
            chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
            chk�ѱ��.Enabled = IIf(chk���.Value = 1, True, False)
            chkδ���.Enabled = IIf(chk���.Value = 1, True, False)
            chkAcc.Enabled = IIf(chk���.Value = 1, True, False)
            chkYesVerifyBack.Enabled = IIf(chk���.Value = 1, True, False)
            If chk���.Value = 0 Then chkYesVerifyBack.Value = 0
        Case ģ���.�������, ģ���.�������, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
            If ģ���.ҩƷ�ƿ� = mlngMode Then chkYesStrike.Enabled = IIf(chk���.Value = 1, True, False)
    End Select
End Sub

Private Sub chk����_Click()
    DTP��ʼʱ��(mint����).Enabled = IIf(chk����.Value = 1, True, False)
    DTP����ʱ��(mint����).Enabled = IIf(chk����.Value = 1, True, False)

    Select Case mlngMode
        Case ģ���.�⹺���
            chkNOVerifyBack.Enabled = IIf(chk����.Value = 1, True, False)
            If chk����.Value = 0 Then chkNOVerifyBack.Value = 0
            chkNoStrike.Enabled = IIf(chk����.Value = 1, True, False)
        Case ģ���.ҩƷ�ƿ� ', ģ���.ҩƷ����, ģ���.��������
            chkNoStrike.Enabled = IIf(chk����.Value = 1, True, False)
    End Select
End Sub

Private Sub chkδ���_Click()
    If chkδ���.Value = 1 Then
        chk�ѱ��.Value = 0
    End If
End Sub

Private Sub chk�޷�Ʊ_Click()
    If chk�޷�Ʊ.Value = 1 Then
        chk�з�Ʊ.Value = 0
    End If
End Sub

Private Sub ChkҩƷ_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        ChkҩƷ.SetFocus
    End If
End Sub

Private Sub ChkҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk����ⷿ_click()
    cbo�ⷿ.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
End Sub

Private Sub Chk����ⷿ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk�ѱ��_Click()
    If chk�ѱ��.Value = 1 Then
        chkδ���.Value = 0
    End If
End Sub

Private Sub chk�ѱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk�з�Ʊ_Click()
    If chk�з�Ʊ.Value = 1 Then
        chk�޷�Ʊ.Value = 0
    End If
End Sub

Private Sub chk�з�Ʊ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    tvw���.Left = txtClass.Left
    tvw���.Top = txtClass.Top + txtClass.Height
    tvw���.Visible = True
    tvw���.SetFocus
        
    gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
              "Where Instr([1], ����, 1) > 0 " & _
              "Order by ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw���
        .Nodes.Clear
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!����, rsTmp!����, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!����
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    gstrSQL = "Select ID, �ϼ�ID, ����, 1 as ĩ��, decode(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') as ����, ���� " & _
                  "From ���Ʒ���Ŀ¼ " & _
                  "Where ���� in (1,2,3) " & _
                  "Start With �ϼ�ID IS NULL Connect By Prior ID=�ϼ�ID Order by level,ID "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��;����")
    
    With rsTmp
        If .EOF Then
            Exit Sub
        End If
        
        '��ҩƷ��;��������װ��
        Do While Not .EOF
            Intĩ�� = IIf(!ĩ�� = 1, 3, 2)
            If IsNull(!�ϼ�ID) Then
                Set nodTmp = tvw���.Nodes.Add("Root" & !����, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
            Else
                Set nodTmp = tvw���.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
            End If
            nodTmp.Tag = !����   '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With

    With tvw���
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Intĩ�� = 1
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Intĩ�� = 2
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Intĩ�� = 3
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Intĩ�� = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdJiXin_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿID As Long
    
    lvw����.Left = txtJiXing.Left
    lvw����.Top = txtJiXing.Top + txtJiXing.Height
    lvw����.Visible = True
    lvw����.SetFocus
    
    On Error GoTo errHandle
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng�ⷿID <> 0 Then
        '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
        gstrSQL = "Select Distinct J.����,J.���� " & _
                  "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                  "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] " & _
                  "Order by J.���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿID)
    Else
        gstrSQL = "Select ����,���� From ҩƷ���� order by ���� "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����")
    End If
    
    With rsTmp
        lvw����.ListItems.Clear
        Do While Not .EOF
            lvw����.ListItems.Add , "K" & !����, !����, 1, 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              " Where (վ�� = [1] Or վ�� is Null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              " Start with �ϼ�ID is null and (վ�� = [1] Or վ�� is Null) " & _
              " connect by prior ID =�ϼ�ID and (վ�� = [1] Or վ�� is Null) order by level,ID"
    
    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "����", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt��Ӧ��.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt��Ӧ��.SetFocus
    txt��Ӧ��.Tag = rsProvider!id
    txt��Ӧ��.Text = rsProvider!����
    
    If mlngMode = ģ���.�������� Then
        Txt������.SetFocus
    ElseIf mlngMode = ģ���.�⹺��� Then
        If Chk������.Value = 1 Then
            txt������.SetFocus
        Else
            Chk������.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmdȡ��_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '�ù��̸�ģ�鶼�еļ��
    If chk����.Value = 0 And chk���.Value = 0 Then
        If mlngMode = ģ���.�������� Then
            MsgBox "�Բ��𣬱���ѡ��һ���Ǽ����ڻ��ߴ�������!", vbInformation, gstrSysName
            chk����.SetFocus
            Exit Sub
        ElseIf mlngMode = ģ���.ҩƷ�ƻ� Then
            If chk����.Value = 0 Then
                MsgBox "�Բ��𣬱���ѡ��һ���������ڻ���������ڻ��߸�������!", vbInformation, gstrSysName
                chk����.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "�Բ��𣬱���ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
            chk����.SetFocus
            Exit Sub
        End If
    End If
    
    If mlngMode <> ģ���.�������� Then
        intNO = Switch(mlngMode = ģ���.�⹺���, 21, mlngMode = ģ���.�������, 24, mlngMode = ģ���.�������, 22, _
                        mlngMode = ģ���.��۵���, 25, mlngMode = ģ���.ҩƷ�ƿ�, 26, mlngMode = ģ���.ҩƷ����, 27, _
                        mlngMode = ģ���.��������, 28, mlngMode = ģ���.ҩƷ�ƻ�, 32)
        lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    End If
    
    '===================��Χѡ�==================
    Select Case mlngMode
        Case ģ���.��������
            If ChkҩƷ.Value = 1 Then
                If TxtҩƷ.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
                    Me.TxtҩƷ.SetFocus
                    Exit Sub
                End If
            End If
            If chk��Ӧ��.Value = 1 Then
                If txt��Ӧ��.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
                    Me.txt��Ӧ��.SetFocus
                    Exit Sub
                End If
            End If
        Case ģ���.ҩƷ�ƻ�
            If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
                txt��ʼNo.Text = zlCommFun.GetFullNo(txt��ʼNo.Text, intNO, lng�ⷿID)
            End If
            If Len(txt����No) < 8 And Len(txt����No) > 0 Then
                txt����No.Text = zlCommFun.GetFullNo(txt����No.Text, intNO, lng�ⷿID)
            End If
            
            SQLCondition.strNO��ʼ = Me.txt��ʼNo
            SQLCondition.strNO���� = Me.txt����No
            SQLCondition.date����ʱ�俪ʼ = CDate(Format(DTP��ʼʱ��(mint����), "yyyy-mm-dd") & " 00:00:00")
            SQLCondition.date����ʱ����� = CDate(Format(DTP����ʱ��(mint����), "yyyy-mm-dd") & " 23:59:59")
            TemporaryInquiries.int�Ѹ��˵��� = chk����.Value
            
        Case ģ���.�������, ģ���.�������
            '�������
            If ChkҩƷ.Value = 1 Then
                If TxtҩƷ.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
                    Me.TxtҩƷ.SetFocus
                    Exit Sub
                End If
            End If
            
            If mlngMode = ģ���.������� Then
                If Chk������.Value = 1 Then
                    If txt������.Tag = 0 Then
                        MsgBox "��ѡ�����ѯ��ҩƷ��������Ϣ��", vbInformation, gstrSysName
                        Me.txt������.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
                txt��ʼNo.Text = zlCommFun.GetFullNo(txt��ʼNo.Text, intNO, lng�ⷿID)
            End If
            If Len(txt����No) < 8 And Len(txt����No) > 0 Then
                txt����No.Text = zlCommFun.GetFullNo(txt����No.Text, intNO, lng�ⷿID)
            End If
            
            SQLCondition.strNO��ʼ = Me.txt��ʼNo
            SQLCondition.strNO���� = Me.txt����No
            TemporaryInquiries.int�������� = chkStrike.Value
            
        Case ģ���.�⹺���
            '�������
            If chkClass.Value = 1 Then
                If txtClass.Tag = "" Then
                    MsgBox "��ѡ��Ҫ��ѯ�ķ�����Ϣ��", vbInformation, gstrSysName
                    Me.txtClass.SetFocus
                    Exit Sub
                End If
            End If
            If chkJiXin.Value = 1 Then
                If txtJiXing.Tag = "" Then
                    MsgBox "��ѡ��Ҫ��ѯ�ļ�����Ϣ��", vbInformation, gstrSysName
                    Me.txtJiXing.SetFocus
                    Exit Sub
                End If
            End If
            If ChkҩƷ.Value = 1 Then
                If TxtҩƷ.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
                    Me.TxtҩƷ.SetFocus
                    Exit Sub
                End If
            End If
            If chk��Ӧ��.Value = 1 Then
                If txt��Ӧ��.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
                    Me.txt��Ӧ��.SetFocus
                    Exit Sub
                End If
            End If
            If Chk������.Value = 1 Then
                If txt������.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��������Ϣ��", vbInformation, gstrSysName
                    Me.txt������.SetFocus
                    Exit Sub
                End If
            End If
            
            If chk�ѱ��.Value = 1 And chkδ���.Value = 0 Then
                SQLCondition.intδ��� = 0
                SQLCondition.int�ѱ�� = 1
            ElseIf chkδ���.Value = 1 And chk�ѱ��.Value = 0 Then
                SQLCondition.intδ��� = 1
                SQLCondition.int�ѱ�� = 0
            End If
            
            SQLCondition.int�������һ����ѯ = 0
            If chk����.Value = 1 And chk���.Value = 1 Then SQLCondition.int�������һ����ѯ = 1
            
            If chk�з�Ʊ.Value = 1 And chk�޷�Ʊ.Value = 0 Then
                SQLCondition.int�з�Ʊ = 1
                SQLCondition.int�޷�Ʊ = 0
            ElseIf chk�޷�Ʊ.Value = 1 And chk�з�Ʊ.Value = 0 Then
                SQLCondition.int�з�Ʊ = 0
                SQLCondition.int�޷�Ʊ = 1
            End If
                
            If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
                txt��ʼNo.Text = zlCommFun.GetFullNo(txt��ʼNo.Text, intNO, lng�ⷿID)
            End If
            If Len(txt����No) < 8 And Len(txt����No) > 0 Then
                txt����No.Text = zlCommFun.GetFullNo(txt����No.Text, intNO, lng�ⷿID)
            End If
            
            SQLCondition.strNO��ʼ = Me.txt��ʼNo
            SQLCondition.strNO���� = Me.txt����No
            TemporaryInquiries.int�������� = chkStrike.Value
            TemporaryInquiries.intδ������� = chkAcc.Value
            TemporaryInquiries.int�Ѳ������ = chkAccStrike.Value
            TemporaryInquiries.intδ����˿� = chkNOVerifyBack.Value
            TemporaryInquiries.int������˿� = chkYesVerifyBack.Value
            
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            '�������
            If chkClass.Value = 1 Then
                If txtClass.Tag = "" Then
                    MsgBox "��ѡ��Ҫ��ѯ�ķ�����Ϣ��", vbInformation, gstrSysName
                    Me.txtClass.SetFocus
                    Exit Sub
                End If
            End If
            If chkJiXin.Value = 1 Then
                If txtJiXing.Tag = "" Then
                    MsgBox "��ѡ��Ҫ��ѯ�ļ�����Ϣ��", vbInformation, gstrSysName
                    Me.txtJiXing.SetFocus
                    Exit Sub
                End If
            End If
            If ChkҩƷ.Value = 1 Then
                If TxtҩƷ.Tag = 0 Then
                    MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
                    Me.TxtҩƷ.SetFocus
                    Exit Sub
                End If
            End If
            
            '������ѯ����
            SQLCondition.int�������һ����ѯ = 0
            If chk����.Value = 1 And chk���.Value = 1 Then SQLCondition.int�������һ����ѯ = 1
            
            If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
                txt��ʼNo.Text = zlCommFun.GetFullNo(txt��ʼNo.Text, intNO, lng�ⷿID)
            End If
            If Len(txt����No) < 8 And Len(txt����No) > 0 Then
                txt����No.Text = zlCommFun.GetFullNo(txt����No.Text, intNO, lng�ⷿID)
            End If
            
            SQLCondition.strNO��ʼ = Me.txt��ʼNo
            SQLCondition.strNO���� = Me.txt����No
            TemporaryInquiries.intδ��˳��� = chkNoStrike.Value
            TemporaryInquiries.int����˳��� = chkYesStrike.Value
            TemporaryInquiries.int�������� = chkStrike.Value
            
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                If chk�Ѵ�ӡ.Value = 1 And chkδ��ӡ.Value = 0 Then
                    SQLCondition.intδ��ӡ = 0
                    SQLCondition.int�Ѵ�ӡ = 1
                ElseIf chkδ��ӡ.Value = 1 And chk�Ѵ�ӡ.Value = 0 Then
                    SQLCondition.intδ��ӡ = 1
                    SQLCondition.int�Ѵ�ӡ = 0
                End If
            End If
    End Select
    
    '�ù��̷�Χѡ���ģ�鶼�е����
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(DTP��ʼʱ��(mint����), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(DTP����ʱ��(mint����), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date���ʱ�俪ʼ = CDate(Format(DTP��ʼʱ��(mint���), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date���ʱ����� = CDate(Format(DTP����ʱ��(mint���), "yyyy-mm-dd") & " 23:59:59")
    TemporaryInquiries.intδ��˵��� = chk����.Value
    TemporaryInquiries.int����˵��� = chk���.Value
    
    '==================��������ѡ�====================
    '��չ��ѯ����
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    Select Case mlngMode
        Case ģ���.��������
            SQLCondition.lng��Ӧ�� = IIf(chk��Ӧ��.Value = 1, Val(txt��Ӧ��.Tag), 0)
            
        Case ģ���.ҩƷ�ƻ�
            SQLCondition.str������ = IIf(Me.txt������ = "", "", Me.txt������ & "%")
            SQLCondition.lng�ƻ����� = IIf(Chk�ƻ�����.Value = 1, Cbo�ƻ�����.ListIndex + 1, 0)
            SQLCondition.lng���Ʒ��� = IIf(chk���Ʒ���.Value = 1, cbo���Ʒ���.ListIndex + 1, 0)
            
        Case ģ���.�������
            SQLCondition.str���� = IIf(Chk������.Value = 1, txt������, "")
            SQLCondition.lng������ = IIf(Chk���.Value = 1, Cbo���.ItemData(Cbo���.ListIndex), 0)
            
        Case ģ���.�⹺���
            SQLCondition.lngҩƷ���� = 0
            SQLCondition.str���� = ""
            
            SQLCondition.lngҩƷ���� = IIf(chkClass.Value = 1, Val(txtClass.Tag), 0)
            SQLCondition.str���� = IIf(chkJiXin.Value = 1, txtJiXing.Tag, "")
            If chk��Ʊ����.Value = 1 Then
                SQLCondition.date��Ʊ������ڿ�ʼ = CDate(Format(dtpStart��Ʊ.Value, "yyyy-mm-dd") & " 00:00:00")
                SQLCondition.date��Ʊ������ڽ��� = CDate(Format(dtpEnd��Ʊ.Value, "yyyy-mm-dd") & " 23:59:59")
            End If
            
            SQLCondition.lng������ = IIf(chk��Ӧ��.Value = 1, Val(txt��Ӧ��.Tag), 0)
            SQLCondition.str���� = IIf(Chk������.Value = 1, txt������, "")
            SQLCondition.str��Ʊ�ſ�ʼ = Me.txt��ʼ��Ʊ��
            SQLCondition.str��Ʊ�Ž��� = Me.txt������Ʊ��
            
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            SQLCondition.lngҩƷ���� = 0
            SQLCondition.str���� = ""
            
            SQLCondition.lngҩƷ���� = IIf(chkClass.Value = 1, Val(txtClass.Tag), 0)
            SQLCondition.str���� = IIf(chkJiXin.Value = 1, txtJiXing.Tag, "")
            If cbo�ⷿ.Visible Then SQLCondition.lng�ⷿ = IIf(Chk����ⷿ.Value = 1, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), 0)
            
    End Select
    '�ù��̸�������ѡ���ģ�鶼�е����
    SQLCondition.lngҩƷ = IIf(ChkҩƷ.Value = 1, Val(TxtҩƷ.Tag), 0)
    SQLCondition.str����� = IIf(Me.Txt����� = "", "", Me.Txt����� & "%")
    SQLCondition.str������ = IIf(Me.Txt������ = "", "", Me.Txt������ & "%")
    
    Unload Me
End Sub

Private Sub Cmd������_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt������.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� as id ,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ���� "
    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt������.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt������.SetFocus
    txt������.Tag = 1
    txt������.Text = rsProvider!����
    
    If mlngMode = ģ���.������� Then
        If Chk���.Visible = True Then
            If Chk���.Value = 1 Then
                Cbo���.SetFocus
            Else
                Chk���.SetFocus
            End If
        End If
    Else '�⹺
        chk��Ʊ����.SetFocus
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    Dim strModeName As String
    
    
    strModeName = Switch(mlngMode = ģ���.�⹺���, "ҩƷ�⹺������", mlngMode = ģ���.�������, "ҩƷ����������", mlngMode = _
        ģ���.�������, "ҩƷ����������", mlngMode = ģ���.��۵���, "ҩƷ�ƿ����", mlngMode = ģ���.ҩƷ�ƿ�, "ҩƷ�ƿ����", mlngMode = _
        ģ���.ҩƷ����, "ҩƷ�ƿ����", mlngMode = ģ���.��������, "ҩƷ�ƿ����", mlngMode = ģ���.ҩƷ�ƻ�, "ҩƷ�ƻ�����", mlngMode = _
        ģ���.��������, "ҩƷ��������")
    
    Select Case mlngMode
        Case ģ���.�⹺���, ģ���.�������, ģ���.�������, ģ���.ҩƷ�ƻ�
            Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������, ģ���.��������
            Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
    End Select
    
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.�������� Then
        If chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            chk��Ӧ��.SetFocus
        End If
    ElseIf mlngMode = ģ���.������� Then
        If Chk������.Value = 1 Then
            txt������.SetFocus
        Else
            Chk������.SetFocus
        End If
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
        If Chk����ⷿ.Value = 1 Then
            cbo�ⷿ.SetFocus
        Else
            Chk����ⷿ.SetFocus
        End If
    ElseIf mlngMode = ģ���.ҩƷ�ƻ� Or mlngMode = ģ���.������� Then
        Txt������.SetFocus
    End If
End Sub

Private Sub dtpEnd��Ʊ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub dtpStart��Ʊ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.dtpEnd��Ʊ.SetFocus
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.DTP����ʱ��(Index).SetFocus
End Sub

Private Sub Form_Activate()
    If mlngMode = ģ���.�⹺��� Then
        SQLCondition.intδ��� = 0
        SQLCondition.int�ѱ�� = 0
        SQLCondition.int�޷�Ʊ = 0
        SQLCondition.int�з�Ʊ = 0
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Then
        SQLCondition.intδ��ӡ = 0
        SQLCondition.int�Ѵ�ӡ = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    LoadSstFilter��Χ
    LoadSstFilter��������
    LoadData '��ʼ��
    
End Sub

Private Sub LoadData()
    Dim StrToday As String
    '���ܣ���������
    
    '�ָ���һ�ε�����
    '�����̸�ģ�鶼���ڵ����
    Me.DTP����ʱ��(mint����) = SQLCondition.date����ʱ�����
    Me.DTP����ʱ��(mint���) = SQLCondition.date���ʱ�����
    Me.DTP��ʼʱ��(mint����) = SQLCondition.date����ʱ�俪ʼ
    Me.DTP��ʼʱ��(mint���) = SQLCondition.date���ʱ�俪ʼ
    Me.chk����.Value = TemporaryInquiries.intδ��˵���
    Me.chk���.Value = TemporaryInquiries.int����˵���
    sstFilter.Tab = 0
    mblnAdvance = False
    
    Select Case mlngMode
        Case ģ���.��������
            TxtҩƷ.Tag = 0
            txt��Ӧ��.Tag = 0
        Case ģ���.ҩƷ�ƻ�
            Me.DTP����ʱ��(mint����) = SQLCondition.date����ʱ�����
            Me.DTP��ʼʱ��(mint����) = SQLCondition.date����ʱ�俪ʼ
            
            Me.chk����.Value = TemporaryInquiries.int�Ѹ��˵���
            
            SQLCondition.lngҩƷ = 0
        Case ģ���.�������, ģ���.�������
            Me.chkStrike.Value = TemporaryInquiries.int��������
            
            Me.TxtҩƷ.Tag = 0
            If mlngMode = ģ���.������� Then Me.txt������.Tag = 0
            
        Case ģ���.�⹺���
            chkStrike.Value = TemporaryInquiries.int��������
            chkAcc.Value = TemporaryInquiries.intδ�������
            chkAccStrike.Value = TemporaryInquiries.int�Ѳ������
            chk�ѱ��.Value = SQLCondition.int�ѱ��
            chkδ���.Value = SQLCondition.intδ���
            chk�з�Ʊ.Value = SQLCondition.int�з�Ʊ
            chk�޷�Ʊ.Value = SQLCondition.int�޷�Ʊ
            chkNOVerifyBack.Value = TemporaryInquiries.intδ����˿�
            chkYesVerifyBack.Value = TemporaryInquiries.int������˿�
            
            Me.txt��Ӧ��.Tag = 0
            Me.TxtҩƷ.Tag = 0
            Me.txt������.Tag = 0
            
            chk�ѱ��.Enabled = IIf(TemporaryInquiries.int����˵��� = 1, True, False)
            chkδ���.Enabled = IIf(TemporaryInquiries.int����˵��� = 1, True, False)
            mstrMatch = IIf(zlDatabase.GetPara("����ƥ��", , , 0) = "0", "%", "")
            
            StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
            dtpStart��Ʊ.Value = DateAdd("m", -1, CDate(StrToday))
            dtpEnd��Ʊ.Value = CDate(StrToday)
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            chkNoStrike.Value = TemporaryInquiries.intδ��˳���
            chkYesStrike.Value = TemporaryInquiries.int����˳���
            chkStrike.Value = TemporaryInquiries.int��������
            
            mblnStock = Check�Ƿ���ҩ����Ա
            mstrMatch = IIf(zlDatabase.GetPara("����ƥ��", , , 0) = "0", "%", "")
            
            Me.TxtҩƷ.Tag = 0
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                mint�������� = Val(zlDatabase.GetPara("��������", glngSys, ģ���.ҩƷ�ƿ�))
                If mint������� = -1 Then
                    Chk����ⷿ.Caption = "����ⷿ"
                Else
                    Chk����ⷿ.Caption = "�Ƴ��ⷿ"
                End If
                If mint�������� = 0 Then    '����Ҫ����
                    chkStrike.Visible = True
                    chkNoStrike.Visible = False
                    chkYesStrike.Visible = False
                Else
                    chkStrike.Visible = False
                    chkNoStrike.Visible = True
                    chkYesStrike.Visible = True
                End If
                
                chk�Ѵ�ӡ.Value = SQLCondition.int�Ѵ�ӡ
                chkδ��ӡ.Value = SQLCondition.intδ��ӡ
            ElseIf mlngMode = ģ���.ҩƷ���� Then
                    Chk����ⷿ.Caption = "���ò���"
            ElseIf mlngMode = ģ���.�������� Then
                    Chk����ⷿ.Caption = "������"
            End If
    End Select
End Sub

Private Sub LoadSstFilter��Χ()
    '���ܣ����÷�Χѡ�����Ŀؼ���ʾ��λ�ü���С��
    
    'ĬȻ����������С
    frmSearch.Height = 4510: sstFilter.Height = 3850: fra��Χ.Height = 2850: fra��������.Height = 2850
    Select Case mlngMode
        Case ģ���.�⹺���
            '�⹺���е���ʾ
            chkAcc.Visible = True
            chkAccStrike.Visible = True
            If gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                chk�ѱ��.Visible = True
                chkδ���.Visible = True
            Else
                chk�з�Ʊ.Top = chk�ѱ��.Top
                chk�޷�Ʊ.Top = chk�з�Ʊ.Top
            End If
            chk�з�Ʊ.Visible = True
            chk�޷�Ʊ.Visible = True
            chkNOVerifyBack.Visible = True
            chkYesVerifyBack.Visible = True
            '����߶�5500��ѡ��ҳ�߶�4810��fra��Χ�߶�4050
            frmSearch.Height = 5800: sstFilter.Height = 5110: fra��Χ.Height = 4150: fra��������.Height = 4150
            '����ȡ����ťCancel
'            cmdȡ��.Cancel = False
            
        Case ģ���.ҩƷ�ƿ�
            '�ƿ���������ɼ�
            mint�������� = Val(zlDatabase.GetPara("��������", glngSys, ģ���.ҩƷ�ƿ�))
            chkNoStrike.Visible = True
            chkYesStrike.Visible = True
            chkStrike.Visible = False
            chkδ��ӡ.Visible = True
            chk�Ѵ�ӡ.Visible = True
            
            frmSearch.Height = 4810: sstFilter.Height = 4150: fra��Χ.Height = 3150: fra��������.Height = 3150
         Case ģ���.ҩƷ�ƻ�
            '�ƻ��ĸ���
            chk����.Visible = True
            lblʱ��(mint����).Visible = True
            DTP��ʼʱ��(mint����).Visible = True
            lbl��(mint����).Visible = True
            DTP����ʱ��(mint����).Visible = True
            chkStrike.Visible = False
            '����߶�5500��ѡ��ҳ�߶�4810��fra��Χ�߶�4050
            frmSearch.Height = 5150: sstFilter.Height = 4450: fra��Χ.Height = 3450: fra��������.Height = 3450
        Case ģ���.��������
            '�������������No�ͳ�������
            lblNO.Visible = False
            txt��ʼNo.Visible = False
            lbl��(mintNo).Visible = False
            txt����No.Visible = False
            chkStrike.Visible = False
            '�ı�Caption
            chk����.Caption = "δ������"
            lblʱ��(mint����).Caption = "�Ǽ�����"
            chk���.Caption = "�Ѵ�����"
            lblʱ��(mint���).Caption = "��������"
            'No�������أ��ı���ʾ�ؼ���top
            chk����.Top = chk����.Top - 240
            lblʱ��(mint����).Top = lblʱ��(mint����).Top - 240
            DTP��ʼʱ��(mint����).Top = DTP��ʼʱ��(mint����).Top - 240
            lbl��(mint����).Top = lbl��(mint����).Top - 240
            DTP����ʱ��(mint����).Top = DTP����ʱ��(mint����).Top - 240
            chk���.Top = chk���.Top - 240
            lblʱ��(mint���).Top = lblʱ��(mint���).Top - 240
            DTP��ʼʱ��(mint���).Top = DTP��ʼʱ��(mint���).Top - 240
            lbl��(mint���).Top = lbl��(mint���).Top - 240
            DTP����ʱ��(mint���).Top = DTP����ʱ��(mint���).Top - 240
            '����߶�5500��ѡ��ҳ�߶�4810��fra��Χ�߶�4050
            frmSearch.Height = 4250: sstFilter.Height = 3550: fra��Χ.Height = 2550: fra��������.Height = 2250
    End Select
End Sub

Private Sub LoadSstFilter��������()
    '���ܣ����ø�������ѡ�����Ŀؼ���ʾ��λ�ü���С��
    
    Select Case mlngMode
        Case ģ���.�⹺���
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            chk��Ӧ��.Visible = True: txt��Ӧ��.Visible = True: cmd��Ӧ��.Visible = True
            Chk������.Visible = True: txt������.Visible = True: Cmd������.Visible = True
            chk��Ʊ����.Visible = True: dtpStart��Ʊ.Visible = True: lbl��(mint��Ʊ����).Visible = True: dtpEnd��Ʊ.Visible = True
            Lbl��Ʊ��.Visible = True: txt��ʼ��Ʊ��.Visible = True: lbl��(mint��Ʊ��).Visible = True: txt������Ʊ��.Visible = True
        Case ģ���.�������
            ChkҩƷ.Top = 480: TxtҩƷ.Top = 480: CmdҩƷ.Top = 480
            Lbl������.Top = 1200: Txt������.Top = 1140: Lbl������.Left = 930:  Txt������.Left = 1650
            Lbl�����.Top = 1800: Txt�����.Top = 1740: Lbl�����.Left = Lbl������.Left: Txt�����.Left = Txt������.Left
         Case ģ���.�������
            Chk������.Visible = True: txt������.Visible = True: Cmd������.Visible = True
            Chk���.Visible = True: Cbo���.Visible = True
            ChkҩƷ.Top = 360: TxtҩƷ.Top = 360: CmdҩƷ.Top = 360
            Chk������.Top = 950: txt������.Top = 950: Cmd������.Top = 950
            Chk���.Top = 1540: Cbo���.Top = 1540
            Lbl������.Top = 2190: Lbl������.Left = Lbl������.Left - 100: Txt������.Top = 2150: Txt������.Left = Cbo���.Left
            Lbl�����.Top = 2190: Txt�����.Top = 2150
        Case ģ���.ҩƷ�ƿ�
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk����ⷿ.Visible = True: cbo�ⷿ.Visible = True
            Lbl������.Top = 2000: Txt������.Top = 1940: Lbl������.Left = 930:  Txt������.Left = 1650
            Lbl�����.Top = 2400: Txt�����.Top = 2340: Lbl�����.Left = Lbl������.Left: Txt�����.Left = Txt������.Left
        Case ģ���.ҩƷ����
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk����ⷿ.Visible = True: cbo�ⷿ.Visible = True
            Lbl������.Top = 2000: Txt������.Top = 1940: Lbl������.Left = 930:  Txt������.Left = 1650
            Lbl�����.Top = 2400: Txt�����.Top = 2340: Lbl�����.Left = Lbl������.Left: Txt�����.Left = Txt������.Left
            Chk����ⷿ.Caption = "���ò���"
        Case ģ���.��������
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk����ⷿ.Visible = True: cbo�ⷿ.Visible = True
            Lbl������.Top = 2000: Txt������.Top = 1940: Lbl������.Left = 930:  Txt������.Left = 1650
            Lbl�����.Top = 2400: Txt�����.Top = 2340: Lbl�����.Left = Lbl������.Left: Txt�����.Left = Txt������.Left
            Chk����ⷿ.Caption = "������"
        Case ģ���.��������
            chk��Ӧ��.Visible = True: txt��Ӧ��.Visible = True: cmd��Ӧ��.Visible = True
            chk��Ӧ��.Caption = "��ҩ��λ": Lbl������.Caption = "�Ǽ���": Lbl�����.Caption = "������"
            ChkҩƷ.Top = 360: TxtҩƷ.Top = 360: CmdҩƷ.Top = 360
            chk��Ӧ��.Top = 750: txt��Ӧ��.Top = 750: cmd��Ӧ��.Top = 750
            Lbl������.Top = 1400: Txt������.Top = 1340: Lbl������.Left = 1050:  Txt������.Left = 1650
            Lbl�����.Top = 1800: Txt�����.Top = 1740: Lbl�����.Left = Lbl������.Left: Txt�����.Left = Txt������.Left
        Case ģ���.ҩƷ�ƻ�
            Chk�ƻ�����.Visible = True: Cbo�ƻ�����.Visible = True
            chk���Ʒ���.Visible = True: cbo���Ʒ���.Visible = True
            lbl������.Visible = True: txt������.Visible = True
            Lbl������.Top = 1700: Txt������.Top = 1640: Lbl������.Left = 870:  Txt������.Left = 1650
            Lbl�����.Top = 2100: Txt�����.Top = 2040: Lbl�����.Left = Lbl������.Left:   Txt�����.Left = Txt������.Left
            lbl������.Top = 2500: txt������.Top = 2440: lbl������.Left = Lbl������.Left: txt������.Left = Txt������.Left
    End Select
End Sub

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByRef colParameter As Collection) As Boolean
    Dim lngloop As Long
    
    GetSearch = False
    mblnCancel = False
    mlngMode = lngMode
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.������� Then mstrSelectTag = ""
    Set mfrmMain = FrmMain
    
    getParameterValue colParameter '��¼���ϴ������Ĳ���ֵ
    If mlngMode = ģ���.������� Or mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.������� Then If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    
    If mblnCancel = True Then Exit Function '���ȡ�����ý�ѡ���������¼������
    setParameterValue colParameter '��ѡ���������¼�������д���ȥ
    
    GetSearch = True
End Function

Private Sub setParameterValue(ByRef colParameter As Collection)
    '���ܣ�������ѡ���������¼�������д��ص�������,��ж����ص�ģ�����
    
    '�����̸�ģ�鶼���ڵ����
    
    CollectionModify colParameter, TemporaryInquiries.intδ��˵���, frmSearch.getKey_δ��˵���: TemporaryInquiries.intδ��˵��� = 0
    CollectionModify colParameter, SQLCondition.date����ʱ�俪ʼ, frmSearch.getKey_����ʱ�俪ʼ: SQLCondition.date����ʱ�俪ʼ = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.date����ʱ�����, frmSearch.getKey_����ʱ�����: SQLCondition.date����ʱ����� = CDate("00:00:00")
    CollectionModify colParameter, TemporaryInquiries.int����˵���, frmSearch.getKey_����˵���: TemporaryInquiries.int����˵��� = 0
    CollectionModify colParameter, SQLCondition.date���ʱ�俪ʼ, frmSearch.getKey_���ʱ�俪ʼ: SQLCondition.date���ʱ�俪ʼ = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.date���ʱ�����, frmSearch.getKey_���ʱ�����: SQLCondition.date���ʱ����� = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.lngҩƷ, frmSearch.getKey_ҩƷID: SQLCondition.lngҩƷ = 0
    CollectionModify colParameter, SQLCondition.str������, frmSearch.getKey_������: SQLCondition.str������ = ""
    CollectionModify colParameter, SQLCondition.str�����, frmSearch.getKey_�����: SQLCondition.str����� = ""
    
    Select Case mlngMode
        Case ģ���.��������
            CollectionModify colParameter, SQLCondition.lng��Ӧ��, frmSearch.getKey_��Ӧ��: SQLCondition.lng��Ӧ�� = 0
            
        Case ģ���.ҩƷ�ƻ�
            CollectionModify colParameter, SQLCondition.strNO��ʼ, frmSearch.getKey_NO��ʼ: SQLCondition.strNO��ʼ = ""
            CollectionModify colParameter, SQLCondition.strNO����, frmSearch.getKey_NO����: SQLCondition.strNO���� = ""
            CollectionModify colParameter, TemporaryInquiries.int�Ѹ��˵���, frmSearch.getKey_�Ѹ��˵���: TemporaryInquiries.int�Ѹ��˵��� = 0
            CollectionModify colParameter, SQLCondition.date����ʱ�俪ʼ, frmSearch.getKey_����ʱ�俪ʼ: SQLCondition.date����ʱ�俪ʼ = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.date����ʱ�����, frmSearch.getKey_����ʱ�����: SQLCondition.date����ʱ����� = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.lng�ƻ�����, frmSearch.getKey_�ƻ�����: SQLCondition.lng�ƻ����� = 0
            CollectionModify colParameter, SQLCondition.lng���Ʒ���, frmSearch.getKey_���Ʒ���: SQLCondition.lng���Ʒ��� = 0
            CollectionModify colParameter, SQLCondition.str������, frmSearch.getKey_������: SQLCondition.str������ = ""
            
        Case ģ���.�������, ģ���.�������
            CollectionModify colParameter, SQLCondition.strNO��ʼ, frmSearch.getKey_NO��ʼ: SQLCondition.strNO��ʼ = ""
            CollectionModify colParameter, SQLCondition.strNO����, frmSearch.getKey_NO����: SQLCondition.strNO���� = ""
            CollectionModify colParameter, TemporaryInquiries.int��������, frmSearch.getKey_��������: TemporaryInquiries.int�������� = 0
            If mlngMode = ģ���.������� Then
                CollectionModify colParameter, SQLCondition.str����, frmSearch.getKey_����: SQLCondition.str���� = ""
                CollectionModify colParameter, SQLCondition.lng������, frmSearch.getKey_������: SQLCondition.lng������ = 0
            End If
            
        Case ģ���.�⹺���
            CollectionModify colParameter, SQLCondition.strNO��ʼ, frmSearch.getKey_NO��ʼ: SQLCondition.strNO��ʼ = ""
            CollectionModify colParameter, SQLCondition.strNO����, frmSearch.getKey_NO����: SQLCondition.strNO���� = ""
            CollectionModify colParameter, TemporaryInquiries.int��������, frmSearch.getKey_��������: TemporaryInquiries.int�������� = 0
            CollectionModify colParameter, TemporaryInquiries.intδ�������, frmSearch.getKey_δ�������: TemporaryInquiries.intδ������� = 0
            CollectionModify colParameter, TemporaryInquiries.int�Ѳ������, frmSearch.getKey_�Ѳ������: TemporaryInquiries.int�Ѳ������ = 0
            CollectionModify colParameter, SQLCondition.intδ���, frmSearch.getKey_�ޱ��: SQLCondition.intδ��� = 0
            CollectionModify colParameter, SQLCondition.int�ѱ��, frmSearch.getKey_�б��: SQLCondition.int�ѱ�� = 0
            CollectionModify colParameter, SQLCondition.int�޷�Ʊ, frmSearch.getKey_�޷�Ʊ: SQLCondition.int�޷�Ʊ = 0
            CollectionModify colParameter, SQLCondition.int�з�Ʊ, frmSearch.getKey_�з�Ʊ: SQLCondition.int�з�Ʊ = 0
            CollectionModify colParameter, TemporaryInquiries.intδ����˿�, frmSearch.getKey_δ����˿�: TemporaryInquiries.intδ����˿� = 0
            CollectionModify colParameter, TemporaryInquiries.int������˿�, frmSearch.getKey_������˿�: TemporaryInquiries.int������˿� = 0
            CollectionModify colParameter, SQLCondition.int�������һ����ѯ, frmSearch.getKey_�������һ����ѯ: SQLCondition.int�������һ����ѯ = 0
            CollectionModify colParameter, SQLCondition.lngҩƷ����, frmSearch.getKey_ҩƷ����: SQLCondition.lngҩƷ���� = 0
            CollectionModify colParameter, SQLCondition.str����, frmSearch.getKey_����: SQLCondition.str���� = ""
            CollectionModify colParameter, SQLCondition.lng������, frmSearch.getKey_��Ӧ��: SQLCondition.lng������ = 0
            CollectionModify colParameter, SQLCondition.str����, frmSearch.getKey_����: SQLCondition.str���� = ""
            CollectionModify colParameter, SQLCondition.date��Ʊ������ڿ�ʼ, frmSearch.getKey_��Ʊ������ڿ�ʼ: SQLCondition.date��Ʊ������ڿ�ʼ = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.date��Ʊ������ڽ���, frmSearch.getKey_��Ʊ������ڽ���: SQLCondition.date��Ʊ������ڽ��� = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.str��Ʊ�ſ�ʼ, frmSearch.getKey_��Ʊ�ſ�ʼ: SQLCondition.str��Ʊ�ſ�ʼ = ""
            CollectionModify colParameter, SQLCondition.str��Ʊ�Ž���, frmSearch.getKey_��Ʊ�Ž���: SQLCondition.str��Ʊ�Ž��� = ""
            
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            CollectionModify colParameter, SQLCondition.strNO��ʼ, frmSearch.getKey_NO��ʼ: SQLCondition.strNO��ʼ = ""
            CollectionModify colParameter, SQLCondition.strNO����, frmSearch.getKey_NO����: SQLCondition.strNO���� = ""
            CollectionModify colParameter, TemporaryInquiries.intδ��˳���, frmSearch.getKey_δ��˳���: TemporaryInquiries.intδ��˳��� = 0
            CollectionModify colParameter, TemporaryInquiries.int����˳���, frmSearch.getKey_����˳���: TemporaryInquiries.int����˳��� = 0
            CollectionModify colParameter, TemporaryInquiries.int��������, frmSearch.getKey_��������: TemporaryInquiries.int�������� = 0
            CollectionModify colParameter, SQLCondition.int�������һ����ѯ, frmSearch.getKey_�������һ����ѯ: SQLCondition.int�������һ����ѯ = 0
            CollectionModify colParameter, SQLCondition.lngҩƷ����, frmSearch.getKey_ҩƷ����: SQLCondition.lngҩƷ���� = 0
            CollectionModify colParameter, SQLCondition.str����, frmSearch.getKey_����: SQLCondition.str���� = ""
            CollectionModify colParameter, SQLCondition.lng�ⷿ, frmSearch.getKey_�ⷿID: SQLCondition.lng�ⷿ = 0
            
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                CollectionModify colParameter, SQLCondition.intδ��ӡ, frmSearch.getKey_δ��ӡ: SQLCondition.intδ��ӡ = 0
                CollectionModify colParameter, SQLCondition.int�Ѵ�ӡ, frmSearch.getKey_�Ѵ�ӡ: SQLCondition.int�Ѵ�ӡ = 0
            End If
    End Select
End Sub

Private Sub CollectionModify(ByRef colParameter As Collection, ByVal varConditionn As Variant, ByVal strConditionnKey As String)
    '���ܣ������޸�ָ��keyֵ��value
    colParameter.Remove strConditionnKey
    colParameter.Add varConditionn, strConditionnKey
End Sub

Private Sub getParameterValue(ByVal colParameter As Collection)
    '���ܣ��������崫�����Ĳ�����ֵ���ô����Ӧ�ı����������ݳ�ʼ��
    
    '��ʱ��ѯ��ʼ��
    '�����̸�ģ�鶼���ڵ����
    
    SQLCondition.date����ʱ�俪ʼ = colParameter(frmSearch.getKey_����ʱ�俪ʼ)
    SQLCondition.date����ʱ����� = colParameter(frmSearch.getKey_����ʱ�����)
    SQLCondition.date���ʱ�俪ʼ = colParameter(frmSearch.getKey_���ʱ�俪ʼ)
    SQLCondition.date���ʱ����� = colParameter(frmSearch.getKey_���ʱ�����)
    TemporaryInquiries.intδ��˵��� = colParameter(frmSearch.getKey_δ��˵���)
    TemporaryInquiries.int����˵��� = colParameter(frmSearch.getKey_����˵���)
    
    Select Case mlngMode
        Case ģ���.ҩƷ�ƻ�
            SQLCondition.date����ʱ�俪ʼ = colParameter(frmSearch.getKey_����ʱ�俪ʼ)
            SQLCondition.date����ʱ����� = colParameter(frmSearch.getKey_����ʱ�����)
            TemporaryInquiries.int�Ѹ��˵��� = colParameter(frmSearch.getKey_�Ѹ��˵���)
            
        Case ģ���.�������, ģ���.�������
            TemporaryInquiries.int�������� = colParameter(frmSearch.getKey_��������)
            
        Case ģ���.�⹺���
            TemporaryInquiries.int�������� = colParameter(frmSearch.getKey_��������)
            TemporaryInquiries.intδ������� = colParameter(frmSearch.getKey_δ�������)
            TemporaryInquiries.int�Ѳ������ = colParameter(frmSearch.getKey_�Ѳ������)
            SQLCondition.int�ѱ�� = colParameter(frmSearch.getKey_�б��)
            SQLCondition.intδ��� = colParameter(frmSearch.getKey_�ޱ��)
            SQLCondition.int�޷�Ʊ = colParameter(frmSearch.getKey_�޷�Ʊ)
            SQLCondition.int�з�Ʊ = colParameter(frmSearch.getKey_�з�Ʊ)
            TemporaryInquiries.intδ����˿� = colParameter(frmSearch.getKey_δ����˿�)
            TemporaryInquiries.int������˿� = colParameter(frmSearch.getKey_������˿�)
            
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            TemporaryInquiries.intδ��˳��� = colParameter(frmSearch.getKey_δ��˳���)
            TemporaryInquiries.int����˳��� = colParameter(frmSearch.getKey_����˳���)
            TemporaryInquiries.int�������� = colParameter(frmSearch.getKey_��������)
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                SQLCondition.int�Ѵ�ӡ = colParameter(frmSearch.getKey_�Ѵ�ӡ)
                SQLCondition.intδ��ӡ = colParameter(frmSearch.getKey_δ��ӡ)
            End If
    End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If tvw���.Visible = True Then
        tvw���.Visible = False
        txtClass.SetFocus
        Cancel = True
        Exit Sub
    End If
    If lvw����.Visible = True Then
        lvw����.Visible = False
        txtJiXing.SetFocus
        Cancel = True
        Exit Sub
    End If
        
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Maker"
                txt������.SetFocus
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
            Case "Booker"
                Txt������.SetFocus
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
            Case "Verify"
                Txt�����.SetFocus
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
            Case "Checker"
                txt������.SetFocus
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
        End Select
        Cancel = True
        Exit Sub
    End If
    
    If Not mfrmMain Is Nothing Then
        Set mfrmMain = Nothing
    End If
    
    Call ReleaseSelectorRS
End Sub

Private Sub lvw����_DblClick()
    Dim i As Integer
    Dim strName As String
    
    With lvw����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked = True Then
                strName = strName & .ListItems(i).Text & ","
            End If
        Next
        lvw����.Visible = False
        txtJiXing.Tag = strName
        txtJiXing.Text = strName
    End With
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
        If ChkҩƷ.Value = 1 Then
            TxtҩƷ.SetFocus
        Else
            ChkҩƷ.SetFocus
        End If
    End If
End Sub

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvw����_DblClick
End Sub

Private Sub lvw����_LostFocus()
    lvw����.Visible = False
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Maker"
                    txt������.Text = .TextMatrix(.Row, 1)
                    txt������.Tag = 1
                    Chk���.SetFocus
                Case "Booker"
                    Txt������ = .TextMatrix(.Row, 2)
                    Txt�����.SetFocus
                Case "Verify"
                    Txt����� = .TextMatrix(.Row, 2)
                    cmdȷ��.SetFocus
                    If mlngMode = ģ���.ҩƷ�ƻ� Then txt������.SetFocus
                    If mlngMode = ģ���.�⹺��� Then txt��ʼ��Ʊ��.SetFocus
                Case "Checker"
                    txt������ = .TextMatrix(.Row, 2)
                    cmdȷ��.SetFocus
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsDepartment As New Recordset
    Dim strStock As String
    Dim strվ������ As String
    
    On Error GoTo errHandle
    strվ������ = GetDeptStationNode(mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
            If cbo�ⷿ.ListCount < 1 Then
                Select Case mlngMode
                    Case ģ���.ҩƷ�ƻ�
                        If Cbo�ƻ�����.ListCount < 1 Then
                            With Cbo�ƻ�����
                                .Clear
                                .AddItem "�¶ȼƻ�", 0
                                .AddItem "���ȼƻ�", 1
                                .AddItem "��ȼƻ�", 2
                                .AddItem "�ܼƻ�", 3
                                .ListIndex = 0
                            End With
                            
                            With cbo���Ʒ���
                                .Clear
                                .AddItem "����ͬ�����β��շ�", 0
                                .AddItem "�ٽ��ڼ�ƽ�����շ�", 1
                                .AddItem "ҩƷ����������շ�", 2
                                .AddItem "ҩƷ�����������շ�", 3
                                .ListIndex = 0
                            End With
                        End If
                    
                    Case ģ���.ҩƷ�ƿ�
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                                & "Where " & IIf(strվ������ <> "", " (a.վ�� = [3] or a.վ�� is null) AND ", "") & " c.�������� = b.���� " _
                                & "  AND Instr([1],b.����,1) > 0 " _
                                & "  AND a.id = c.����id " _
                                & "  AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
                    Case ģ���.ҩƷ����
                        strStock = "O"
                        gstrSQL = " Select C.ID " & _
                            " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                            " Where " & IIf(strվ������ <> "", " (c.վ�� = [3] or c.վ�� is null) AND ", "") & " A.��������=B.���� And A.����ID=C.ID " & _
                            "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                            "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])"
                        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                            & "Where " & IIf(strվ������ <> "", " (a.վ�� = [3] or a.վ�� is null) AND ", "") & " c.�������� = b.���� " _
                            & "  AND Instr([1],b.����,1) > 0 " _
                            & "  AND a.id = c.����id " _
                            & "  AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')" _
                            & IIf(mblnStock, "", " And a.ID IN (Select Distinct ���ò���ID From ҩƷ���ÿ��� Where ���ò���ID IN (" & gstrSQL & "))")
                    Case ģ���.��������
                       gstrSQL = "SELECT b.Id,b.���� " _
                               & "FROM ҩƷ�������� A, ҩƷ������ B " _
                               & "Where A.���id = B.ID AND A.���� = 11 "
                    Case ģ���.��۵���, ģ���.ҩƷ�̵�
                        If Chk����ⷿ.Visible = True Then
                            Chk����ⷿ.Visible = False
                            cbo�ⷿ.Visible = False
                        End If
                        Exit Sub
                End Select
                
                If mlngMode = ģ���.��۵��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Or mlngMode = ģ���.ҩƷ�̵� Then
                    Set rsDepartment = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock, UserInfo.�û�ID, gstrNodeNo)
                
                    With cbo�ⷿ
                        Do While Not rsDepartment.EOF
                            .AddItem rsDepartment.Fields(1)
                            .ItemData(.NewIndex) = rsDepartment.Fields(0)
                            rsDepartment.MoveNext
                        Loop
                        If .ListCount > 0 Then .ListIndex = 0
                    End With
                    rsDepartment.Close
                End If
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvw���_DblClick()
    With tvw���
        If .SelectedItem.Text <> "" Then
            If .SelectedItem.Key Like "Root*" Then Exit Sub
            txtClass.Tag = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "_") + 1)
            txtClass.Text = .SelectedItem.Text
            .Visible = False
        End If
    End With
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
        If chkJiXin.Value = 1 Then
            txtJiXing.SetFocus
        Else
            chkJiXin.SetFocus
        End If
    End If
End Sub

Private Sub tvw���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then tvw���_DblClick
End Sub

Private Sub tvw���_LostFocus()
    tvw���.Visible = False
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(Trim(txtClass.Text))
        If strTemp <> "" Then
            tvw���.Left = txtClass.Left
            tvw���.Top = txtClass.Top + txtClass.Height
            tvw���.Visible = True
            tvw���.SetFocus
            
            gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
                      "Where Instr([1], ����, 1) > 0 " & _
                      "Order by ���� "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
            With tvw���
                .Nodes.Clear
                Do While Not rsTmp.EOF
                    Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!����, rsTmp!����, 2, 2)
                    nodTmp.Tag = "Root" & rsTmp!����
                    rsTmp.MoveNext
                Loop
                rsTmp.Close
            End With
            
            gstrSQL = "Select ID, �ϼ�id, ����, 1 As ĩ��, ����, ����" & _
                        " From (Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, ����" & _
                               " From ���Ʒ���Ŀ¼" & _
                               " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                     " (���� Like [1] Or ���� Like [1] Or ���� Like [1])" & _
                               " Start With �ϼ�id Is Null" & _
                               " Connect By Prior ID = �ϼ�id" & _
                               " Union " & _
                               " Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, ����" & _
                               " From ���Ʒ���Ŀ¼" & _
                               " Where ID In (Select �ϼ�id" & _
                                            " From ���Ʒ���Ŀ¼" & _
                                            " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                                  " (���� Like [1] Or ���� Like [1] Or ���� Like [1])))" & _
                        " Start With �ϼ�id Is Null" & _
                        " Connect By Prior ID = �ϼ�id" & _
                        " Order By Level, ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & strTemp & mstrMatch)
            
            With rsTmp
                If .EOF Then
                    Exit Sub
                End If
                
                '��ҩƷ��;��������װ��
                Do While Not .EOF
                    Intĩ�� = IIf(!ĩ�� = 1, 3, 2)
                    If IsNull(!�ϼ�ID) Then
                        Set nodTmp = tvw���.Nodes.Add("Root" & !����, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
                    Else
                        Set nodTmp = tvw���.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
                    End If
                    nodTmp.Tag = !����   '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
                    .MoveNext
                Loop
            End With
        
            With tvw���
                .Nodes(1).Selected = True
                If .Nodes(1).Children <> 0 Then
                    Intĩ�� = 1
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(2).Children <> 0 Then
                    Intĩ�� = 2
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(3).Children <> 0 Then
                    Intĩ�� = 3
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                Else
                    Intĩ�� = 0
                    .Nodes(1).Selected = True
                    .SelectedItem.Selected = True
                End If
                If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
            End With
        Else
            If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
                If chkJiXin.Value = 1 Then
                    txtJiXing.SetFocus
                Else
                    chkJiXin.SetFocus
                End If
            End If
        End If
        
    ElseIf KeyCode = vbKeyDelete Then
        txtClass.Tag = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtJiXing_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿID As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind <> "" Then
            lvw����.Left = txtJiXing.Left
            lvw����.Top = txtJiXing.Top + txtJiXing.Height
            lvw����.Visible = True
            lvw����.SetFocus
            
            On Error GoTo errHandle
            lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            If lng�ⷿID <> 0 Then
                '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
                gstrSQL = "Select Distinct J.����,J.���� " & _
                          "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                          "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] and (j.���� like [2] or j.���� like [2] or j.���� like [2]) " & _
                          "Order by J.���� "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿID, "%" & strFind & mstrMatch)
            Else
                gstrSQL = "Select ����,���� From ҩƷ���� where ���� like [1] or ���� like [1] or ���� like [1] order by ���� "
                Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����", "%" & strFind & mstrMatch)
            End If
            
            With rsTmp
                If .RecordCount = 0 Then
                    lvw����.Visible = False
                    MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                     txtJiXing.SetFocus: Exit Sub
                End If
                lvw����.ListItems.Clear
                Do While Not .EOF
                    lvw����.ListItems.Add , "K" & !����, !����, 1, 1
                    .MoveNext
                Loop
            End With
        Else
            If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
                If ChkҩƷ.Value = 1 Then
                    TxtҩƷ.SetFocus
                Else
                    ChkҩƷ.SetFocus
                End If
            End If
        End If
        
    ElseIf KeyCode = vbKeyDelete Then
        txtJiXing.Tag = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(txt������.Text) = "" Then
            If mlngMode = ģ���.�⹺��� Then
                txt��ʼ��Ʊ��.SetFocus
            Else
                Me.cmdȷ��.SetFocus
            End If
            Exit Sub
        End If
        txt������.Text = UCase(txt������.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.txt������ & "%", _
                        Me.txt������ & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Checker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Left = sstFilter.Left + fra��������.Left + txt������.Left
                    .Height = txt������.Top - sstFilter.Top - fra��������.Top - 50
                    .Top = sstFilter.Top + fra��������.Top + txt������.Top - .Height
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt������ = IIf(IsNull(!����), "", !����)
                If mlngMode = ģ���.�⹺��� Then
                    txt��ʼ��Ʊ��.SetFocus
                Else
                    Me.cmdȷ��.SetFocus
                End If
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt��Ӧ��)) <> "" Then
        txt��Ӧ�� = UCase(txt��Ӧ��)
        vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd)
        
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) " & _
                  "  And ĩ��=1 And substr(����,1,1)=1 " & _
                  "  And (���� like [1] Or ���� like [1] or ���� like [1] Or zlSpellCode(����) Like [2] Or zlWbCode(����) Like [2])" & _
                  "  Start with �ϼ�ID is null and (վ�� = [3] Or վ�� is Null) connect by prior ID =�ϼ�ID and (վ�� = [3] Or վ�� is Null) "
        Set RecTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%", txt��Ӧ�� & "%", gstrNodeNo)
        
        
        If blnCancel Then txt��Ӧ��.SetFocus: Exit Sub
        
        If RecTmp.State = 0 Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            KeyCode = 0
            txt��Ӧ��.Tag = 0
            txt��Ӧ��.SelStart = 0
            txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
            Exit Sub
        End If
        
        txt��Ӧ�� = RecTmp!����
        txt��Ӧ��.Tag = RecTmp!id
                  
    End If
    
    If mlngMode = ģ���.�������� Then
        Txt������.SetFocus
    ElseIf mlngMode = ģ���.�⹺��� Then
        If Chk������.Value = 1 Then
            txt������.SetFocus
        Else
            Chk������.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������Ʊ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdȷ��.SetFocus
End Sub

Private Sub txt��ʼNo_GotFocus()
    If sstFilter.Tab = 1 Then sstFilter.Tab = 0
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '��ʼ׼��
    intNO = Switch(mlngMode = ģ���.�⹺���, 21, mlngMode = ģ���.�������, 22, mlngMode = _
        ģ���.�������, 24, mlngMode = ģ���.��۵���, 25, mlngMode = ģ���.ҩƷ�ƿ�, 26, mlngMode = _
        ģ���.ҩƷ����, 27, mlngMode = ģ���.��������, 28, mlngMode = ģ���.ҩƷ�ƻ�, 32)
    
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNo(txt��ʼNo.Text, intNO, lng�ⷿID)
        End If
        txt����No.SetFocus
    End If
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '��ʼ׼��
    intNO = Switch(mlngMode = ģ���.�⹺���, 21, mlngMode = ģ���.�������, 22, mlngMode = _
        ģ���.�������, 24, mlngMode = ģ���.��۵���, 25, mlngMode = ģ���.ҩƷ�ƿ�, 26, mlngMode = _
        ģ���.ҩƷ����, 27, mlngMode = ģ���.��������, 28, mlngMode = ģ���.ҩƷ�ƻ�, 32)
    
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt����No) < 8 And Len(txt����No) > 0 Then
            txt����No.Text = zlCommFun.GetFullNo(txt����No.Text, intNO, lng�ⷿID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = 100
End Sub

Private Sub txtJiXing_GotFocus()
    txtJiXing.SelStart = 0
    txtJiXing.SelLength = 100
End Sub

Private Sub Txt��ʼ��Ʊ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            If mlngMode = ģ���.�⹺��� Then
                txt��ʼ��Ʊ��.SetFocus
            ElseIf mlngMode = ģ���.ҩƷ�ƻ� Then
                Me.txt������.SetFocus
            Else
                Me.cmdȷ��.SetFocus
            End If
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)
        
        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", _
                        Me.Txt����� & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt�����.Top + Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra��������.Top - Txt�����.Top - Txt�����.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt����� = IIf(IsNull(!����), "", !����)
                If mlngMode = ģ���.�⹺��� Then
                    txt��ʼ��Ʊ��.SetFocus
                ElseIf mlngMode = ģ���.ҩƷ�ƻ� Then
                    Me.txt������.SetFocus
                Else
                    Me.cmdȷ��.SetFocus
            End If
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub txt������_GotFocus()
    txt������.SelStart = 0
    txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
End Sub

Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt������.hWnd)
    
    On Error GoTo errHandle
    
    If KeyCode = vbKeyReturn Then
        If Trim(txt������) <> "" Then
            txt������ = UCase(txt������)
            
            gstrSQL = "Select ���� as id,����,���� From ҩƷ������ " & _
                      "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) " & _
              "Order By ����"
    
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ������", False, "", "", False, False, _
                    True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & Me.txt������ & "%", Me.txt������ & "%", gstrNodeNo)
            
            If blnCancel Then txt������.SetFocus: Exit Sub
            
            If rsTemp.State = 0 Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                txt������.Tag = 0
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
                Exit Sub
            End If
            
            txt������ = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            txt������.Tag = 1
        End If
        
        If mlngMode = ģ���.������� Then
            If Chk���.Visible = True Then
                If Chk���.Value = 1 Then
                    Cbo���.SetFocus
                Else
                    Chk���.SetFocus
                End If
            End If
        Else '�⹺
            chk��Ʊ����.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            Txt�����.SetFocus
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", _
                        Me.Txt������ & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt������.Top + Txt������.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt������.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra��������.Top - Txt������.Top - Txt������.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt������ = IIf(IsNull(!����), "", !����)
                Me.Txt�����.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    Dim strModeName As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) <> "" Then
        sngLeft = Me.Left + sstFilter.Left + fra��������.Left + TxtҩƷ.Left
        sngTop = Me.Top + sstFilter.Top + fra��������.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
        If sngTop + 3630 > Screen.Height Then
            sngTop = sngTop - TxtҩƷ.Height - 3630
        End If
        
        strkey = Trim(TxtҩƷ.Text)
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        strModeName = Switch(mlngMode = ģ���.�⹺���, "ҩƷ�⹺������", mlngMode = ģ���.�������, "ҩƷ����������", mlngMode = _
            ģ���.�������, "ҩƷ����������", mlngMode = ģ���.��۵���, "ҩƷ�ƿ����", mlngMode = ģ���.ҩƷ�ƿ�, "ҩƷ�ƿ����", mlngMode = _
            ģ���.ҩƷ����, "ҩƷ�ƿ����", mlngMode = ģ���.��������, "ҩƷ�ƿ����", mlngMode = ģ���.ҩƷ�ƻ�, "ҩƷ�ƻ�����", mlngMode = _
            ģ���.��������, "ҩƷ��������")
        
        Select Case mlngMode
            Case ģ���.�⹺���, ģ���.�������, ģ���.�������, ģ���.ҩƷ�ƻ�
                Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
            Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������, ģ���.��������
                Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
        End Select
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then Exit Sub
        If gintҩƷ������ʾ = 1 Then
            TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        TxtҩƷ.Tag = RecReturn!ҩƷid
    End If
    
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.�������� Then
        If chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            chk��Ӧ��.SetFocus
        End If
    ElseIf mlngMode = ģ���.������� Then
        If Chk������.Value = 1 Then
            txt������.SetFocus
        Else
            Chk������.SetFocus
        End If
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Then
        If Chk����ⷿ.Value = 1 Then
            cbo�ⷿ.SetFocus
        Else
            Chk����ⷿ.SetFocus
        End If
    ElseIf mlngMode = ģ���.ҩƷ�ƻ� Or mlngMode = ģ���.������� Then
        Txt������.SetFocus
    End If
    
End Sub

Private Sub TxtҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Cbo�ⷿ_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0

    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Cbo�ⷿ_Validate(Cancel As Boolean)
    If cbo�ⷿ.ListCount > 0 Then
        If cbo�ⷿ.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Cbo���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtҩƷ_GotFocus()
    TxtҩƷ.SelStart = 0
    TxtҩƷ.SelLength = Len(TxtҩƷ.Text)
End Sub

Private Sub txt��Ӧ��_GotFocus()
    txt��Ӧ��.SelStart = 0
    txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
End Sub

Private Sub Txt������_GotFocus()
    Txt������.SelStart = 0
    Txt������.SelLength = Len(Txt������.Text)
End Sub

Private Sub Txt�����_GotFocus()
    Txt�����.SelStart = 0
    Txt�����.SelLength = Len(Txt�����.Text)
End Sub

Private Sub txt������_GotFocus()
    txt������.SelStart = 0
    txt������.SelLength = Len(txt������.Text)
End Sub

Private Sub Cbo�ƻ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cbo���Ʒ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Function Check�Ƿ���ҩ����Ա() As Boolean
    Dim rsDepend As ADODB.Recordset
    
    On Error GoTo errHandle
    '�ж��ǲ���ҩ����Աʹ�ñ�ģ��
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [2] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr('HIJKLMN', b.����, 1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " _
            & "  And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) "
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.�û�ID, gstrNodeNo)
                  
    Check�Ƿ���ҩ����Ա = (rsDepend.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    
    If mlngMode = ģ���.�⹺��� Then
        gstrSQL = "Select id,�ϼ�ID,����,����,ĩ��,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And ���� is Not NULL " & _
              "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is NULL Connect by prior id=�ϼ�id"
        Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ӧ��", gstrNodeNo)
        With rsCompete
            If .EOF Then
                .Close
                MsgBox "ҩƷ��Ӧ����Ϣ��ȫ�����ڹ�ҩ��λ����������ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    gstrSQL = "Select ����,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null"
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "ҩƷ������", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "ҩƷ��������Ϣ��ȫ,�����ֵ����������ҩƷ��������Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If mlngMode = ģ���.������� Then
        gstrSQL = "SELECT B.Id,b.���� " _
                & "FROM ҩƷ�������� A, ҩƷ������ B " _
                & "Where A.���id = B.ID AND A.���� = 4 "
        Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsCompete
            If .EOF Then
                MsgBox "ҩƷ�������û��������Ӧ������������ҩƷ������࣡", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveFirst
            Do While Not .EOF
                Cbo���.AddItem .Fields(1)
                Cbo���.ItemData(Cbo���.NewIndex) = .Fields(0)
                .MoveNext
            Loop
            Cbo���.ListIndex = 0
            .Close
        End With
    End If
    
    CheckCompete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

