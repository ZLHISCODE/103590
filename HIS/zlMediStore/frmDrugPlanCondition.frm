VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugPlanCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmDrugPlanCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab sstConditon 
      Height          =   7680
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   13547
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "�ƻ�(&1)"
      TabPicture(0)   =   "frmDrugPlanCondition.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl�ⷿ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lvw����"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra�ƻ�����"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra��ʽ"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra�ƻ�����"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk�������ƻ�����"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbo�ⷿ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkBaseMedi"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra����ҩ"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkOnlyBaseMedi"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frm�������"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chk�������ֿ��"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Chk����ȡ��ȡ���޵�ҩƷ"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkClearZeroPlan"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chk�ο�����"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "fra��������"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Chk����"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "ҩƷ����(&2)"
      TabPicture(1)   =   "frmDrugPlanCondition.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "tvw��;"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "��Ӧ��(&3)"
      TabPicture(2)   =   "frmDrugPlanCondition.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tvw������λ"
      Tab(2).Control(1)=   "chk�б굥λ"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "��Դҩ��(&4)"
      TabPicture(3)   =   "frmDrugPlanCondition.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvwҩ��"
      Tab(3).Control(1)=   "Label2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "��Դ�ⷿ(&5)"
      TabPicture(4)   =   "frmDrugPlanCondition.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(2)=   "lvw�ⷿ"
      Tab(4).ControlCount=   3
      Begin VB.CheckBox chk�б굥λ 
         Caption         =   "���ϴι�Ӧ�����б굥λΪ׼(&W)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   -74880
         TabIndex        =   50
         Top             =   420
         Width           =   2985
      End
      Begin VB.CheckBox Chk���� 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   840
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2925
         Width           =   675
      End
      Begin VB.Frame fra�������� 
         Caption         =   " �������� "
         Enabled         =   0   'False
         Height          =   1710
         Left            =   3720
         TabIndex        =   38
         Top             =   4980
         Width           =   3495
         Begin VB.TextBox txt�������� 
            Height          =   300
            Left            =   1560
            TabIndex        =   40
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt�������� 
            Height          =   300
            Left            =   1560
            TabIndex        =   39
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lbl�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&X)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   390
            TabIndex        =   42
            Top             =   330
            Width           =   990
         End
         Begin VB.Label lbl�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   390
            TabIndex        =   41
            Top             =   720
            Width           =   990
         End
      End
      Begin VB.CheckBox chk�ο����� 
         Appearance      =   0  'Flat
         Caption         =   "�����������ƻ�����"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   1140
         Width           =   3120
      End
      Begin VB.CheckBox chkClearZeroPlan 
         Appearance      =   0  'Flat
         Caption         =   "�������ƻ�����Ϊ0��ҩƷ��¼"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   900
         Width           =   2760
      End
      Begin VB.CheckBox Chk����ȡ��ȡ���޵�ҩƷ 
         Appearance      =   0  'Flat
         Caption         =   "����ȡ�������޵�ҩƷ"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   660
         Width           =   2205
      End
      Begin VB.CheckBox chk�������ֿ�� 
         Appearance      =   0  'Flat
         Caption         =   "�������ֿ������"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Frame frm������� 
         Caption         =   " �������ѡ��"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   2220
         Width           =   7095
         Begin VB.CheckBox chk���� 
            Caption         =   "��ȡ��ͨҩ"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��ȡ����ҩ"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   32
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��ȡ������ҩ"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   31
            Top             =   280
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��ȡ����ҩ"
            Height          =   180
            Index           =   3
            Left            =   5040
            TabIndex        =   30
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkOnlyBaseMedi 
         Appearance      =   0  'Flat
         Caption         =   "��������ҩ��"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   28
         Top             =   900
         Width           =   1800
      End
      Begin VB.Frame fra����ҩ 
         Caption         =   " ����ҩѡ�� "
         Height          =   645
         Left            =   120
         TabIndex        =   24
         Top             =   1500
         Width           =   7095
         Begin VB.OptionButton opt����ҩ 
            Caption         =   "����ȡ�ǳ���ҩ"
            Height          =   180
            Index           =   1
            Left            =   4440
            TabIndex        =   27
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton opt����ҩ 
            Caption         =   "����ȡ����ҩ"
            Height          =   180
            Index           =   0
            Left            =   2400
            TabIndex        =   26
            Top             =   300
            Width           =   1575
         End
         Begin VB.OptionButton opt����ҩ 
            Caption         =   "�������Ƿ񳣱�ҩ"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   300
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkBaseMedi 
         Appearance      =   0  'Flat
         Caption         =   "��������ҩ��"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   23
         Top             =   900
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   276
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   330
         Width           =   3210
      End
      Begin VB.CheckBox Chk�������ƻ����� 
         Appearance      =   0  'Flat
         Caption         =   "�������ƻ�����"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   660
         Width           =   1560
      End
      Begin VB.Frame fra�ƻ����� 
         Caption         =   " ���Ʒ��� "
         Height          =   1710
         Left            =   120
         TabIndex        =   9
         Top             =   4980
         Width           =   3435
         Begin VB.OptionButton opt���� 
            Caption         =   "�Զ���������շ�(&5)"
            Height          =   195
            Index           =   4
            Left            =   735
            TabIndex        =   17
            Top             =   1440
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ҩƷ�ճ��������շ�(&4)"
            Height          =   195
            Index           =   3
            Left            =   735
            TabIndex        =   13
            Top             =   1170
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ҩƷ����������շ�(&3)"
            Height          =   195
            Index           =   2
            Left            =   735
            TabIndex        =   12
            Top             =   885
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�ٽ��ڼ�ƽ�����շ�(&2)"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   11
            Top             =   585
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����ͬ�����Բ��շ�(&1)"
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin VB.Frame fra��ʽ 
         Caption         =   " ����������ʽ"
         Height          =   1710
         Left            =   3720
         TabIndex        =   43
         Top             =   4980
         Width           =   3495
         Begin VB.OptionButton opt���� 
            Caption         =   "��������޲����ƻ�����"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��������޲����ƻ�����"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Frame fra�ƻ����� 
         Caption         =   " �ƻ����� "
         Height          =   765
         Left            =   120
         TabIndex        =   4
         Top             =   6780
         Width           =   7095
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "�¶ȼƻ�(&A)"
            Height          =   210
            Index           =   0
            Left            =   1845
            TabIndex        =   8
            Top             =   405
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "���ȼƻ�(&B)"
            Height          =   210
            Index           =   1
            Left            =   3555
            TabIndex        =   7
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "��ȼƻ�(&C)"
            Height          =   210
            Index           =   2
            Left            =   5130
            TabIndex        =   6
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "�ܼƻ�(&W)"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   405
            Width           =   1290
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   " �Զ�������"
         Height          =   765
         Left            =   120
         TabIndex        =   18
         Top             =   6780
         Visible         =   0   'False
         Width           =   7095
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   303759363
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Left            =   3225
            TabIndex        =   20
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   303759363
            CurrentDate     =   36263
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   360
            TabIndex        =   22
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2985
            TabIndex        =   21
            Top             =   420
            Width           =   180
         End
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   1680
         Left            =   120
         TabIndex        =   47
         Top             =   3180
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   2963
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw��; 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   49
         Top             =   660
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12065
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvw������λ 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   51
         Top             =   660
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12065
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvw�ⷿ 
         Height          =   5880
         Left            =   -74880
         TabIndex        =   56
         Top             =   900
         Width           =   7068
         _ExtentX        =   12462
         _ExtentY        =   10372
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvwҩ�� 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   52
         Top             =   660
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   12065
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(��������ѡ���ⷿΪ[ȫԺ]ʱĬ��ͳ�����пⷿ��棬����Ϊ��ǰ�ⷿ���)"
         Height          =   180
         Left            =   -74760
         TabIndex        =   58
         Top             =   650
         Width           =   6144
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ���пⷿ��Ϊͳ�ƿ�������Ŀⷿ"
         Height          =   180
         Left            =   -74760
         TabIndex        =   57
         Top             =   420
         Width           =   3060
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ�����Ĭ����ȡ���м���ҩƷ"
         Height          =   180
         Left            =   1680
         TabIndex        =   55
         Top             =   2940
         Width           =   2700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�����ѡ����࣬�򲻲����ƻ�����"
         Height          =   180
         Left            =   -74760
         TabIndex        =   54
         Top             =   420
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ����ҩ����Ϊͳ��ҩƷ���۵ķ�ҩҩ����������ѡʱĬ��Ϊ����ҩ��"
         Height          =   180
         Left            =   -74760
         TabIndex        =   53
         Top             =   420
         Width           =   5760
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   2940
         Width           =   630
      End
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         Caption         =   "�ⷿ(&K)"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   1
      Top             =   8040
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1800
      Top             =   7920
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
            Picture         =   "frmDrugPlanCondition.frx":0098
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":0F72
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":13C4
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":1816
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   8040
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugPlanCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mstr��;ID As String
Private mstr���� As String
Private mlng�ⷿID As Long
Private mint�ƻ����� As Integer
Private mint���Ʒ��� As Integer             '0-����ͬ�����Բ��շ���1-�ٽ��ڼ�ƽ�����շ���2-ҩƷ����������շ���3-ҩƷ�����������շ���4-�Զ���������շ�
Private mbln���� As Boolean
Private mint���� As Integer
Private mint���� As Integer
Private mbln��ҩ�ⷿ As Boolean                     '��ҩ�ⷿ
Private mfrmMain As Form
Private mbln�ƻ����� As Boolean
Private mstr������ID As String
Private mbln�б굥λ As Boolean
Private mstrBeginDate As String
Private mstrEndDate As String
Private mbln�����ǿ�� As Boolean
Private mblnClearZeroPlan As Boolean
Private mblnBaseMedi As Boolean
Private mblnOnlyBaseMedi As Boolean
Private mintStock As Integer            '����ҩѡ��0-ֻ��ȡ����ҩ��1-ֻ��ȡ�ǳ���ҩ��2-�������Ƿ񳣱�ҩ��
Private mbln������ʽ As Boolean         ' false-���޷�ʽ true-���޷�ʽ
Private mintPlanPoint As Integer        'ȫԺ�ƻ�����վ�� 0-Ҫ��վ�㣬1-����վ��
Private mstrToxicologyClass As String       '�������
Private mbln�����������ƻ� As Boolean   '�����������ƻ�����
Private mstr��Դҩ�� As String               '��ʽ:ҩ��id1,ҩ��id2...
Private mstr��Դ�ⷿ As String               '��ʽ:ҩ��id1,ҩ��id2...
Private mstrAll��Դҩ�� As String       '������Դҩ������ʽ:ҩ��id1,ҩ��id2...
Private mstrAll��Դ�ⷿ As String       '������Դҩ������ʽ:ҩ��id1,ҩ��id2...

Private Enum zlDrugPlan
    P0_����ͬ�����Բ��շ� = 0
    P1_�ٽ��ڼ�ƽ�����շ� = 1
    P2_ҩƷ����������շ� = 2
    P3_ҩƷ�����������շ� = 3
    P4_�Զ���������շ� = 4
End Enum
Public Function GetCondition(FrmMain As Form, ByRef str��;ID, ByRef str���� As String, _
    ByRef lng�ⷿID As Long, ByRef int�ƻ����� As Integer, ByRef int���Ʒ��� As Integer, _
    ByRef bln���� As Boolean, ByRef int���� As Integer, ByRef int���� As Integer, ByRef bln�ƻ����� As Boolean, _
    ByRef str������ID As String, ByRef bln�б굥λ As Boolean, ByRef strBeginDate As String, ByRef strEndDate As String, _
    ByRef bln�����ǿ�� As Boolean, ByRef blnClearZeroPlan As Boolean, ByRef blnBaseMedi As Boolean, ByRef intStock As Integer, _
    ByRef bln������ʽ As Boolean, ByRef blnOnlyBaseMedi As Boolean, ByRef strToxicologyClass As String, ByRef bln�����������ƻ� As Boolean, _
    ByRef str��Դҩ�� As String, ByRef str��Դ�ⷿ As String, Optional ByRef strAll��Դҩ�� As String, Optional ByRef strAll��Դ�ⷿ As String) As Boolean

    mstr��;ID = ""
    mstr���� = ""
    mlng�ⷿID = 0
    mint�ƻ����� = 0
    mint���Ʒ��� = 0
    mblnSelect = False
    mblnClearZeroPlan = False
    mblnBaseMedi = False
    mintStock = 0
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    bln�б굥λ = mbln�б굥λ
    str������ID = mstr������ID
    
    str��;ID = mstr��;ID
    str���� = mstr����
    lng�ⷿID = mlng�ⷿID
    int�ƻ����� = mint�ƻ�����
    int���Ʒ��� = mint���Ʒ��� + 1
    bln���� = mbln����
    int���� = mint����
    int���� = mint����
    bln�ƻ����� = mbln�ƻ�����
    strBeginDate = mstrBeginDate
    strEndDate = mstrEndDate
    bln�����ǿ�� = mbln�����ǿ��
    blnClearZeroPlan = mblnClearZeroPlan
    blnBaseMedi = mblnBaseMedi
    intStock = mintStock
    bln������ʽ = mbln������ʽ
    blnOnlyBaseMedi = mblnOnlyBaseMedi
    strToxicologyClass = mstrToxicologyClass
    bln�����������ƻ� = mbln�����������ƻ�
    str��Դҩ�� = mstr��Դҩ��
    str��Դ�ⷿ = mstr��Դ�ⷿ
    strAll��Դҩ�� = mstrAll��Դҩ��
    strAll��Դ�ⷿ = mstrAll��Դ�ⷿ
End Function


Private Sub cmdCancel_Click()
    mblnSelect = False
    Unload Me
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim Str�ڼ� As String
    Dim intMonth As Integer
    Dim i As Integer
    Dim bln���� As Boolean
        
    If opt����(3).Value Then
        '���������������С�ڿ����������
        '���������������������������Ϊ��
        If Trim(txt��������.Text) = "" Then
            MsgBox "������������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Trim(txt��������.Text) = "" Then
            MsgBox "������������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt��������.Text) Then
            MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt��������.Text) Then
            MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Val(txt��������.Text) <= 0 Then
            MsgBox "���������������С���㣡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Val(txt��������.Text) <= 0 Then
            MsgBox "���������������С���㣡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Val(txt��������.Text) < Val(txt��������.Text) Then
            MsgBox "���������������С�ڿ������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        If Val(txt��������.Text) > 300 Then
            MsgBox "��������������ܴ���300�죡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Sub
        End If
        mint���� = Val(txt��������.Text)
        mint���� = Val(txt��������.Text)
    End If

    mstr��;ID = ""
    For intItem = 1 To tvw��;.Nodes.count
        If tvw��;.Nodes(intItem).Key = "Root" And tvw��;.Nodes(intItem).Checked Then
            mstr��;ID = "���з���"
            Exit For
        End If
        
        If tvw��;.Nodes(intItem).Key <> "Root" And _
            tvw��;.Nodes(intItem).Key <> "_�г�ҩ" And _
            tvw��;.Nodes(intItem).Key <> "_�в�ҩ" And _
            tvw��;.Nodes(intItem).Key <> "_����ҩ" And _
            tvw��;.Nodes(intItem).Checked Then
            mstr��;ID = mstr��;ID & "," & Mid(tvw��;.Nodes(intItem).Key, 2)
        End If
    Next
    
    If mstr��;ID <> "" And mstr��;ID <> "���з���" Then
        mstr��;ID = Mid(mstr��;ID, 2)
    End If
    
    mstr���� = ""
    
    intItems = Me.Lvw����.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If Lvw����.ListItems(intItem).Checked Then
                mstr���� = mstr���� & "," & "'" & Lvw����.ListItems(intItem).Text & "'"
            End If
        Next
    End If
    If mstr���� <> "" Then mstr���� = Mid(mstr����, 2)
    
    mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    
    frmDrugPlanCard.LblTitle.Tag = cbo�ⷿ.Text

    For intItem = 0 To opt�ƻ�.count - 1
       If opt�ƻ�(intItem).Value Then
           frmDrugPlanCard.txt�ƻ�����.Caption = Mid(opt�ƻ�(intItem).Caption, 1, InStr(1, opt�ƻ�(intItem).Caption, "(") - 1)
           mint�ƻ����� = intItem + 1
           Exit For
       End If
    Next

    For intItem = 0 To opt����.count - 1
       If opt����(intItem).Value Then
           frmDrugPlanCard.txt���Ʒ���.Caption = Mid(opt����(intItem).Caption, 1, InStr(1, opt����(intItem).Caption, "(") - 1)
           mint���Ʒ��� = intItem
           Exit For
       End If
    Next
    
    mstr������ID = ""
    For i = 1 To tvw������λ.Nodes.count
        If tvw������λ.Nodes(i).Key <> "Root" And _
            tvw������λ.Nodes(i).Checked Then
            If tvw������λ.Nodes(i).Tag = "1" Then
                mstr������ID = mstr������ID & "," & Mid(tvw������λ.Nodes(i).Key, 2)
            End If
        End If
    Next
    If mstr������ID <> "" Then mstr������ID = Mid(mstr������ID, 2)
    mbln�б굥λ = chk�б굥λ.Value = 1
    
    mbln���� = (Chk����ȡ��ȡ���޵�ҩƷ.Value = 1)
    mbln�ƻ����� = (Chk�������ƻ�����.Value <> 1)
    mbln�����ǿ�� = (chk�������ֿ��.Value = 1)
    mblnClearZeroPlan = (chkClearZeroPlan.Value = 1)
    mblnBaseMedi = (chkBaseMedi.Value = 1)
    mblnOnlyBaseMedi = (chkOnlyBaseMedi.Value = 1)
    mintStock = IIf(opt����ҩ(0).Value = True, 0, IIf(opt����ҩ(1).Value = True, 1, 2))
    mbln�����������ƻ� = (chk�ο�����.Value = 1)
    
    If mint���Ʒ��� = zlDrugPlan.P2_ҩƷ����������շ� Or mint���Ʒ��� = zlDrugPlan.P4_�Զ���������շ� Then
        mstrBeginDate = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd")
        mstrEndDate = Format(dtp����ʱ��.Value, "yyyy-mm-dd")
    End If
    
    If opt����.Value = True Then
        mbln������ʽ = False
    Else
        mbln������ʽ = True
    End If
    
    For i = 0 To chk����.count - 1
        If chk����(i).Value = 0 Then
            bln���� = True
            Exit For
        End If
    Next
    
    mstrToxicologyClass = ""
    If bln���� = True Then
        If chk����(0).Value = 1 Then
            mstrToxicologyClass = " t.�������='��ͨҩ'"
        End If
        If chk����(1).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = " t.�������='����ҩ'"
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.�������='����ҩ'"
            End If
        End If
        If chk����(2).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = "  t.������� ='����I��' or t.������� ='����II��' "
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.������� ='����I��' or t.������� ='����II��'"
            End If
        End If
        If chk����(3).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = " t.������� ='����ҩ'"
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.������� ='����ҩ'"
            End If
        End If
        
        If mstrToxicologyClass <> "" Then
            mstrToxicologyClass = "(" & mstrToxicologyClass & ")"
        End If
    End If
    
    mstr��Դҩ�� = ""
    intItems = Me.lvwҩ��.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If lvwҩ��.ListItems(intItem).Checked Then
                mstr��Դҩ�� = IIf(mstr��Դҩ�� = "", "", mstr��Դҩ�� & ",") & Mid(lvwҩ��.ListItems(intItem).Key, 2)
            End If
        Next
    End If
    
    mstr��Դ�ⷿ = ""
    intItems = Me.lvw�ⷿ.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If lvw�ⷿ.ListItems(intItem).Checked Then
                mstr��Դ�ⷿ = IIf(mstr��Դ�ⷿ = "", "", mstr��Դ�ⷿ & ",") & Mid(lvw�ⷿ.ListItems(intItem).Key, 2)
            End If
        Next
    End If
    
    If mstr��Դ�ⷿ <> "" Then
        If InStr(1, "," & mstr��Դ�ⷿ & ",", "," & cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex) & ",") = 0 Then
            mstr��Դ�ⷿ = mstr��Դ�ⷿ & "," & cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
        End If
    End If
    
    '���û��ѡ��ҩƷ���࣬��ʾ�Ƿ����
    If mstr��;ID = "" Then
        If MsgBox("δѡ��ҩƷ���࣬�������յļƻ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub cbo�ⷿ_Click()
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    '�����ȫԺ�ƻ�����ȡ����ҩƷ����
    On Error GoTo errHandle
    If Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.ListIndex) = 0 Then
        gstrSQL = "Select ����,���� From ҩƷ���� Order by ����"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ����")
        
'        opt����ҩ(0).Enabled = False
'        opt����ҩ(1).Enabled = False
    Else
        '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
        mbln��ҩ�ⷿ = False
        gstrSQL = "Select 1 From ��������˵�� " & _
                 " Where �������� Like '��ҩ%' And ����ID=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
        
        If Not rsTemp.EOF Then mbln��ҩ�ⷿ = True
    
        gstrSQL = "Select Distinct J.����,J.���� " & _
                 " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
                 " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
                 " And A.ִ�п���ID=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    End If

    Lvw����.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!���� = "����")
            End If
            Lvw����.ListItems.Add , "K" & !����, !����, , 1
            .MoveNext
        Loop
        If mbln��ҩ�ⷿ And blnEXIST = False Then
            Lvw����.ListItems.Add , "KK1", "����", , 1
        End If
    End With
    If Chk����.Value <> 2 Then
        Chk����_Click
    Else
        Chk����.Value = 0
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node
    Dim i As Integer
    Dim blnSelectStock As String
    Dim strIco As String, strID As String
    Dim strTemp As String
    Dim objItem As ListItem
    
    On Error GoTo errH

    blnSelectStock = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƻ�����", "�ⷿ", "0")
    mintPlanPoint = Val(zlDataBase.GetPara("ȫԺ�ƻ�����վ��", glngSys, 1330, 0))
    
    sstConditon.Tab = 0
    
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With

    If zlStr.IsHavePrivs(gstrprivs, "���пⷿ") Then
        If blnSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
    
    '��;
    gstrSQL = "Select Level as ��,ID,�ϼ�ID,����,DECODE(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') As ���� " & _
        " From ���Ʒ���Ŀ¼" & _
        " Where ���� in (1,2,3)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Order by Level"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption)

    Set objNode = tvw��;.Nodes.Add(, , "Root", "������;", "Item")
    Set objNode = tvw��;.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", "Item")
    Set objNode = tvw��;.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", "Item")
    Set objNode = tvw��;.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", "Item")

    Do While Not rsTmp.EOF
        If rsTmp!�� = 1 Then
            Set objNode = tvw��;.Nodes.Add("_" & rsTmp!����, 4, "_" & rsTmp!Id, rsTmp!����, "Item")
        Else
            Set objNode = tvw��;.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!Id, rsTmp!����, "Item")
        End If
        rsTmp.MoveNext
    Loop
    tvw��;.Nodes("Root").Selected = True
    tvw��;.Nodes("Root").Expanded = True
    '0-Ҫ��վ�㣬1-����վ��
    mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    If mlng�ⷿID <> 0 Or (mlng�ⷿID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
        strTemp = "(վ�� = [1] Or վ�� is Null) And "
    End If
    gstrSQL = "" & _
        "   Select Level as ��,ID,�ϼ�ID,����||'-'||���� ����,ĩ�� " & _
        "   From ��Ӧ��" & _
        "   where " & strTemp & "(substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null)" & _
        "   Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        "   Order by Level"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ӧ��", gstrNodeNo)
    
    tvw������λ.Nodes.Clear
    Set objNode = tvw������λ.Nodes.Add(, , "Root", "����ҩƷ������", "Folder")
    objNode.Sorted = True
    Do While Not rsTmp.EOF
        strIco = IIf(Val(NVL(rsTmp!ĩ��)) = 1, "Card", "Folder")
        If rsTmp!�� = 1 Then
            Set objNode = tvw������λ.Nodes.Add("Root", 4, "_" & rsTmp!Id, rsTmp!����, strIco)
            strID = strID & rsTmp!Id & ";"
        Else
            If InStr(strID, rsTmp!Id & ";") = 0 Then
                Set objNode = tvw������λ.Nodes.Add("Root", 4, "_" & rsTmp!Id, rsTmp!����, strIco)
            Else
                Set objNode = tvw������λ.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!Id, rsTmp!����, strIco)
            End If
        End If
        If strIco = "Card" Then
            objNode.Tag = "1"
        End If
        objNode.Sorted = True
        rsTmp.MoveNext
    Loop
    tvw������λ.Nodes("Root").Selected = True
    tvw������λ.Nodes("Root").Expanded = True
    
    Me.dtp����ʱ�� = Sys.Currentdate
    Me.dtp��ʼʱ�� = DateAdd("m", -1, Me.dtp����ʱ��)
    fra��ʽ.Visible = False
    opt����.Value = True    'Ĭ���ǰ�����
    
    mint���Ʒ��� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�ƻ�����", "�������ñ༭��ʽ", 0))
    If mint���Ʒ��� >= 0 And mint���Ʒ��� <= 4 Then
        opt����(mint���Ʒ���).Value = True
    Else
        opt����(0).Value = True
    End If
    
    '��Դҩ��
    mstr��Դҩ�� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�ƻ�����", "��Դҩ��", "")
    gstrSQL = "Select Distinct a.id,a.����,a.���� " & _
        " From ���ű� a,��������˵�� b " & _
        " Where a.id=b.����id And b.��������  In ('��ҩ��','��ҩ��','��ҩ��') And TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'  Order By ���� "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩ��")
    
    lvwҩ��.ListItems.Clear
    mstrAll��Դҩ�� = ""
    With rsTmp
        Do While Not .EOF
            Set objItem = lvwҩ��.ListItems.Add(, "K" & !Id, "[" & !���� & "]" & !����, , 1)
                        
            If InStr(1, "," & mstr��Դҩ�� & ",", "," & !Id & ",") > 0 Then
                objItem.Checked = True
            End If
            
            mstrAll��Դҩ�� = IIf(mstrAll��Դҩ�� = "", "", mstrAll��Դҩ�� & ",") & !Id
            
            .MoveNext
        Loop
    End With
    
    '��Դ�ⷿ
    mstr��Դ�ⷿ = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�ƻ�����", "��Դ�ⷿ", "")
    gstrSQL = "Select Distinct a.id,a.����,a.���� " & _
        " From ���ű� a,��������˵�� b " & _
        " Where a.id=b.����id And b.��������  In ('��ҩ��','��ҩ��','��ҩ��','��ҩ��', '��ҩ��', '��ҩ��') And TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'  Order By ���� "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩ��")
    
    lvw�ⷿ.ListItems.Clear
    mstrAll��Դ�ⷿ = ""
    With rsTmp
        Do While Not .EOF
            Set objItem = lvw�ⷿ.ListItems.Add(, "K" & !Id, "[" & !���� & "]" & !����, , 1)
                        
            If InStr(1, "," & mstr��Դ�ⷿ & ",", "," & !Id & ",") > 0 Then
                objItem.Checked = True
            End If
            
            mstrAll��Դ�ⷿ = IIf(mstrAll��Դ�ⷿ = "", "", mstrAll��Դ�ⷿ & ",") & !Id
            
            .MoveNext
        Loop
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƻ�����", "�������ñ༭��ʽ", mint���Ʒ���)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƻ�����", "��Դҩ��", mstr��Դҩ��)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƻ�����", "��Դ�ⷿ", mstr��Դ�ⷿ)
End Sub


Private Sub opt����_Click(index As Integer)
    fra��������.Visible = False
    fra�ƻ�����.Visible = False
    fra����.Visible = False
    chk�ο�����.Enabled = True
    
    Select Case index
    Case zlDrugPlan.P0_����ͬ�����Բ��շ�, zlDrugPlan.P1_�ٽ��ڼ�ƽ�����շ�
        '0-����ͬ�����Բ��շ���1-�ٽ��ڼ�ƽ�����շ�
        fra�ƻ�����.Visible = True
        fra��������.Visible = True
        fra��������.Enabled = False
        fra��������.ZOrder 0
    Case zlDrugPlan.P2_ҩƷ����������շ�
        'ҩƷ����������շ�
        '��ʾ���䣬����������ʽ
        chk�ο�����.Value = 0
        chk�ο�����.Enabled = False
        fra����.Visible = True
        fra��ʽ.Visible = True
        fra��ʽ.ZOrder 0
    Case zlDrugPlan.P3_ҩƷ�����������շ�
        'ҩƷ�����������շ�
        fra�ƻ�����.Visible = True
        fra��������.Visible = True
        fra��������.Enabled = True
        fra��������.ZOrder 0
    Case zlDrugPlan.P4_�Զ���������շ�
        '�Զ���������շ�
        fra����.Visible = True
        fra��������.Visible = True
        fra��������.Enabled = False
        fra��������.ZOrder 0
    End Select
    
    mint���Ʒ��� = index

End Sub

Private Sub opt�ƻ�_Click(index As Integer)
    opt����(0).Enabled = True
    opt����(1).Enabled = True
    opt����(2).Enabled = True
    opt����(3).Enabled = True

    Select Case index
    Case 1
        If opt����(3).Value Then
            opt����(3).Value = False
            opt����(0).Value = True
        End If
        opt����(3).Enabled = False
    Case 2
        If opt����(0).Value Or opt����(3).Value Then
            opt����(0).Value = False
            opt����(3).Value = False
            opt����(1).Value = True
        End If
        opt����(0).Enabled = False
        opt����(3).Enabled = False
    Case 3
        If opt����(0).Value Then
            opt����(0).Value = False
            opt����(1).Value = True
        End If
        opt����(0).Enabled = False
    End Select
End Sub



Private Sub tvw������λ_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim blnAllUnCheck As Boolean
    
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked, False
    
    blnAllUnCheck = True
    Do While Not Node Is Nothing
        If Node.Checked = True Then
            blnAllUnCheck = False
            Exit Do
        End If
        Set Node = Node.Next
    Loop
    
    If blnAllUnCheck Then
        chk�б굥λ.Value = 0
        chk�б굥λ.Enabled = False
    ElseIf chk�б굥λ.Enabled = False Then
        chk�б굥λ.Enabled = True
    End If
End Sub

Private Sub tvw��;_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub

Private Sub SetParentNode1(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer

    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If tvw��;.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw��;.Nodes(intIdx).Next.index
            Loop
            If intIdx = Node.LastSibling.index Then
                If tvw��;.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If

        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck
        End If
    End If
End Sub


Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean, Optional blnTvw��; As Boolean = True)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If blnTvw��; = True Then
                    If tvw��;.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw��;.Nodes(intIdx).Next.index
                Else
                    If tvw������λ.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw������λ.Nodes(intIdx).Next.index
                End If
            Loop
            If intIdx = Node.LastSibling.index Then
                If blnTvw��; = True Then
                       If tvw��;.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                Else
                       If tvw������λ.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck, blnTvw��;
        End If
    End If
End Sub
Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw��;.Nodes.count
        If tvw��;.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function


Private Sub Chk����_Click()
    If Chk����.Value = 2 Then Exit Sub
    Call SetSelect(Lvw����, Chk����.Value)
End Sub

Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(Lvw����, Item)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            Chk����.Value = 1
        ElseIf intCount > 0 Then
            Chk����.Value = 2
        Else
            Chk����.Value = 0
        End If
    End With
End Sub
