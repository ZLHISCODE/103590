VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentifyNBYKT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����һ��ͨ���ʶ��"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   Icon            =   "frmIdentifyNBYKT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton opt 
      Caption         =   "���ũ����"
      Height          =   225
      Index           =   6
      Left            =   6900
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   210
      Width           =   1275
   End
   Begin VB.OptionButton opt 
      Caption         =   "���ҽ����"
      Height          =   225
      Index           =   5
      Left            =   5610
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      Caption         =   "ũ����"
      Height          =   225
      Index           =   4
      Left            =   4710
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.OptionButton opt 
      Caption         =   "ҽ����"
      Height          =   225
      Index           =   3
      Left            =   3750
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.CheckBox chk������ 
      Caption         =   "�¿�"
      Height          =   225
      Left            =   1140
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   675
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Index           =   0
      Left            =   5730
      TabIndex        =   64
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Index           =   1
      Left            =   6990
      TabIndex        =   65
      Top             =   6600
      Width           =   1100
   End
   Begin TabDlg.SSTab sstab 
      Height          =   4425
      Left            =   240
      TabIndex        =   10
      Top             =   2010
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   7805
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "������Ϣ"
      TabPicture(0)   =   "frmIdentifyNBYKT.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl�Ա�"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl��������"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl֤������"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl֤����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl����״��"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEMAIL"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl˵��"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl������"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblʡ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl��"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl�ֵ�"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl��ַ"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl�ʱ�"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl�绰"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl������λ"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl��λ��ַ"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl��λ�ʱ�"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl��λ�绰"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblְҵ"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl�ֻ���"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl��������"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl�����绰"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl������"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl����"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl���պ�"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt����"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cbo�Ա�"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "msk��������"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cbo֤������"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt֤����"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cbo����״��"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtEMAIL"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txt˵��"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt������"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Frame1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtʡ"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txt��"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txt�ֵ�"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt��ַ"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt�ʱ�"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt�绰"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt������λ"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt��λ��ַ"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt��λ�ʱ�"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt��λ�绰"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cboְҵ"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt�ֻ���"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt��������"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt�����绰"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt������"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt���պ�"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cbo������"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).ControlCount=   53
      Begin VB.ComboBox cbo������ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt���պ� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt������ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt�����绰 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   63
         Top             =   3930
         Width           =   1485
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   61
         Top             =   3930
         Width           =   1485
      End
      Begin VB.TextBox txt�ֻ��� 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   59
         Top             =   3930
         Width           =   1485
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   300
         Left            =   6330
         TabIndex        =   57
         Text            =   "cboְҵ"
         Top             =   3540
         Width           =   1515
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   55
         Top             =   3540
         Width           =   1485
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   53
         Top             =   3540
         Width           =   1485
      End
      Begin VB.TextBox txt��λ��ַ 
         Height          =   300
         Left            =   4770
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3150
         Width           =   3045
      End
      Begin VB.TextBox txt������λ 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   49
         Top             =   3150
         Width           =   2295
      End
      Begin VB.TextBox txt�绰 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   47
         Top             =   2760
         Width           =   1485
      End
      Begin VB.TextBox txt�ʱ� 
         Height          =   300
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt��ַ 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txt�ֵ� 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   41
         Top             =   2370
         Width           =   1485
      End
      Begin VB.TextBox txt�� 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   39
         Top             =   2370
         Width           =   1485
      End
      Begin VB.TextBox txtʡ 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   37
         Top             =   2370
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   60
         TabIndex        =   35
         Top             =   2070
         Width           =   7875
      End
      Begin VB.TextBox txt������ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   12
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtEMAIL 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1320
         Width           =   1515
      End
      Begin VB.ComboBox cbo����״�� 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txt֤���� 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   22
         Top             =   930
         Width           =   1515
      End
      Begin VB.ComboBox cbo֤������ 
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   930
         Width           =   1515
      End
      Begin MSMask.MaskEdBox msk�������� 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   1020
         TabIndex        =   18
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   540
         Width           =   1515
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   14
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lbl���պ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���պ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5700
         TabIndex        =   33
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   31
         Top             =   1770
         Width           =   360
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
         Left            =   390
         TabIndex        =   29
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label lbl�����绰 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����绰"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5520
         TabIndex        =   62
         Top             =   3990
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   60
         Top             =   3990
         Width           =   720
      End
      Begin VB.Label lbl�ֻ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ֻ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   58
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label lblְҵ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   56
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label lbl��λ�绰 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   54
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   52
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lbl��λ��ַ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��ַ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   50
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lbl������λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   48
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lbl�绰 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   46
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl�ʱ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʱ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3540
         TabIndex        =   44
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl��ַ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ַ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   570
         TabIndex        =   42
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl�ֵ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ֵ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   40
         Top             =   2430
         Width           =   360
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3330
         TabIndex        =   38
         Top             =   2430
         Width           =   180
      End
      Begin VB.Label lblʡ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʡ/��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   36
         Top             =   2430
         Width           =   450
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
         Left            =   390
         TabIndex        =   11
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   27
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label lblEMAIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   25
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label lbl����״�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   23
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl֤���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "֤����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5700
         TabIndex        =   21
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lbl֤������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "֤������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   19
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   17
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   13
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.OptionButton opt 
      Caption         =   "���֤"
      Height          =   225
      Index           =   2
      Left            =   2820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.OptionButton opt 
      Caption         =   "����"
      Height          =   225
      Index           =   1
      Left            =   2010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   675
   End
   Begin VB.OptionButton opt 
      Caption         =   "���￨"
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   960
      MaxLength       =   50
      TabIndex        =   9
      Top             =   570
      Width           =   4005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFact 
      Height          =   915
      Left            =   300
      TabIndex        =   67
      Top             =   960
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   1614
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmIdentifyNBYKT.frx":0028
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl��״̬ 
      AutoSize        =   -1  'True
      Caption         =   "��ʧע��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5100
      TabIndex        =   66
      Top             =   600
      Width           =   840
   End
   Begin VB.Label lbl��ˢ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ˢ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   630
      Width           =   540
   End
End
Attribute VB_Name = "frmIdentifyNBYKT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr�������� As String
Private mstr���� As String
Private mstrUser As String
Private mstrPwd As String
Private mbln��Ϣת�� As Boolean
Private mstr������ַ As String
Private mstr������ As String
Private mdomOutput As New MSXML2.DOMDocument
Dim intCount As Integer             '��¼�¿�,�ɿ�ʹ�ô���,�Ա��´�����ȱʡ

Private Enum MSHCol
    ����
    �Ա�
    סַ
    ����ҽԺ
End Enum

Public Function ReadCard(ByVal str������ַ As String, ByVal strUser As String, ByVal strPwd As String, ByVal bln��Ϣת�� As Boolean) As String
    mstr������ = ""
    mstr�������� = ""
    mstrUser = strUser
    mstrPwd = strPwd
    mbln��Ϣת�� = bln��Ϣת��
    mstr������ַ = str������ַ
    Me.Show 1
    ReadCard = mstr������
End Function

Private Sub chk������_Click()
    mstr�������� = IIf(chk������.value = 1, "ͨ�þ��￨", "���￨")
End Sub

Private Sub cmdCard_Click(Index As Integer)
    Dim blnNew As Boolean           '�Ƿ����µĲ��˵���
    Dim strSQL As String
    Dim str������ As String
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�����½������Ĳ����޵�����,������ID��Ϊ�������ϴ�,�Ժ���ʹ�õ��ò���,���µ�����
    
    If Index = 0 Then
        If Me.txt����.Text = "" Then
            MsgBox "������������Ϊ�գ�"
            Exit Sub
        End If
        If Val(lbl��״̬.Tag) <> 0 Then
            MsgBox "��ǰ��״̬Ϊ��" & lbl��״̬.Caption & "��������ʹ�ã�"
            Exit Sub
        End If
        
        'ȷ��,�������²�����Ϣ
        str������ = txt������.Text
        '����Ƿ���ڸò��˵���Ϣ
        '1������д˵����ţ�˵�����ڸò���
        '2��������˷�����¼�оɿ��Ŵ��ڣ�˵�����ڸò���
        strSQL = " Select * From ������Ϣ Where IC����=[1]"
        'Call OpenRecordset(rsTemp, "����Ƿ���ڸò��˵���Ϣ", strSQL)
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����Ƿ���ڸò��˵���Ϣ", str������)
        If rsTemp.RecordCount = 0 Then
            strSQL = " Select * From ������Ϣ Where ����ID=(Select ����ID From ���˷�����¼ Where �¿���=[1])"
            'Call OpenRecordset(rsTemp, "����Ƿ���ڸò��˵���Ϣ", strSQL)
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����Ƿ���ڸò��˵���Ϣ", Me.txt����.Text)
            If rsTemp.RecordCount = 0 Then
                blnNew = True
            End If
        End If
        
        If blnNew Then
            '�ޱ����ʻ�����Ϊû�в�����Ϣ
            If lng����ID = 0 Then lng����ID = gobjDatabase.GetNextNO(1)
            strSQL = "zl_������Ϣ_Insert(" & lng����ID & ",NULL,NULL,'�Է�ҽ��'," & _
                "'" & txt����.Text & "','" & cbo�Ա�.Text & "'," & DateDiff("yyyy", msk��������.Text, gobjDatabase.CurrentDate()) & "," & _
                "To_Date('" & Me.msk��������.Text & "','YYYY-MM-DD')," & _
                "NULL,'" & IIf(Me.cbo֤������.ListIndex = 0, Me.txt֤����.Text, "") & "',NULL,'" & cboְҵ.Text & "'," & _
                "NULL,NULL,NULL,NULL,'" & Me.txt��ַ.Text & "','" & Me.txt�绰.Text & "','" & Me.txt�ʱ�.Text & "'," & _
                "'" & Me.txt��������.Text & "',NULL,NULL,'" & Me.txt�����绰.Text & "',NULL,'" & txt������λ.Text & "','" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��λ��ַ.Text & "'," & _
                "NULL,NULL,NULL,NULL,To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),NULL,NULL,NULL,NULL,NULL,'" & IIf(Me.cbo֤������.ListIndex = 0, "", Me.txt֤����.Text) & "')"
        Else
            lng����ID = rsTemp!����ID
            strSQL = "zl_������Ϣ_Update(" & _
                lng����ID & "," & Nvl(rsTemp!�����, "NULL") & "," & Nvl(rsTemp!סԺ��, "NULL") & ",'" & Nvl(rsTemp!�ѱ�) & "'," & _
                "'" & Nvl(rsTemp!ҽ�Ƹ��ʽ) & "','" & txt����.Text & "','" & Me.cbo�Ա�.Text & "'," & DateDiff("yyyy", msk��������.Text, gobjDatabase.CurrentDate()) & "," & _
                "To_Date('" & msk��������.Text & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsTemp!�����ص�), "", rsTemp!�����ص�) & "','" & IIf(Me.cbo֤������.ListIndex = 0, Me.txt֤����.Text, "") & "'," & _
                "'" & IIf(IsNull(rsTemp!���), "", rsTemp!���) & "','" & cboְҵ.Text & "'," & _
                "'" & IIf(IsNull(rsTemp!����), "", rsTemp!����) & "','" & IIf(IsNull(rsTemp!����), "", rsTemp!����) & "'," & _
                "'" & IIf(IsNull(rsTemp!ѧ��), "", rsTemp!ѧ��) & "','" & IIf(IsNull(rsTemp!����״��), "", rsTemp!����״��) & "'," & _
                "'" & txt��ַ.Text & "','" & txt�绰.Text & "','" & txt�ʱ�.Text & "','" & txt��������.Text & "'," & _
                "'" & IIf(IsNull(rsTemp!��ϵ�˹�ϵ), "", rsTemp!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsTemp!��ϵ�˵�ַ), "", rsTemp!��ϵ�˵�ַ) & "'," & _
                "'" & txt�����绰.Text & "'," & IIf(IsNull(rsTemp!��ͬ��λID), "NULL", rsTemp!��ͬ��λID) & "," & _
                "'" & txt������λ.Text & "','" & txt��λ�绰.Text & "'," & _
                "'" & txt��λ�ʱ�.Text & "','" & IIf(IsNull(rsTemp!��λ������), "", rsTemp!��λ������) & "'," & _
                "'" & IIf(IsNull(rsTemp!��λ�ʺ�), "", rsTemp!��λ�ʺ�) & "','" & IIf(IsNull(rsTemp!������), "", rsTemp!������) & "'," & _
                "" & IIf(IsNull(rsTemp!������), "NULL", rsTemp!������) & "," & Nvl(rsTemp!����, "NULL") & ")"
        End If
        gcnConnect.Execute strSQL, , adCmdStoredProc
        
        '����IC����,���￨��
        If InStr(1, "1005,1006,1007", mshFact.Tag) <> 0 Then
            strSQL = "zl_������Ϣ_������Ϣ(" & lng����ID & ",'���￨��','''" & txt������.Text & "''')"        '����ҽ�����صĿ���
            gcnConnect.Execute strSQL, , adCmdStoredProc
        End If
        strSQL = "zl_������Ϣ_������Ϣ(" & lng����ID & ",'IC����','''" & IIf(str������ = "", lng����ID, str������) & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
        strSQL = "zl_������Ϣ_������Ϣ(" & lng����ID & ",'һ��ͨ����ʱ��','''" & Me.Tag & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
        strSQL = "zl_������Ϣ_������Ϣ(" & lng����ID & ",'��������','''" & mstr�������� & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
'        strSQL = "zl_������Ϣ_������Ϣ(" & lng����ID & ",'��ע','''" & txtʡ.Text & "|" & txt��.Text & "|" & txt�ֵ�.Text & "|" & txt��λ��ַ.Text & "|" & txt�ֻ���.Text & "|" & txtEMAIL.Text & "''')"
'        gcnConnect.Execute strSQL, , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'ʡ','" & txtʡ.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'��','" & txt��.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'�ֵ�','" & txt�ֵ�.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'��λ��ַ','" & txt��λ��ַ.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'�ֻ���','" & txt�ֻ���.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'EMAIL','" & txtEMAIL.Text & "')", , adCmdStoredProc
        
        'todo:ֻҪ������ϵͳ�ǲ����²���,�����һ�����˷�����¼
        If rsTemp.RecordCount = 0 Then
            '������õľɵľ��￨,txt������϶�Ϊ��,��ʱӦ�ý�����ľɿ���д�벡�˷�����¼��
            strSQL = "zl_���˷�����¼_������(" & lng����ID & ",' " & IIf(txt������.Text = "", Me.txt����.Text, txt������.Text) & "',NULL," & _
                "'" & mshFact.TextMatrix(mshFact.Row, ����ҽԺ) & "'," & Me.cbo������.ItemData(Me.cbo������.ListIndex) & ",'" & mstr���� & "','" & cmdCard(0).Tag & "')"
            gcnConnect.Execute strSQL, , adCmdStoredProc
        End If
        
        mstr������ = IIf(str������ = "", lng����ID, str������)
    End If
    
  '  MsgBox lng����ID & "|" & mstr������
    Unload Me
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    With Me.cbo����״��
        .AddItem "�ѻ�"
        .ItemData(.NewIndex) = 0
        .AddItem "δ��"
        .ItemData(.NewIndex) = 1
        .AddItem "ɥż"
        .ItemData(.NewIndex) = 2
        .AddItem "���"
        .ItemData(.NewIndex) = 3
        .AddItem "����"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo�Ա�
        .AddItem "��"
        .ItemData(.NewIndex) = 0
        .AddItem "Ů"
        .ItemData(.NewIndex) = 1
        .AddItem "����"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo������
        .AddItem "ҽ����"
        .ItemData(.NewIndex) = 0
        .AddItem "ũ����"
        .ItemData(.NewIndex) = 1
        .AddItem "���￨"
        .ItemData(.NewIndex) = 2
        .AddItem "������"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo֤������
        .AddItem "���֤"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    strSQL = "Select ����,���� From ְҵ Order by ����"
    Call OpenRecordset(rsTemp, "��ȡְҵ", strSQL)
    With rsTemp
        Do While Not .EOF
            Me.cboְҵ.AddItem !����
            .MoveNext
        Loop
        Me.cboְҵ.ListIndex = 0
    End With
    
    Call ClearCons
    
    '������������¿��ɿ�ʹ�ô����趨ȱʡ
    Dim lng�¿� As Long, lng�ɿ� As Long
    '����ע������¿�,�ɿ��ۼƴ���
    lng�ɿ� = Val(GetSetting("ZLSOFT", "����һ��ͨ", "�ɿ�", 0))
    lng�¿� = Val(GetSetting("ZLSOFT", "����һ��ͨ", "�¿�", 0))
    chk������.value = IIf(lng�¿� >= lng�ɿ�, 1, 0)
    mstr�������� = IIf(lng�¿� >= lng�ɿ�, "ͨ�þ��￨", "���￨")
    Call InitMsh
End Sub

Private Sub mshFact_EnterCell()
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    
    If mshFact.TextMatrix(mshFact.Row, ����) = "" Then
        sstab.Enabled = True
        txt����.Text = ""
        Exit Sub
    End If
    msk��������.Text = "2000-01-01"         '��Щ����û�г�������,���ȱʡֵ��
    cmdCard(0).Tag = ""
    
    '����ѡ��Ĳ�����ʾ��ϸ��Ϣ
    Select Case mshFact.Tag
    Case "1001"
        Me.lbl��״̬.Tag = 0
        Me.lbl��״̬.Caption = "����"
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If mshFact.TextMatrix(mshFact.Row, ����) = nodRow.selectSingleNode("name").Text And mshFact.TextMatrix(mshFact.Row, סַ) = nodRow.selectSingleNode("address").Text Then
                '������סַ��ͬ�������
                'Me.txt������.Text = nodRow.selectSingleNode("id").selectSingleNode("cardNumber").Text '�ɿ��Ų����µ����￨�ֶ���,�����ظ����ݳ���
                Me.txt������.Text = ""
                Me.txtEMAIL.Text = nodRow.selectSingleNode("email").Text
                Me.txt��λ��ַ.Text = nodRow.selectSingleNode("companyAddress").Text
                Me.txt��λ�绰.Text = nodRow.selectSingleNode("companyPhone").Text
                Me.txt��λ�ʱ�.Text = nodRow.selectSingleNode("companyPostcode").Text
                Me.txt������.Text = ""  '�Ͽ����˿϶�û�е�����
                Me.txt��ַ.Text = nodRow.selectSingleNode("address").Text
                Me.txt�绰.Text = nodRow.selectSingleNode("homePhone").Text
                Me.txt������λ.Text = nodRow.selectSingleNode("company").Text
                Me.txt�����绰.Text = nodRow.selectSingleNode("folkPhoneNumber").Text
                Me.txt��������.Text = nodRow.selectSingleNode("folkName").Text
                Me.txt�ֵ�.Text = nodRow.selectSingleNode("street").Text
                Me.txt��.Text = nodRow.selectSingleNode("district").Text
                Me.txtʡ.Text = nodRow.selectSingleNode("province").Text
                Me.txt�ֻ���.Text = nodRow.selectSingleNode("mobile").Text
                Me.txt����.Text = nodRow.selectSingleNode("name").Text
                Me.txt�ʱ�.Text = nodRow.selectSingleNode("homePostcode").Text
                Me.txt���պ�.Text = ""
                Me.txt֤����.Text = ""
                If Val(nodRow.selectSingleNode("cftype").Text) = 0 Then
                    Me.txt֤����.Text = nodRow.selectSingleNode("cfnumber").Text
                End If
                'Me.Tag = nodRow.selectSingleNode("createTime").Text     '���没����Ϣ�Ľ���ʱ��,���油��ʱҪ��
                Me.cbo����״��.ListIndex = Val(nodRow.selectSingleNode("wedlock").Text)
                Me.cbo�Ա�.ListIndex = Val(nodRow.selectSingleNode("sex").Text)
                Me.cbo֤������.ListIndex = IIf(Val(nodRow.selectSingleNode("cftype").Text) = 0, 0, 1)
                Me.cboְҵ.Text = nodRow.selectSingleNode("metier").Text
                
                If nodRow.selectSingleNode("birthday").Text <> "" Then
                    Me.msk��������.Text = Mid(nodRow.selectSingleNode("birthday").Text, 1, 10)
                End If
            End If
        Next
    Case Else
        'ֻ������һ����¼
        Me.txt���պ�.Text = ""
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If nodRow.nodeName = "CardInfoRet" Then
                Me.txt������.Text = nodRow.selectSingleNode("cardNumber").Text
                '���¿�״̬(0��������1����ͣ��2��ע����3����ʧ��4����ע����5����ʧע��)
                Me.lbl��״̬.Tag = Val(nodRow.selectSingleNode("cardStatus").Text)
                Select Case Val(Me.lbl��״̬.Tag)
                Case 0
                    Me.lbl��״̬.Caption = "����"
                Case 1
                    Me.lbl��״̬.Caption = "��ͣ"
                Case 2
                    Me.lbl��״̬.Caption = "ע��"
                Case 3
                    Me.lbl��״̬.Caption = "��ʧ"
                Case 4
                    Me.lbl��״̬.Caption = "��ע��"
                Case 5
                    Me.lbl��״̬.Caption = "��ʧע��"
                End Select
                cmdCard(0).Tag = nodRow.selectSingleNode("reportTime").Text '�ɿ��Ľ���ʱ��
            End If
            If nodRow.nodeName = "TPersonBasalInfo" Then
                Me.txt������.Text = nodRow.selectSingleNode("personid").Text
                Me.txt����.Text = nodRow.selectSingleNode("name").Text
                Me.txtEMAIL.Text = nodRow.selectSingleNode("email").Text
                Me.cbo�Ա�.ListIndex = Val(nodRow.selectSingleNode("sex").Text)
                If nodRow.selectSingleNode("birthday").Text <> "" Then
                    Me.msk��������.Text = Mid(nodRow.selectSingleNode("birthday").Text, 1, 10)
                End If
                Me.cbo����״��.ListIndex = Val(nodRow.selectSingleNode("wedlock").Text)
                Me.cbo֤������.ListIndex = IIf(Val(nodRow.selectSingleNode("cftype").Text) = 0, 0, 1)
                Me.txt֤����.Text = ""
                If Val(nodRow.selectSingleNode("cftype").Text) = 0 Then
                    Me.txt֤����.Text = nodRow.selectSingleNode("cfnumber").Text
                End If
                Me.Tag = nodRow.selectSingleNode("createTime").Text     '���没����Ϣ�Ľ���ʱ��,���油��ʱҪ��
            End If
            If nodRow.nodeName = "TPersonExtendInfo" Then
                Me.txt��λ��ַ.Text = nodRow.selectSingleNode("companyAddress").Text
                Me.txt��λ�绰.Text = nodRow.selectSingleNode("companyPhone").Text
                Me.txt��λ�ʱ�.Text = nodRow.selectSingleNode("companyPostcode").Text
                Me.txt��ַ.Text = nodRow.selectSingleNode("address").Text
                Me.txt�绰.Text = nodRow.selectSingleNode("homePhone").Text
                Me.txt������λ.Text = nodRow.selectSingleNode("company").Text
                Me.txt�����绰.Text = nodRow.selectSingleNode("folkPhoneNumber").Text
                Me.txt��������.Text = nodRow.selectSingleNode("folkName").Text
                Me.txt�ֵ�.Text = nodRow.selectSingleNode("street").Text
                Me.txt��.Text = nodRow.selectSingleNode("district").Text
                Me.txtʡ.Text = nodRow.selectSingleNode("province").Text
                Me.txt�ֻ���.Text = nodRow.selectSingleNode("mobile").Text
                Me.txt�ʱ�.Text = nodRow.selectSingleNode("homePostcode").Text
                Me.cboְҵ.Text = nodRow.selectSingleNode("metier").Text
            End If
        Next
    
    End Select
    
    If mshFact.Tag = "1004" Then Me.txt���պ�.Text = Me.txt����.Text
    If Me.Tag <> "" Then Me.Tag = Format(Mid(Me.Tag, 1, 10), "YYYYMMdd") & Format(Mid(Me.Tag, 12, 8), "HHmmss")
    If cmdCard(0).Tag <> "" Then cmdCard(0).Tag = Format(Mid(cmdCard(0).Tag, 1, 10), "YYYYMMdd") & Format(Mid(cmdCard(0).Tag, 12, 8), "HHmmss")
    
    Exit Sub
errHand:
    MsgBox "��ʾָ��������Ϣʱ��������:" & Err.Description
    Resume
End Sub

Private Sub opt_Click(Index As Integer)
    chk������.Enabled = (Index = 0)
    
    Select Case Index
    Case 0
        If chk������.value = 1 Then
            mstr�������� = "ͨ�þ��￨"
        Else
            mstr�������� = "���￨"
        End If
    Case 1
        mstr�������� = "����"
    Case 2
        mstr�������� = "���֤"
    Case 3
        mstr�������� = "ҽ����"
    Case 4
        mstr�������� = "ũ����"
    Case 5
        mstr�������� = "���ҽ����"
    Case 6
        mstr�������� = "���ũ����"
    End Select
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstr���� = ""
    Call ClearCons
    Call InitMsh
    sstab.Enabled = False
    Call ReadPatient
End Sub

Private Sub ClearCons()
    Dim objControl As Object
    
    Me.lbl��״̬.Tag = 0
    Me.lbl��״̬.Caption = ""
    For Each objControl In Me.Controls
        If UCase(objControl.Container.Name) = "SSTAB" Then
            Select Case Mid(UCase(objControl.Name), 1, 3)
            Case "TXT"
                objControl.Text = ""
            Case "CBO"
                objControl.ListIndex = 0
            End Select
        End If
    Next
    
    If Me.opt(4).value Or Me.opt(6).value Then
        Me.cbo������.ListIndex = 1
    ElseIf Me.opt(3).value Or Me.opt(5).value Then
        Me.cbo������.ListIndex = 0
    ElseIf Me.opt(0).value Then
        Me.cbo������.ListIndex = 2
    Else
        Me.cbo������.ListIndex = 3
    End If
    
End Sub

Private Function ReadPatient() As Boolean
    Dim strType As String
    Dim strPatient As String
    On Error GoTo errHand
    
    If opt(0).value Then
        If chk������.value = 1 Then
            strType = "1005"
        Else
            strType = "1001"
        End If
    ElseIf opt(1).value Then
        strType = "1003"
    ElseIf opt(2).value Then
        strType = "1002"
    ElseIf opt(3).value Or opt(4).value Then
        strType = 1004
    ElseIf opt(5).value Then
        strType = 1007
    ElseIf opt(6).value Then
        strType = 1006
    End If
    
    '����WebServices��ȡ�������
    '1�� ��ѯ����
    'a)  .����˵��
    'i.���Ͳ�������
    'SearchType:  ��ѯ��������
    'ii.��ѯ�ؼ��ֲ���
    'Cardnumber:  ����
    'Sfzj:  ���֤��
    'Jzkmm:  ���￨����
    'Bxh:  ���պ�
    '
    'b)  .��������
    'i.  getPersonInfo(String SearchType,String [��ѯ�ؼ��ֲ���])
    'ii.���Ͳ�������˵��
    '1001: ͨ�����Ų�ѯ�ɿ���Ϣ
    '1002: ͨ�����֤�����ѯ����Ϣ
    '1003: ͨ�����￨�����ѯ����Ϣ
    '1004: ͨ�����պŲ�ѯ����Ϣ
    '1005: ͨ�����Ų�ѯ����Ϣ
    If Not ���ýӿ�("getPersonInfo", strType, txt����.Text) Then
        sstab.Enabled = True
        txt����.Text = ""
        Exit Function
    End If
    '��������Ϣ��������
    Me.mshFact.Tag = strType
    Call AnalysePatient
    
    '����ע������¿�,�ɿ��ۼƴ���
    Dim lng�¿� As Long, lng�ɿ� As Long, str���� As String
    lng�ɿ� = Val(GetSetting("ZLSOFT", "����һ��ͨ", "�ɿ�", 0))
    lng�¿� = Val(GetSetting("ZLSOFT", "����һ��ͨ", "�¿�", 0))
    str���� = Format(gobjDatabase.CurrentDate(), "yyyyMMdd")
    If str���� <> GetSetting("ZLSOFT", "����һ��ͨ", "����", "") Then lng�¿� = 0: lng�ɿ� = 0
    If opt(0).value Then
        If chk������.value = 1 Then
            lng�¿� = lng�¿� + 1
        Else
            lng�ɿ� = lng�ɿ� + 1
        End If
        Call SaveSetting("ZLSOFT", "����һ��ͨ", "�ɿ�", lng�ɿ�)
        Call SaveSetting("ZLSOFT", "����һ��ͨ", "�¿�", lng�¿�)
        Call SaveSetting("ZLSOFT", "����һ��ͨ", "����", str����)
    End If
    
    If mstr�������� = "����" Then mstr���� = Me.txt����.Text
    If opt(2).value Then Me.txt����.Text = ""
    mshFact.Row = 1: mshFact.Col = 0
    Call mshFact_EnterCell
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function AnalysePatient() As Boolean
    Dim intRow As Integer
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    '����������Ϣ���ش�
    
    intRow = 1
    Select Case mshFact.Tag
    Case "1001"
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            mshFact.TextMatrix(intRow, ����) = nodRow.selectSingleNode("name").Text
            mshFact.TextMatrix(intRow, �Ա�) = IIf(nodRow.selectSingleNode("sex").Text = "0", "��", "Ů")
            mshFact.TextMatrix(intRow, סַ) = nodRow.selectSingleNode("address").Text
            mshFact.TextMatrix(intRow, ����ҽԺ) = nodRow.selectSingleNode("id").selectSingleNode("hospitalid").Text
            intRow = intRow + 1
            mshFact.Rows = mshFact.Rows + 1
        Next
    Case Else       '1005
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If nodRow.nodeName = "CardInfoRet" Then
                mshFact.TextMatrix(intRow, ����ҽԺ) = nodRow.selectSingleNode("hospitalNumber").Text
            End If
            If nodRow.nodeName = "TPersonBasalInfo" Then
                mshFact.TextMatrix(intRow, ����) = nodRow.selectSingleNode("name").Text
                mshFact.TextMatrix(intRow, �Ա�) = IIf(nodRow.selectSingleNode("sex").Text = "0", "��", "Ů")
            End If
            If nodRow.nodeName = "TPersonExtendInfo" Then
                mshFact.TextMatrix(intRow, סַ) = nodRow.selectSingleNode("address").Text
            End If
        Next
    End Select
    
    AnalysePatient = True
    Exit Function
errHand:
    MsgBox "װ�ز�����Ϣʱ��������:" & Err.Description
End Function

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Dim varReturn As Variant
    varReturn = IIf(IsNull(varValue), DefaultValue, varValue)
    Nvl = Replace(varReturn, "'", "")
End Function

Private Function GetElemnetValue(ByVal Name As String) As String
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        GetElemnetValue = xmlElement.Text
'    Else
'        'ȡ��
'        Debug.Assert False
    End If
End Function

Private Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'���ܣ��õ�ָ�����Ե�ֵ
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Private Function ���ýӿ�(ByVal strFunction As String, ByVal strType As String, ByVal strKey As String) As Boolean
'    ----------------------------------------------------------------
    '��������   �����ýӿں���
    '��д��     ������
'    ��д����   ��2009-07-31
'    ----------------------------------------------------------------
    Dim str���� As String, lng���к� As Long, str������Ϣ As String
    Dim strURL As String, strSoapRequest As String
    Dim objHttp As MSXML2.XMLHTTP
    On Error GoTo errHand
    
    Set objHttp = New MSXML2.XMLHTTP
    strURL = mstr������ַ & "?op=" & strFunction
    
    strSoapRequest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>" & _
                "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">" & _
                "<soapenv:Header>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:user>" & mstrUser & "</ns:user>" & _
                        "<ns:pwd>" & mstrPwd & "</ns:pwd>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Header>" & _
                "<soapenv:Body>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:SearchType>" & strType & "</ns:SearchType>"
    Select Case strType
    Case "1001"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1002"
        strSoapRequest = strSoapRequest & "<ns:Sfzh>" & strKey & "</ns:Sfzh>"
    Case "1003"
        strSoapRequest = strSoapRequest & "<ns:Jzkmm>" & strKey & "</ns:Jzkmm>"
    Case "1004"
        strSoapRequest = strSoapRequest & "<ns:Bxh>" & strKey & "</ns:Bxh>"
    Case "1005"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1006"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1007"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    End Select
                        
    strSoapRequest = strSoapRequest & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Body>" & _
                "</soapenv:Envelope>"
    If mbln��Ϣת�� = False Then
        objHttp.Open "post", strURL, False
        objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
        objHttp.setRequestHeader "SOAPAction", strURL
        
        '���ݷ��ص�״̬��Ϣ���ж��Ƿ�ɹ�
        objHttp.send (strSoapRequest)
        If objHttp.status <> 200 Then
            MsgBox "������Ϣ��[" & objHttp.status & "]" & objHttp.responseText
            Exit Function
        End If
    Else
        'д������
        If Not SendRequest(str����, lng���к�, strFunction, strURL, strSoapRequest) Then Exit Function
        
        '��ʾ�ȴ�����
        If frmWait.SendRequest(str����, lng���к�, str������Ϣ) = False Then
            If str������Ϣ <> "" Then MsgBox "������Ϣ��" & str������Ϣ
            Exit Function
        End If
    End If
    
    '�ϵ����ô�
    Set mdomOutput = New MSXML2.DOMDocument
    If mbln��Ϣת�� = False Then
        If mdomOutput.loadXML(objHttp.responseText) = False Then
            MsgBox "���׺�����" & strFunction & "���������ݸ�ʽ����ȷ��"
            Exit Function
        End If
    Else
        If mdomOutput.loadXML(str������Ϣ) = False Then
            MsgBox "���׺�����" & strFunction & "���������ݸ�ʽ����ȷ��"
            Exit Function
        End If
    End If
    
    ���ýӿ� = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function SendRequest(str���� As String, lng���к� As Long, _
    ByVal strFuncName As String, ByVal strURL As String, ByVal strSoapRequest As String) As Boolean
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '������������д�����ݱ�
    
    str���� = Format(gobjDatabase.CurrentDate, "yyyyMMdd")
    lng���к� = gobjDatabase.GetNextId("��Ϣת��")
    
    gcnConnect.BeginTrans
    blnTrans = True
    
    '��������
    gcnConnect.Execute "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "')", , adCmdStoredProc
    
    '�������������
    intCount = Len(strSoapRequest) \ 1000
    If Len(strSoapRequest) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strSoapRequest, intRow * 1000 + 1, 1000)
        gcnConnect.Execute "zl_��Ϣת��_Insert('" & str���� & "'," & lng���к� & "," & intRow + 1 & ",'" & strRow & "')", , adCmdStoredProc
    Next
    
    gcnConnect.CommitTrans
    blnTrans = False
    SendRequest = True
    Exit Function
errHand:
    If blnTrans Then gcnConnect.RollbackTrans
    MsgBox Err.Description
End Function

Private Sub InitMsh()
    With mshFact
        .Clear
        .Rows = 2: .Cols = 4
        .TextMatrix(0, ����) = "����"
        .TextMatrix(0, �Ա�) = "�Ա�"
        .TextMatrix(0, סַ) = "סַ"
        .TextMatrix(0, ����ҽԺ) = "����ҽԺ"
        .ColWidth(����) = 1200
        .ColWidth(�Ա�) = 500
        .ColWidth(סַ) = 3000
        .ColWidth(����ҽԺ) = 1000
    End With
End Sub
