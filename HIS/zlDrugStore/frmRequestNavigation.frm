VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequestNavigation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ��������Զ�������"
   ClientHeight    =   7605
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmRequestNavigation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8145
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7170
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":1582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicSetup 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   -15
      Width           =   1485
      Begin VB.Image imgSetup 
         Height          =   6645
         Left            =   60
         Picture         =   "frmRequestNavigation.frx":328C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5265
      TabIndex        =   1
      Top             =   7095
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   330
      TabIndex        =   2
      Top             =   7095
      Width           =   1230
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6630
      TabIndex        =   0
      Top             =   7095
      Width           =   1230
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmRequestNavigation.frx":8872
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":974C
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":9B9E
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":9FF0
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   -120
      Width           =   6555
      Begin VB.CheckBox chk����� 
         Caption         =   "����ⷿ�޿��ÿ��ʱ�����������¼"
         Height          =   180
         Left            =   360
         TabIndex        =   59
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "���췽ʽ"
         Height          =   4600
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   6135
         Begin VB.CheckBox chkLowerLimit 
            Caption         =   "�̶�������������"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   2880
            TabIndex        =   51
            Top             =   2085
            Width           =   1875
         End
         Begin VB.OptionButton optMode 
            Caption         =   "4����ҩƷ�Ĵ������ޣ������ۺϿ���"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   45
            Top             =   2655
            Width           =   3285
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�̶�������������"
            Enabled         =   0   'False
            Height          =   225
            Left            =   690
            TabIndex        =   43
            Top             =   1305
            Width           =   1875
         End
         Begin VB.CheckBox chkLowerLimit 
            Caption         =   "�̶�������������"
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   690
            TabIndex        =   42
            Top             =   2085
            Width           =   1875
         End
         Begin VB.OptionButton optMode 
            Caption         =   "6��������������"
            Height          =   195
            Index           =   5
            Left            =   330
            TabIndex        =   15
            Top             =   3960
            Width           =   3255
         End
         Begin VB.CheckBox chk��������С�������� 
            Caption         =   "��������С����������ҩƷ"
            Height          =   225
            Left            =   690
            TabIndex        =   36
            Top             =   480
            Width           =   2715
         End
         Begin VB.OptionButton optMode 
            Caption         =   "5������ָ��ʱ�䷶Χ�ڵ����쵥"
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   14
            Top             =   3390
            Width           =   3285
         End
         Begin VB.OptionButton optMode 
            Caption         =   "3����ҩƷ�Ĵ�������"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   13
            Top             =   1875
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "2����ҩƷ�Ĵ�������"
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   12
            Top             =   1080
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "1������ָ��ʱ�䷶Χ��ҩƷ��������"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   3405
         End
         Begin VB.OptionButton optMode 
            Caption         =   "7������ָ��ʱ�䷶Χ��ҩƷ��������"
            Height          =   180
            Index           =   6
            Left            =   330
            TabIndex        =   52
            Top             =   240
            Width           =   3405
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ȡ���������޵�ҩƷ����ʹ��ǰ�ⷿ�Ĵ�����ʼ�ձ��������ޱ�׼�������������쵥"
            ForeColor       =   &H00004000&
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   46
            Top             =   2880
            Width           =   5460
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ������������Լ��趨�Ŀ�����ޡ�����������������������������쵥"
            ForeColor       =   &H00004000&
            Height          =   510
            Index           =   5
            Left            =   390
            TabIndex        =   20
            Top             =   4200
            Width           =   5400
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ����쵥��δ�������������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   4
            Left            =   390
            TabIndex        =   19
            Top             =   3630
            Width           =   4680
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ��ǰ�ⷿ��ҩƷ������ʼ�ձ��������ޱ�׼�������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   18
            Top             =   2340
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ��ǰ�ⷿ��ҩƷ������ʼ�ձ��������ޱ�׼�������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   1560
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ָ����ʱ�䷶Χ����ҩƷ��������Ϊ���ݣ��������ε����쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   5400
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ָ����ʱ�䷶Χ����ҩƷ�����ۣ���ҩ����Ϊ���ݣ��������ε����쵥��ע�⣺�������쵥����ʼʱ��Ϊ�ϸ����쵥�Ľ�ֹʱ��."
            ForeColor       =   &H00004000&
            Height          =   780
            Index           =   6
            Left            =   360
            TabIndex        =   53
            Top             =   480
            Width           =   5295
         End
      End
      Begin VB.CheckBox chk�������� 
         Caption         =   "����������Ϊ�ο�����"
         Height          =   180
         Left            =   2880
         TabIndex        =   54
         Top             =   1750
         Width           =   2415
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "�ǳ���ҩƷ"
         Height          =   180
         Index           =   2
         Left            =   3600
         TabIndex        =   49
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "����ҩƷ"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   48
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDrugType 
         Caption         =   "ȫ��ҩƷ"
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   47
         Top             =   1440
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "ȷ����������Ϊ����"
         Height          =   180
         Left            =   360
         TabIndex        =   44
         Top             =   1750
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   6255
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   2235
      End
      Begin VB.Label lblDrugType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��ҩƷ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   50
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��һ���������������쵥�ķ�ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��׼�����ĸ��ⷿ������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   780
         Width           =   2730
      End
      Begin VB.Label lbl�ⷿ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   2
      Left            =   1470
      TabIndex        =   37
      Top             =   -120
      Width           =   6555
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Width           =   6255
      End
      Begin MSComctlLib.TreeView tvw��; 
         Height          =   5370
         Left            =   150
         TabIndex        =   39
         Top             =   1035
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   9472
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
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ָ��ҩƷ��������С��Χ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   41
         Top             =   240
         Width           =   6300
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����ѡ��ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   40
         Top             =   810
         Width           =   1560
      End
   End
   Begin VB.Frame fraStep 
      Height          =   7095
      Index           =   1
      Left            =   1470
      TabIndex        =   21
      Top             =   -120
      Width           =   6555
      Begin VB.CheckBox chk�����в�ҩ 
         Caption         =   "�в�ҩ"
         Height          =   180
         Left            =   3600
         TabIndex        =   57
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk�����г�ҩ 
         Caption         =   "�г�ҩ"
         Height          =   180
         Left            =   2460
         TabIndex        =   56
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk��������ҩ 
         Caption         =   "����ҩ"
         Height          =   180
         Left            =   1320
         TabIndex        =   55
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   5370
         TabIndex        =   33
         Top             =   3090
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   5370
         TabIndex        =   31
         Top             =   2700
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   285
         Left            =   4170
         TabIndex        =   26
         Top             =   1350
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   158400515
         CurrentDate     =   38096
      End
      Begin MSComctlLib.ListView lvwSelect 
         Height          =   4845
         Left            =   150
         TabIndex        =   24
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8546
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   285
         Left            =   4170
         TabIndex        =   28
         Top             =   1980
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   158400515
         CurrentDate     =   38096
      End
      Begin VB.Label lbl���ʷ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ʷ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   58
         Top             =   6120
         Width           =   780
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&T)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   32
         Top             =   3150
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&X)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl����޶����� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����޶�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   29
         Top             =   2460
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbl������������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   35
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lbl����ѡ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lbl����ʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&E)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   27
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lbl��ʼʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   25
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ڶ�����ָ����������С��Χ�����в�ҩ��Ч��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   23
         Top             =   240
         Width           =   6300
      End
   End
End
Attribute VB_Name = "frmRequestNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ģʽ
    ���������� = 0
    �������� = 1
    �������� = 2
    ���������� = 3
    �������쵥δ���� = 4
    ������������
    ������������
End Enum
Private mstr���� As String
Private mblnOK As Boolean
Private mlngStockID As Long                 '����ⷿID
Private mintAutoType As Integer             '�Զ�����ʱ�����췽ʽ��1-����������;2-��������;3-��������;4-����������;5-�������쵥δ����;6-������������;7-������������
Private mintCheck As Integer                '��������
Private mblnFirst  As Boolean
Private mblnStart As Boolean
Private mfrmMain As Object
Private mintStep As Integer
Private mIntCol�������� As Integer
Private mIntCol��д���� As Integer
Private mstr�ⷿid As String
Private mstr���ʷ��� As String          '������¼ѡ��Ĳ��ʷ���
Private mstr���� As String      '5-��ҩ��6-��ҩ 7-��ҩ

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��
Private mintUnit As Integer                 '��λ
Private mint�����γ��� As Integer           '0-�������γ���,1-�����γ���
Private Sub CheckAll(ByVal myNodes As Nodes, ByVal blnCheck As Boolean)
    Dim tmpNode As Node
    
    For Each tmpNode In myNodes
        tmpNode.Checked = True
        If tmpNode.Child > 0 Then
            Call CheckAll(tmpNode, blnCheck)
        End If
    Next
End Sub

Private Function IniҩƷ����() As Boolean
    'ҩƷ��;����
    Dim lng����id As Long
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node
    
    mstr���ʷ��� = ""
    If chk��������ҩ.Value = 1 Then
        mstr���ʷ��� = "1"
        mstr���� = "5"
    End If
    
    If mstr���ʷ��� <> "" Then
        If chk�����г�ҩ.Value = 1 And chk�����г�ҩ.Visible = True Then
            mstr���ʷ��� = mstr���ʷ��� & ",2"
            mstr���� = mstr���� & ",6"
        End If
    Else
        If chk�����г�ҩ.Value = 1 And chk�����г�ҩ.Visible = True Then
            mstr���ʷ��� = "2"
            mstr���� = "6"
        End If
    End If
    
    If mstr���ʷ��� <> "" Then
        If chk�����в�ҩ.Value = 1 And chk�����в�ҩ.Visible = True Then
            mstr���ʷ��� = mstr���ʷ��� & ",3"
            mstr���� = mstr���� & ",7"
        End If
    Else
        If chk�����в�ҩ.Value = 1 And chk�����в�ҩ.Visible = True Then
            mstr���ʷ��� = "3"
            mstr���� = "7"
        End If
    End If
    
    If mstr���ʷ��� = "" Then
        MsgBox "��ѡ����ʷ��࣡", vbInformation, gstrSysName
        IniҩƷ���� = False
        Exit Function
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select Level As ��, a.ID, a.�ϼ�id, a.����, Decode(a.����, 1, '����ҩ', 2, '�г�ҩ', '�в�ҩ') As ����" & _
                " From ���Ʒ���Ŀ¼ a" & _
                " Where a.���� in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList)))" & _
                " Start With a.�ϼ�id Is Null" & _
                " Connect By Prior a.ID = a.�ϼ�id" & _
                " Order By Level"

    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ��;����", mstr���ʷ���)

    tvw��;.Nodes.Clear
    Set objNode = tvw��;.Nodes.Add(, , "Root", "���з���", "Item")
    
    If InStr(mstr���ʷ���, "1") > 0 Then
        Set objNode = tvw��;.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", "Item")
    End If
    If InStr(mstr���ʷ���, "2") > 0 Then
        Set objNode = tvw��;.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", "Item")
    End If
    If InStr(mstr���ʷ���, "3") > 0 Then
        Set objNode = tvw��;.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", "Item")
    End If
    
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
    
    IniҩƷ���� = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RunByStep(ByVal intStep As Integer)
    Dim bln���Ϸ��� As Boolean
    
    Select Case intStep
        Case 0  '��һ��
            fraStep(0).Visible = True
            fraStep(0).ZOrder
            
            fraStep(1).Visible = False
            fraStep(2).Visible = False
            
            cmdPrevious.Enabled = False
            cmdNext.Caption = "��һ��(&N)"
        Case 1  '�ڶ���
            fraStep(1).Visible = True
            fraStep(1).ZOrder
            
            fraStep(0).Visible = False
            fraStep(2).Visible = False
            
            cmdPrevious.Enabled = True
            cmdNext.Caption = "��һ��(&N)"
            
            '�������͵�λ��
            Call ResizeDrug
            '�������Ϸ���λ��
            If chk��������ҩ.Visible = False And chk�����г�ҩ.Visible = True Then
                chk�����г�ҩ.Left = chk��������ҩ.Left
            End If
            If chk��������ҩ.Visible = True And chk�����г�ҩ.Visible = False And chk�����в�ҩ.Visible = True Then
                chk�����в�ҩ.Left = chk�����г�ҩ.Left
            End If
            If chk��������ҩ.Visible = False And chk�����г�ҩ.Visible = False And chk�����в�ҩ.Visible = True Then
                chk�����в�ҩ.Left = chk��������ҩ.Left
            End If
            
        Case 2  '������
            bln���Ϸ��� = IniҩƷ����
            If bln���Ϸ��� = True Then
                fraStep(2).Visible = True
                fraStep(2).ZOrder
                
                fraStep(0).Visible = False
                fraStep(1).Visible = False
                
                cmdPrevious.Enabled = True
                cmdNext.Caption = "���(&F)"
            Else
                Exit Sub
            End If
        Case 3  '���
            If optMode(������������) Then
                '���������������С�ڿ����������
                '���������������������������Ϊ��
                If dtp��ʼʱ��.Value > dtp����ʱ��.Value Then
                     MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��", vbInformation, gstrSysName
                     Call RunByStep(1)
                     dtp��ʼʱ��.SetFocus
                     Exit Sub
                End If
                
                If Trim(txt��������.Text) = "" Then
                    MsgBox "������������������", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Trim(txt��������.Text) = "" Then
                    MsgBox "������������������", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(txt��������.Text) Then
                    MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(txt��������.Text) Then
                    MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Val(txt��������.Text) <= 0 Then
                    MsgBox "���������������С���㣡", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Val(txt��������.Text) <= 0 Then
                    MsgBox "���������������С���㣡", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Val(txt��������.Text) < Val(txt��������.Text) Then
                    MsgBox "���������������С�ڿ������������", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
                If Val(txt��������.Text) > 300 Then
                    MsgBox "��������������ܴ���300�죡", vbInformation, gstrSysName
                    mintStep = 1
                    Call RunByStep(1)
                    txt��������.SetFocus
                    Exit Sub
                End If
            ElseIf optMode(������������) Then
                If dtp��ʼʱ��.Value > dtp����ʱ��.Value Then
                    MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��", vbInformation, gstrSysName
                    Call RunByStep(1)
                    If dtp��ʼʱ��.Enabled = True Then
                        dtp��ʼʱ��.SetFocus
                    Else
                        dtp����ʱ��.SetFocus
                    End If
                    Exit Sub
                End If
            End If
            
            '��������
            Call Get���ʹ�
            If Not CheckData Then Exit Sub
            
            mblnOK = True
            Unload Me
    End Select
End Sub

'Private Sub cbo���ʷ���_Click()
'    If cbo���ʷ���.ListIndex >= 0 Then
'        Call IniҩƷ����(cbo���ʷ���.ItemData(cbo���ʷ���.ListIndex))
'    End If
'End Sub


Private Sub chkLowerLimit_Click(index As Integer)
    If chkLowerLimit(index).Value = 1 Then
        chkLowerLimit(Abs(index - 1)).Value = 0
    End If
    
    If chkLowerLimit(0).Value = 1 Then
        lblTip(2).Caption = "ʹ��ǰ�ⷿ��ҩƷ���������ٱ��������ޱ�׼�������������쵥"
    ElseIf chkLowerLimit(1).Value = 1 Then
        lblTip(2).Caption = "ʹ��ǰ�ⷿ��ҩƷ���������ٱ��������ޱ�׼�������������쵥"
    Else
        lblTip(2).Caption = "ʹ��ǰ�ⷿ��ҩƷ������ʼ�ձ��������ޱ�׼�������������쵥"
    End If
End Sub

Private Sub chk����_Click()
    If chk����.Value = 1 Then
        lblTip(1).Caption = "ʹ��ǰ�ⷿ��ҩƷ���������ٱ��������ޱ�׼�������������쵥"
    Else
        lblTip(1).Caption = "ʹ��ǰ�ⷿ��ҩƷ������ʼ�ձ��������ޱ�׼�������������쵥"
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    mintStep = IIf(mintStep = 3, 3, mintStep + 1)
    Call RunByStep(mintStep)
End Sub

Private Sub cmdPrevious_Click()
    mintStep = IIf(mintStep = 0, 0, mintStep - 1)
    Call RunByStep(mintStep)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    '----ȱʡѡ�����м���----
    
    If Not mblnFirst Then Exit Sub
    
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
    
    If mintAutoType <> 7 Then
        optMode(0).Visible = True
        chk��������С��������.Visible = True
        lblTip(0).Visible = True
        
        optMode(1).Visible = True
        chk����.Visible = True
        lblTip(1).Visible = True
        
        optMode(2).Visible = True
        chkLowerLimit(0).Visible = True
        chkLowerLimit(1).Visible = True
        lblTip(2).Visible = True
        
        optMode(3).Visible = True
        lblTip(3).Visible = True
        
        optMode(4).Visible = True
        lblTip(4).Visible = True
        
        optMode(5).Visible = True
        lblTip(5).Visible = True
        
        optMode(6).Visible = False
        lblTip(6).Visible = False
    Else
        optMode(0).Visible = False
        chk��������С��������.Visible = False
        lblTip(0).Visible = False
        
        optMode(1).Visible = False
        chk����.Visible = False
        lblTip(1).Visible = False
        
        optMode(2).Visible = False
        chkLowerLimit(0).Visible = False
        chkLowerLimit(1).Visible = False
        lblTip(2).Visible = False
        
        optMode(3).Visible = False
        lblTip(3).Visible = False
        
        optMode(4).Visible = False
        lblTip(4).Visible = False
        
        optMode(5).Visible = False
        lblTip(5).Visible = False
        
        optMode(6).Visible = True
        lblTip(6).Visible = True
        optMode(6).Value = True
    End If
    
    Call RunByStep(0)
    lvwSelect.ListItems(1).Checked = True
    Call lvwSelect_ItemCheck(lvwSelect.ListItems(1))
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim dateCurDate As Date
    Dim intStock As Integer
    Dim strվ��Ȩ�� As String
    Dim int�������� As Integer
    Dim int���� As Integer
    Dim int���������� As Integer
    
    mblnStart = False
    mblnFirst = True
    mintStep = 0
    
    int�������� = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "����������Ϊ�ο�����", 0)))
    int���� = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ȷ����������Ϊ����", 0)))
    int���������� = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "����������", 0)))
    mint�����γ��� = Val(zlDataBase.GetPara("ҩƷ�����γ���", glngSys, 1343, 0))
    
    Me.chk��������.Value = IIf(int�������� = 1, 1, 0)
    Me.chk����.Value = IIf(int���� = 1, 1, 0)
    Me.chk�����.Value = IIf(int���������� = 1, 1, 0)
    
    On Error GoTo errHandle

    '----��ȡҩƷ����----
    'û��ҩƷ����ʱ�Կ��Լ�����ֻ�ǲ����޶�ҩƷ�ļ���
    gstrSQL = "Select ����,���� From ҩƷ����"
    Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "��ȡ��������")
    
    Me.lvwSelect.ListItems.Clear
    Me.lvwSelect.ListItems.Add , "R", "���м���", , 1
    With rsTemp
        Do While Not .EOF
            Me.lvwSelect.ListItems.Add , "K" & !����, !����, , 1
            .MoveNext
        Loop
    End With
    
    '����ȡ�ò���ӵ�еĲ���
'    cbo���ʷ���.Clear
    gstrSQL = " Select distinct substr(��������,1,2) ��������" & _
              " From ��������˵��" & _
              " Where ����ID = [1] ANd �������� IN ('��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��')"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����ȡ�ò���ӵ�еĲ���]", mlngStockID)
    
    With rsTemp
        Do While Not .EOF
'            cbo���ʷ���.AddItem IIf(!�������� Like "��ҩ*", "����ҩ", IIf(!�������� Like "��ҩ*", "�г�ҩ", "�в�ҩ"))
'            cbo���ʷ���.ItemData(cbo���ʷ���.NewIndex) = IIf(!�������� Like "��ҩ*", 5, IIf(!�������� Like "��ҩ*", 6, 7))
            
            If !�������� Like "��ҩ*" Then
                chk��������ҩ.Visible = True
                chk��������ҩ.Value = 1
            ElseIf !�������� Like "��ҩ*" Then
                chk�����г�ҩ.Visible = True
                chk�����г�ҩ.Value = 1
            Else
                chk�����в�ҩ.Visible = True
                chk�����в�ҩ.Value = 1
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then
'            cbo���ʷ���.ListIndex = 0
'        Else
            Exit Sub
        End If
    End With
    
    '----��ȡҩƷ�ⷿ----
    'gstrSQL = ReturnSQL(mlngStockID, False)
    'Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��������Ŀⷿ]", mlngStockID)
    Set rsTemp = ReturnSQL(mlngStockID, Me.Caption & "[��ȡ��������Ŀⷿ]", False, 1343)
    
    If rsTemp.EOF Then
        MsgBox "û���κοⷿ�������죬����[������������]��ҩƷ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    With cbo�ⷿ
        .Clear
        mstr�ⷿid = ""
        Do While Not rsTemp.EOF
            If InStr(1, mstr�ⷿid, "|" & rsTemp!Id & "|") = 0 Then
                .AddItem rsTemp!����
                .ItemData(.NewIndex) = rsTemp!Id
                mstr�ⷿid = mstr�ⷿid & "|" & rsTemp!Id & "|"
                If rsTemp!ҩ������ = 1 And intStock = 0 Then
                    intStock = .NewIndex
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        .ListIndex = intStock
    End With
    
    dateCurDate = zlDataBase.Currentdate()
    '----����ȱʡ��ʱ�䷶Χ��һ���£�----
    Me.dtp��ʼʱ��.Value = Format(DateAdd("m", -1, dateCurDate), "yyyy-MM-dd") & " 00:00:00"
    Me.dtp����ʱ��.Value = Format(dateCurDate, "yyyy-MM-dd HH:mm:ss")
        
    If mintAutoType = 7 Then
        '��������췽ʽ��ȡ�ϴ���˵����쵥��������Ϊ��ʼʱ��
        gstrSQL = " Select a.Ƶ�� As ����ʱ�� " & _
            " From ҩƷ�շ���¼ A, " & _
            " (Select Nvl(Max(�������), Sysdate) As ������� " & _
            " From ҩƷ�շ���¼ " & _
            " Where ���� = 6 And ���� = 7 And ���ϵ�� = 1 And �ⷿid + 0 = [1] And ������� Between Sysdate - 60 And Sysdate) B " & _
            " Where a.���� = 6 And a.���� = 7 And a.���ϵ�� = 1 And a.�ⷿid + 0 = [1] And a.������� = b.������� And Rownum = 1 "

'        gstrSQL = "Select Max(�������) As ������� From ҩƷ�շ���¼ " & _
'            " Where ���� = 6 And ���� = 7 And ���ϵ�� = 1 And �ⷿid = [1] And ������� Between Sysdate - 60 And Sysdate "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "mnuEditAddAutoBySale_Click", mlngStockID)
        
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp!����ʱ��) Then
                Me.dtp��ʼʱ��.Value = Format(DateAdd("s", 1, rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss")
                Me.dtp��ʼʱ��.Enabled = False
            End If
        End If
    End If
    mblnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "����������Ϊ�ο�����", Me.chk��������.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "ȷ����������Ϊ����", Me.chk����.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "����������", Me.chk�����.Value)
End Sub

Private Sub lvwSelect_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim intItem As Integer, intItems As Integer, intSelectItems As Integer
    Dim BlnSelect As Boolean
    
    intItems = lvwSelect.ListItems.count
    If Item.Key = "R" Then
        'ȫ���ȫѡ
        If lvwSelect.Tag = "" Then
            BlnSelect = Item.Checked
            For intItem = 2 To intItems
                lvwSelect.ListItems(intItem).Checked = BlnSelect
            Next
        End If
    Else
        lvwSelect.Tag = "1"     '��ʾ����Ҫ�����¼�
        intSelectItems = 0
        For intItem = 2 To intItems
            If lvwSelect.ListItems(intItem).Checked Then intSelectItems = intSelectItems + 1
        Next
        If intSelectItems = intItems - 1 Then
            'ȫѡ
            lvwSelect.ListItems(1).Checked = True
        Else
            'ûѡ�κ�һ��
            lvwSelect.ListItems(1).Checked = False
        End If
        lvwSelect.Tag = ""
    End If
End Sub

Private Sub Get���ʹ�()
    Dim intItem As Integer, intItems As Integer
    mstr���� = ""
    intItems = lvwSelect.ListItems.count
    
    If lvwSelect.ListItems(1).Checked Then
        'ȫѡ
        mstr���� = "1"
    Else
        For intItem = 2 To intItems
            If lvwSelect.ListItems(intItem).Checked Then
                mstr���� = mstr���� & ",'" & Mid(lvwSelect.ListItems(intItem).Key, 2) & "'"
            End If
        Next
        If mstr���� <> "" Then
            mstr���� = "(" & Mid(mstr����, 2) & ")"
        Else
            '��ѡ��ļ���Ϊ��
            mstr���� = "-1"
        End If
    End If
End Sub

Private Sub ResizeDrug()
    Dim blnEnable As Boolean
    '�ж��Ƿ������û�������������
    blnEnable = (optMode(�������쵥δ����) Or optMode(����������) Or optMode(������������) Or optMode(������������))
    lbl������������.Visible = blnEnable
    lbl��ʼʱ��.Visible = blnEnable
    lbl����ʱ��.Visible = blnEnable
    dtp��ʼʱ��.Visible = blnEnable
    dtp����ʱ��.Visible = blnEnable
    
    If blnEnable Then
        lvwSelect.Width = lbl��ʼʱ��.Left - 200 - lvwSelect.Left
    Else
        lvwSelect.Width = fraStep(1).Width - 200 - lvwSelect.Left
    End If
    
    blnEnable = optMode(������������)
    lbl����޶�����.Visible = blnEnable
    lbl��������.Visible = blnEnable
    lbl��������.Visible = blnEnable
    txt��������.Visible = blnEnable
    txt��������.Visible = blnEnable
End Sub

Public Function ShowNavigation(ByVal frmParent As Object, ByVal lngStockid As Long, ByRef intAutoType As Integer, ByRef strEndTime As String, ByRef bln����״̬ As Boolean) As Boolean
    On Error Resume Next
    mlngStockID = lngStockid
    mintAutoType = intAutoType  '1-��ͨ�Զ�����;7-�������Զ�����
    mblnOK = False
    Set mfrmMain = frmParent
    Me.Show 1, frmParent
    ShowNavigation = mblnOK
    intAutoType = mintAutoType
    If Me.chk��������.Value = 0 Then
        bln����״̬ = True
    End If
    If mintAutoType = 7 Then
        strEndTime = Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    End If
End Function

Private Function CheckData() As Boolean
    Dim lng���� As Long
    Dim lngTargetID As Long             'Ŀ��ⷿ��ID
    Dim str���� As String
    Dim rsCheck As New ADODB.Recordset
    Dim str��;ID As String
    Dim n As Integer
    
    '����Ƿ���ڷ��������ļ�¼��ʼ��ֻ�����������бȽϣ�����ֵ�ʱ���ٰ��Ƿ���ȷ��������������Σ�
    On Error GoTo ErrHand
    CheckData = False

    For n = 1 To tvw��;.Nodes.count
        If tvw��;.Nodes(n).Key <> "Root" And _
            tvw��;.Nodes(n).Key <> "_�г�ҩ" And _
            tvw��;.Nodes(n).Key <> "_�в�ҩ" And _
            tvw��;.Nodes(n).Key <> "_����ҩ" And _
            tvw��;.Nodes(n).Checked Then
            str��;ID = str��;ID & "," & Mid(tvw��;.Nodes(n).Key, 2)
        End If
    Next

    If str��;ID <> "" Then
        str��;ID = Mid(str��;ID, 2)
    End If
    
    gstrSQL = ""
    str���� = IIf(mstr���� = "1", "", IIf(mstr���� = "-1", " And 1=2", " And (C.���� IN " & mstr���� & " Or C.���� Is NULL)"))
    lngTargetID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    
    If optMode(����������) Then
        '�����ȷ���Σ���ҩƷ�����û�м�¼��ҩƷ���ݣ�����ȡ����
        gstrSQL = "" & _
                 " Select Distinct Nvl(A.��������,0) ��������,Nvl(B.��������,0) ��������,Nvl(B.ʵ������,0) ʵ������,Nvl(B.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(B.ʵ�ʲ��,0) ʵ�ʲ��, " & _
                 "        D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���, " & _
                 "        D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱� " & _
                 " From (Select �ⷿid, ҩƷid, Sum(Nvl(ʵ������, 0) * Nvl(����, 1)) �������� " & _
                 " From ҩƷ�շ���¼ Where �ⷿid = [2] And ���� In (7,8,9,10,11) And ���ϵ�� = -1 And " & _
                 " ������� Between [3] And [4] Group By �ⷿid, ҩƷid Having Sum(Nvl(ʵ������, 0) * Nvl(����, 1)) > 0) A,ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F,�շ���Ŀ���� E,�շѼ�Ŀ P, ������ĿĿ¼ M,���Ʒ���Ŀ¼ L, " & _
                 "      (Select ҩƷID,Sum(Nvl(��������,0)) ��������,sum(Nvl(ʵ������,0)) ʵ������,Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ�� ,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��" & _
                 "      From ҩƷ��� Where �ⷿID=[1] And ����=1 Group By ҩƷID) B "
        If chk��������С��������.Value = 1 Then
            gstrSQL = gstrSQL & ",(Select ҩƷID,Sum(Nvl(��������,0)) ��������,sum(Nvl(ʵ������,0)) ʵ������,Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ�� ,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��" & _
                " From ҩƷ��� Where �ⷿID=[2] And ����=1 Group By ҩƷID) K "
        End If
        gstrSQL = gstrSQL & "" & _
                 " Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=D.ҩƷID And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist))) " & str���� & _
                 " And D.ҩƷID=F.ID And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� X Where ִ�п���ID=[2] And X.�շ�ϸĿid = D.ҩƷid) " & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� Y Where ִ�п���ID=[1] And Y.�շ�ϸĿid = D.ҩƷid) " & _
                 " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & IIf(chk�����.Value = 1, " And Nvl(b.��������, 0) <> 0", "") & _
                 IIf(chk��������С��������.Value = 1, " And A.ҩƷID=K.ҩƷID(+) And Nvl(A.��������,0)>Nvl(K.��������,0)", "") & _
                 " Order By F.���� "
       Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")), mstr����, str��;ID)
    ElseIf optMode(��������) Then
       gstrSQL = "Select Distinct " & IIf(chk����.Value = 1, "Nvl(A.����,0)", "Nvl(A.����,0)-Sum(Nvl(B.��������,0))") & " ��������,Sum(Nvl(K.��������,0)) ��������,Sum(Nvl(K.ʵ������,0)) ʵ������,Sum(Nvl(K.ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(K.ʵ�ʲ��,0)) ʵ�ʲ��,  " & _
                "         D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                "         D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱� " & _
                "  From (Select ҩƷid,���� From ҩƷ�����޶� Where �ⷿID=[2] And Nvl(����,0)>0" & _
                " ) A, " & _
                "       ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F,�շ���Ŀ���� E,�շѼ�Ŀ P,������ĿĿ¼ M,���Ʒ���Ŀ¼ L , " & _
                "       (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K," & _
                "       (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[1] Group By ִ�п���ID,�շ�ϸĿID) I, " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "       From ҩƷ��� Where �ⷿID=[2] And ����=1" & _
                "       Group by ҩƷID) B,  " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "      From ҩƷ��� Where �ⷿID=[1] And ����=1" & _
                "      Group by ҩƷID) K " & _
                " Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.ҩƷID+0=D.ҩƷID And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=K.ҩƷID(+) And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str���� & _
                " And D.ҩƷID=F.ID And D.ҩƷID=K.�շ�ϸĿID And D.ҩƷID=I.�շ�ϸĿID " & _
                " And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
                " Having Nvl(A.����,0)-Sum(Nvl(B.��������,0))>0 " & IIf(chk�����.Value = 1, " And Sum(Nvl(k.��������, 0))<>0 ", "") & _
                " Group By Nvl(A.����,0),D.ҩƷID,F.����,F.����,E.����,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                " D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ, D.�Ƿ񳣱� "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, mstr����, str��;ID)
    ElseIf optMode(��������) Then
        gstrSQL = "Select Distinct " & IIf(chkLowerLimit(0).Value = 1, "Nvl(A.����,0)", IIf(chkLowerLimit(1).Value = 1, "Nvl(A.����,0)", "Nvl(A.����,0)-Sum(Nvl(B.��������,0))")) & " ��������,Sum(Nvl(K.��������,0)) ��������,Sum(Nvl(K.ʵ������,0)) ʵ������,Sum(Nvl(K.ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(K.ʵ�ʲ��,0)) ʵ�ʲ��,  " & _
                "         D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                "         D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱�  " & _
                "  From (Select ҩƷid,����,���� From ҩƷ�����޶� Where �ⷿID=[2] And Nvl(����,0)>0" & IIf(chkLowerLimit(1).Value = 1, " And Nvl(����,0)>0", "") & _
                " ) A, " & _
                "       ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F, �շ���Ŀ���� E,�շѼ�Ŀ P,������ĿĿ¼ M,���Ʒ���Ŀ¼ L, " & _
                "      (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K," & _
                "      (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[1] Group By ִ�п���ID,�շ�ϸĿID) I, " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "       From ҩƷ��� Where �ⷿID=[2] And ����=1" & _
                "       Group by ҩƷID) B,  " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "      From ҩƷ��� Where �ⷿID=[1] And ����=1" & _
                "      Group by ҩƷID) K " & _
                "  Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in (select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.ҩƷID+0=D.ҩƷID And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=K.ҩƷID(+) And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in(select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str���� & _
                "  And D.ҩƷID=F.ID And D.ҩƷID=K.�շ�ϸĿID And D.ҩƷID=I.�շ�ϸĿID " & _
                "  And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
                "  Having Nvl(A.����,0)-Sum(Nvl(B.��������,0))>0 " & IIf(chk�����.Value = 1, " And Sum(Nvl(k.��������, 0))<>0 ", "") & _
                "  Group By Nvl(A.����,0),Nvl(A.����,0),D.ҩƷID,F.����,F.����,E.����,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                "        D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ, D.�Ƿ񳣱� "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, mstr����, str��;ID)
    ElseIf optMode(����������) Then
        '��ȡ���������޵�ҩƷ����������������д��������
        gstrSQL = "Select Distinct Nvl(A.����,0)-Sum(Nvl(B.��������,0)) As ��������, Nvl(A.����,0)-Sum(Nvl(B.��������,0)) As ������������,Sum(Nvl(K.��������,0)) ��������,Sum(Nvl(K.ʵ������,0)) ʵ������,Sum(Nvl(K.ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(K.ʵ�ʲ��,0)) ʵ�ʲ��,  " & _
                "         D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                "         D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱�  " & _
                "  From (Select ҩƷid,����,���� From ҩƷ�����޶� Where �ⷿID=[2] And Nvl(����,0)>0 And Nvl(����,0)>0" & _
                " ) A, " & _
                "       ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F, �շ���Ŀ���� E,�շѼ�Ŀ P,������ĿĿ¼ M,���Ʒ���Ŀ¼ L," & _
                "      (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K," & _
                "      (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[1] Group By ִ�п���ID,�շ�ϸĿID) I, " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "       From ҩƷ��� Where �ⷿID=[2] And ����=1" & _
                "       Group by ҩƷID) B,  " & _
                "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                "      From ҩƷ��� Where �ⷿID=[1] And ����=1" & _
                "      Group by ҩƷID) K " & _
                "  Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))") & _
                " And A.ҩƷID+0=D.ҩƷID And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=K.ҩƷID(+) And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in(select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & str���� & _
                "  And D.ҩƷID=F.ID And D.ҩƷID=K.�շ�ϸĿID And D.ҩƷID=I.�շ�ϸĿID " & _
                "  And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                GetPriceClassString("P") & _
                " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
                "  Having Nvl(A.����,0)-Sum(Nvl(B.��������,0))>0 " & IIf(chk�����.Value = 1, " And Sum(Nvl(k.��������, 0))<>0 ", "") & _
                "  Group By Nvl(A.����, 0),Nvl(A.����,0),D.ҩƷID,F.����,F.����,E.����,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                "        D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ,D.�Ƿ񳣱� "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, mstr����, str��;ID)
    ElseIf optMode(�������쵥δ����) Then
        '�������쵥δ��������������And Nvl(A.��ҩ��ʽ,0)=1 ������Ϊ���ʱ����ɾ�����쵥�������ƿⵥ����˵ģ���־�Ѿ�û���ˣ�
        gstrSQL = "select Distinct A.��������,Nvl(B.��������,0) ��������,Nvl(B.ʵ������,0) ʵ������,Nvl(B.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(B.ʵ�ʲ��,0) ʵ�ʲ��, " & _
                 "        D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���, " & _
                 "        D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱� " & _
                 " from (Select �ⷿid, ҩƷid, Sum(Nvl(��д����,0) - Nvl(ʵ������,0)) �������� " & _
                 " From ҩƷ�շ���¼ Where �ⷿid = [1] And �Է�����id = [2] And ���� = 6 And " & _
                 " ������� Between [3] And [4] Group By �ⷿid, ҩƷid Having Sum(��д���� - ʵ������) > 0) A,ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F,�շ���Ŀ���� E,�շѼ�Ŀ P,������ĿĿ¼ M,���Ʒ���Ŀ¼ L , " & _
                 "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                 "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                 "      From ҩƷ��� Where �ⷿID=[1] And ����=1" & _
                 "      Group by ҩƷID) B " & _
                 " Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=D.ҩƷID And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in(select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))" & str���� & _
                 " And D.ҩƷID=F.ID And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� X Where ִ�п���ID=[2] And X.�շ�ϸĿid = D.ҩƷid) " & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� Y Where ִ�п���ID=[1] And Y.�շ�ϸĿid = D.ҩƷid) " & _
                 " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & IIf(chk�����.Value = 1, " And Nvl(B.��������,0)<>0 ", "") & _
                 " Order By F.���� "
         Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")), mstr����, str��;ID)
    ElseIf optMode(������������) Then
        'ҩ��������������ã��ȼ������ĳҩ����Ʒ����ƽ����������
        '1������ĳҩ����Ʒ����ƽ��������=����(ʱ�䳤�̵��Զ��壬����Ƚ��鷳���������ó�����ȫ��ʱ��)ĳҩ����Ʒ��ҩƷ�������ܺ�/��������
        '2��ҩ��������� = ����ĳҩ����Ʒ��ҩƷ��ƽ�������� * ��Ҫ�趨�����������(ʱ����Զ���)
        '3��ҩ��������� = ����ĳҩ����Ʒ��ҩƷ��ƽ�������� * ��Ҫ�趨�����������(ʱ����Զ���)
        '4��ҩ���Զ�����ƻ�Ϊ:
        '   (1)��ĳҩ����Ʒ���ֿ����>= ҩ��������ޣ�����������ƻ�
        '   (2)��ĳҩ����Ʒ���ֿ����< ҩ��������ޣ���������ƻ�
        '   (3)��ҩ������ƻ�=��ҩ���������-���п����
        lng���� = CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd")) - CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd")) + 1
        If lng���� <= 0 Then lng���� = 1
            gstrSQL = "Select Distinct Nvl(A.�������,0)-Nvl(B.��������,0) ��������,Nvl(K.��������,0) ��������,Nvl(K.ʵ������,0) ʵ������,Nvl(K.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(K.ʵ�ʲ��,0) ʵ�ʲ��,  " & _
                 "         D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���,  " & _
                 "         D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱�  " & _
                 "  From (SELECT A.ҩƷID,A.������,A.������*" & Val(txt��������.Text) & " AS �������,A.������*" & Val(txt��������.Text) & " AS �������" & _
                 "        FROM " & _
                 "           (SELECT ҩƷID,SUM(NVL(ʵ������,0)*NVL(����,1))/" & lng���� & " AS ������" & _
                 "           FROM ҩƷ�շ���¼ WHERE �ⷿID+0=[2] AND ���� IN (8,9,10)" & _
                 "           AND ������� BETWEEN [3] AND [4] " & _
                 "           GROUP BY ҩƷID) A ) A," & _
                 "       ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F,�շ���Ŀ���� E,�շѼ�Ŀ P,������ĿĿ¼ M,���Ʒ���Ŀ¼ L,"
            gstrSQL = gstrSQL & _
                 "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                 "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                 "       From ҩƷ��� Where �ⷿID=[2] And ����=1" & _
                 "       Group by ҩƷID) B,  " & _
                 "       (Select ҩƷID,Sum(Nvl(��������,0)) ��������,Sum(Nvl(ʵ������,0)) ʵ������, " & _
                 "       Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��   " & _
                  "      From ҩƷ��� Where �ⷿID=[1] And ����=1" & _
                  "      Group by ҩƷID) K " & _
                 "  Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.ҩƷID+0=D.ҩƷID And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=K.ҩƷID(+) And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in(select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist))) " & str���� & _
                 "  And D.ҩƷID=F.ID " & _
                 "  AND Nvl(B.��������,0)<A.������� " & _
                 "  And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� X Where ִ�п���ID=[2] And X.�շ�ϸĿid = D.ҩƷid) " & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� Y Where ִ�п���ID=[1] And Y.�շ�ϸĿid = D.ҩƷid) " & _
                 " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & IIf(chk�����.Value = 1, " And Nvl(K.��������,0)<>0 ", "") & _
                 " Order By F.���� "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")), mstr����, str��;ID)
    ElseIf optMode(������������) Then
        '����ָ��ʱ�䷶Χ�ڵ����ۣ���ҩ��������Ϊ���ε���������
        gstrSQL = "" & _
                 " Select Distinct Nvl(A.��������,0) ��������,Nvl(B.��������,0) ��������,Nvl(B.ʵ������,0) ʵ������,Nvl(B.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(B.ʵ�ʲ��,0) ʵ�ʲ��, " & _
                 "        D.ҩƷID,F.����,F.���� As ͨ����,E.���� As ��Ʒ��,F.�Ƿ���,D.ҩ�����,D.ҩ������,P.�ּ� �ۼ�,F.���,F.����,D.ԭ����,D.���Ч��,D.�ӳ���, " & _
                 "        D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ,F.���㵥λ �ۼ۵�λ, D.�Ƿ񳣱� " & _
                 " From  (Select  �ⷿid, ҩƷid, Sum(Nvl(ʵ������, 0) * Nvl(����, 1)) �������� " & _
                 " From ҩƷ�շ���¼ Where �ⷿid = [2] And ���� In (8, 9, 10) And ���ϵ�� = -1 And " & _
                 " ������� Between [3] And [4] Group By �ⷿid, ҩƷid Having Sum(Nvl(ʵ������, 0) * Nvl(����, 1)) > 0) A,ҩƷ��Ϣ C,ҩƷ��� D,�շ���ĿĿ¼ F,�շ���Ŀ���� E,�շѼ�Ŀ P, ������ĿĿ¼ M,���Ʒ���Ŀ¼ L," & _
                 "      (Select ҩƷID,Sum(Nvl(��������,0)) ��������,sum(Nvl(ʵ������,0)) ʵ������,Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ�� ,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��" & _
                 "      From ҩƷ��� Where �ⷿID=[1] And ����=1 Group By ҩƷID) B " & _
                 " Where D.ҩ��ID=M.ID And M.����ID=L.ID And B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And ����(+) = 1 AND L.���� In (1,2,3) " & IIf(str��;ID = "", "", " And L.ID in(select * from Table(Cast(f_Num2List([6]) As zlTools.t_NumList)))") & _
                 " And A.ҩƷID+0=B.ҩƷID(+) And A.ҩƷID+0=D.ҩƷID And D.ҩ��ID=C.ҩ��ID And D.ҩƷID=P.�շ�ϸĿID And F.��� in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))" & str���� & _
                 " And D.ҩƷID=F.ID And SysDate Between P.ִ������ And Nvl(P.��ֹ����,Sysdate) " & _
                 GetPriceClassString("P") & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� X Where ִ�п���ID=[2] And X.�շ�ϸĿid = D.ҩƷid) " & _
                 " And Exists (Select ִ�п���ID From �շ�ִ�п��� Y Where ִ�п���ID=[1] And Y.�շ�ϸĿid = D.ҩƷid) " & _
                 " And (F.����ʱ�� Is Null Or To_char(F.����ʱ��,'yyyy-MM-dd')='3000-01-01') " & IIf(chk�����.Value = 1, " And Nvl(B.��������,0)<>0 ", "") & _
                 " Order By F.���� "
       Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ڷ��������ļ�¼]", lngTargetID, mlngStockID, CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")), mstr����, str��;ID)
    End If
    
    If rsCheck.RecordCount = 0 Then
        MsgBox "û�ҵ����������ļ�¼��", vbInformation, gstrSysName
        mintStep = mintStep - 1
        Exit Function
    End If
    
    On Error GoTo 0
    Call WriteResult(rsCheck)
    
    Dim intCount As Integer
    With frmRequestDrugCard
        For intCount = 0 To .cboStock.ListCount - 1
            If .cboStock.ItemData(intCount) = lngTargetID Then
                .cboStock.ListIndex = intCount: Exit For
            End If
        Next
    End With
    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub WriteResult(ByVal rsCheck As ADODB.Recordset)
    Dim strUnit As String
    Dim lngTargetID As Long
    Dim blnAdd As Boolean
    Dim bln��ʾ As Boolean, bln�ⷿ As Boolean
    Dim bln���� As Boolean, bln��ҩ As Boolean       'bln����-����ϵͳ����������顱���û������������Ƿ�����޿���ҩƷ��bln��ҩ-��ǰҩƷ�Ƿ���ʱ�ۻ�����ҩƷ
    Dim dbl�������� As Double, dbl��д���� As Double, dbl����ϵ�� As Double
    Dim rsStock As New ADODB.Recordset  'ҩƷ���
    Dim rsTemp  As New ADODB.Recordset
    Dim blnStock As Boolean             '�Ƿ񳣱�ҩƷ
    Dim blnShowMsg As Boolean
        
    On Error GoTo errHandle
    lngTargetID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    Call GetPara(lngTargetID)
    bln�ⷿ = CheckStock(lngTargetID)
    strUnit = GetDrugUnit(mlngStockID, "ҩƷ�������")
    
    Call GetDrugDigit(lngTargetID, "ҩƷ�������", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    '׼���������ݣ�ȫ�������۵�λΪ׼��������SetColValue������ת���������ϵ��Ϊ��ǰ��λ��ϵ����
    With rsCheck
        Do While Not .EOF
            dbl�������� = Calc_Clique(!ҩƷID, !��������)
            'ȷ������ҩƷ
            blnStock = IIf(IsNull(!�Ƿ񳣱�), False, !�Ƿ񳣱�)
            blnStock = Not blnStock
            If optDrugType(1).Value Then
                If blnStock = False Then GoTo Continue
            ElseIf optDrugType(2).Value Then
                If blnStock = True Then GoTo Continue
            End If
            
            If mint�����γ��� = 1 Then
                gstrSQL = " Select Nvl(��������,0) ��������,Nvl(ʵ������,0) ʵ������,Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ��," & _
                          "     Nvl(����,0) ����,Ч��,�ϴ����� ����,�ϴβ��� ����,ԭ����,NVL(�ϴι�Ӧ��ID,0) �ϴι�Ӧ��ID,��׼�ĺ� " & _
                          " From ҩƷ��� Where �ⷿID=[1] And ҩƷID=[2] And ����=1"
                If gtype_UserSysParms.P150_ҩƷ���������㷨 = 0 Then
                    gstrSQL = gstrSQL & " Order by Nvl(����,0)"
                Else
                    gstrSQL = gstrSQL & " Order by Ч��,Nvl(����,0)"
                End If
                
                Set rsStock = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ��ҩƷ�Ŀ��]", lngTargetID, CLng(!ҩƷID))
                
                blnAdd = False
                If rsStock.RecordCount <> 0 Then
                    '�п���ҩƷ��������ʱ��ҩƷ���˲���
                    Do While Not rsStock.EOF
                        If dbl�������� >= rsStock!�������� Then
                            dbl��д���� = IIf(rsStock!�������� > 0, rsStock!��������, 0)
                            'У����д����
                            dbl��д���� = Calc_Clique(!ҩƷID, dbl��д����, True)
                        Else
                            '����ҪУ������Ϊ�������У�����������ж����֧����ʣ������Ȼ�Ƿ���Ҫ��ģ����Բ���У��
                            dbl��д���� = dbl��������
                        End If
                        
                        '��ȷ���Σ�������Ҫ������������д���쵥
                        dbl����ϵ�� = IIf(strUnit = "סԺ��λ", !סԺ��װ, IIf(strUnit = "���ﵥλ", !�����װ, IIf(strUnit = "ҩ�ⵥλ", !ҩ���װ, 1)))
                        If dbl��д���� <> 0 Then
                            
                            If SetColValue(!ҩƷID, "[" & !���� & "]", !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), IIf(IsNull(!���), "", !���), IIf(IsNull(rsStock!����), "", rsStock!����), _
                                IIf(strUnit = "סԺ��λ", !סԺ��λ, IIf(strUnit = "���ﵥλ", !���ﵥλ, IIf(strUnit = "ҩ�ⵥλ", !ҩ�ⵥλ, !�ۼ۵�λ))), _
                                !�ۼ�, IIf(IsNull(rsStock!����), "", rsStock!����), IIf(IsNull(rsStock!Ч��), "", rsStock!Ч��), zlStr.Nvl(!���Ч��, 0), !ҩ�����, IIf(IsNull(!��������), 0, !��������), _
                                IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !�ӳ��� / 100, _
                                IIf(strUnit = "סԺ��λ", !סԺ��װ, IIf(strUnit = "���ﵥλ", !�����װ, IIf(strUnit = "ҩ�ⵥλ", !ҩ���װ, 1))), _
                                rsStock!����, dbl��д����, !ҩ������, !�Ƿ���, zlStr.Nvl(rsStock!�ϴι�Ӧ��ID, 0), _
                                IIf(IsNull(rsStock!��׼�ĺ�), "", rsStock!��׼�ĺ�), blnStock, IIf(IsNull(rsStock!ԭ����), "", rsStock!ԭ����)) Then blnAdd = True
                        End If
                        
                        dbl�������� = dbl�������� - dbl��д����
                        If dbl�������� = 0 Then Exit Do
                        rsStock.MoveNext
                    Loop
                    If dbl�������� > 0 And blnAdd Then
                        'δ�����������ȫ���������һ�е�ҩƷ��
                        If Me.chk��������.Value = 1 Then
                            frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��������) = zlStr.FormatEx(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��������)) + dbl�������� / dbl����ϵ��, mintNumberDigit, , True)
                        Else
                            frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����) = zlStr.FormatEx(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����)) + dbl�������� / dbl����ϵ��, mintNumberDigit, , True)
                        End If
                        
                        If chk����.Value = 1 Then
                            If Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����)) <> Int(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����))) Then
                                frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����) = zlStr.FormatEx(Int(Val(frmRequestDrugCard.mshBill.TextMatrix(frmRequestDrugCard.mshBill.rows - 2, mIntCol��д����))) + 1, mintNumberDigit, , True)
                            End If
                        End If
                    End If
                Else
                    '����������ʱ�����Ե�ҩƷ���˲���
                    '�������Ϊ�����ֹ��������ִ���������
                    If mintCheck <> 2 Then
                        gstrSQL = " Select Nvl(A.ҩ�����,0) ҩ�����,Nvl(A.ҩ������,0) ҩ������,Nvl(B.�Ƿ���,0) ʱ��, a.�ϴι�Ӧ��ID " & _
                                  " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                                  " Where A.ҩƷID = B.ID And A.ҩƷID = [1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ҩƷ���ڳ���ⷿ�Ƿ������ʱ�۵�����]", CLng(!ҩƷID))
                        
                        bln��ҩ = (rsTemp!ʱ�� = 1) Or IIf(bln�ⷿ, (rsTemp!ҩ����� = 1), (rsTemp!ҩ������ = 1))
                        If Not bln��ҩ Then
                            If Not bln��ʾ Then
                                If mintCheck = 1 Then
                                    bln���� = (MsgBox("ҩƷ��治��,�Ƿ�������죿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
                                Else
                                    bln���� = True
                                End If
                                bln��ʾ = True
                            End If
                            If bln���� Then
                                'Ϊ�޿��ҩƷ���������¼
                                If dbl�������� <> 0 Then
                                    Call SetColValue(!ҩƷID, "[" & !���� & "]", !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), IIf(IsNull(!���), "", !���), "", _
                                        IIf(strUnit = "סԺ��λ", !סԺ��λ, IIf(strUnit = "���ﵥλ", !���ﵥλ, IIf(strUnit = "ҩ�ⵥλ", !ҩ�ⵥλ, !�ۼ۵�λ))), _
                                        !�ۼ�, "", "", zlStr.Nvl(!���Ч��, 0), !ҩ�����, IIf(IsNull(!��������), 0, !��������), _
                                        IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !�ӳ��� / 100, _
                                        IIf(strUnit = "סԺ��λ", !סԺ��װ, IIf(strUnit = "���ﵥλ", !�����װ, IIf(strUnit = "ҩ�ⵥλ", !ҩ���װ, 1))), _
                                        0, dbl��������, !ҩ������, !�Ƿ���, IIf(rsTemp Is Nothing, 0, zlStr.Nvl(rsTemp!�ϴι�Ӧ��ID, 0)), "", blnStock, "")
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                '���ݴ����¼����������
                If dbl�������� <> 0 Then
'                    If mintCheck = 2 And dbl�������� > Val(IIf(IsNull(!��������), 0, !��������)) Then
'                        '��治���ֹ
'                        If blnShowMsg = False Then
'                            MsgBox "����ⷿ�������˳����飬��治��ʱ���������������ݡ�", vbInformation, gstrSysName
'                            blnShowMsg = True
'                        End If
'                    Else
                        '�ϴι�Ӧ��ID
                        gstrSQL = "select �ϴι�Ӧ��ID, �ϴβ��� from ҩƷ��� where ҩƷID = [1] "
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-�ϴι�Ӧ��", CLng(!ҩƷID))
                    
                        '��������
                        Call SetColValue(!ҩƷID, "[" & !���� & "]", !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), IIf(IsNull(!���), "", !���), _
                            IIf(rsTemp Is Nothing, zlStr.Nvl(!����), zlStr.Nvl(rsTemp!�ϴβ���)), _
                            IIf(strUnit = "סԺ��λ", !סԺ��λ, IIf(strUnit = "���ﵥλ", !���ﵥλ, IIf(strUnit = "ҩ�ⵥλ", !ҩ�ⵥλ, !�ۼ۵�λ))), _
                            !�ۼ�, "", "", zlStr.Nvl(!���Ч��, 0), !ҩ�����, IIf(IsNull(!��������), 0, !��������), _
                            IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !�ӳ��� / 100, _
                            IIf(strUnit = "סԺ��λ", !סԺ��װ, IIf(strUnit = "���ﵥλ", !�����װ, IIf(strUnit = "ҩ�ⵥλ", !ҩ���װ, 1))), _
                            0, dbl��������, !ҩ������, !�Ƿ���, IIf(rsTemp Is Nothing, 0, zlStr.Nvl(rsTemp!�ϴι�Ӧ��ID, 0)), "", blnStock, zlStr.Nvl(!ԭ����))
'                    End If
                End If
            End If
Continue:
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

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal lngҩƷid As Long, ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal str��� As String, _
    ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�������� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal dbl�ӳ��� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal Dbl���� As Double, ByVal intҩ������ As Integer, ByVal int�Ƿ��� As Integer, _
    ByVal lng�ϴι�Ӧ��ID As Long, ByVal str��׼�ĺ� As String, ByVal bln�Ƿ񳣱� As Boolean, ByVal strԭ���� As String) As Boolean
    
    Dim intDrugNameShow As Integer
    Dim strҩ�� As String
    Dim intCount As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strҩƷ��Դ As String, str����ҩ�� As String
    
    Dim mconIntColҩ��  As Integer    ' = 2
    Dim mconIntCol��Ʒ�� As Integer  '= 3
    Dim mconIntCol��Դ As Integer       '=4
    Dim mconIntCol��� As Integer   '=5
    Dim mconIntCol���   As Integer  '= 6
    Dim mconIntCol��������  As Integer ' = 7
    Dim mconIntCol���Ч��  As Integer   ' = 8
    Dim mconIntCol��������  As Integer   ' = 9
    Dim mconIntcol�ӳ��� As Integer     '= 10
    Dim mconIntColʵ�ʽ�� As Integer    '= 11
    Dim mconIntColʵ�ʲ�� As Integer   ' = 12
    Dim mconIntCol����ϵ�� As Integer    '= 13

    Dim mconIntCol���� As Integer    '= 14
    Dim mconIntCol���� As Integer    '= 15
    Dim mconIntColԭ���� As Integer    '= 16
    Dim mconIntCol��λ As Integer  ' = 17
    Dim mconIntCol���� As Integer    '= 18
    Dim mconIntColЧ�� As Integer     '= 19
    Dim mconIntCol��׼�ĺ� As Integer '= 20
    Dim mconintcol��ǰ��� As Integer
    Dim mconintcol�Է���� As Integer
    Dim mconIntCol��д���� As Integer  ' = 21
    Dim mconIntCol�������� As Integer
    Dim mconIntColʵ������ As Integer  ' = 22
    Dim mconIntCol�ɹ��� As Integer     '= 23
    Dim mconIntCol�ɹ���� As Integer   '= 24
    Dim mconIntCol�ۼ� As Integer   '= 25
    Dim mconIntCol�ۼ۽�� As Integer   ' = 26
    Dim mconintCol��� As Integer   '= 27
    Dim mconIntCol�ϴι�Ӧ��ID As Long '=28
    Dim mconIntColҩƷ��������� As Integer
    Dim mconIntColҩƷ���� As Integer
    Dim mconIntColҩƷ���� As Integer
    Dim mconIntCol����ҩ�� As Integer
    Dim intCol����ҩƷ As Integer

    Dim numʵ������ As Double
    Dim rsTemp As New ADODB.Recordset
    mconIntColҩ�� = 2
    mconIntCol��Ʒ�� = 3
    mconIntCol��Դ = 4
    mconIntCol����ҩ�� = 5
    mconIntCol��� = 6
    mconIntCol��� = 7
    mconIntCol�������� = 8
    mconIntCol���Ч�� = 9
    mconIntCol�������� = 10
    mconIntcol�ӳ��� = 11
    mconIntColʵ�ʽ�� = 12
    mconIntColʵ�ʲ�� = 13
    mconIntCol����ϵ�� = 14

    mconIntCol���� = 15
    mconIntCol���� = 16
    mconIntColԭ���� = 17
    mconIntCol��λ = 18
    mconIntCol���� = 19
    mconIntColЧ�� = 20
    mconIntCol��׼�ĺ� = 21
    mconintcol��ǰ��� = 22
    mconintcol�Է���� = 23
    mconIntCol�������� = 24: mIntCol�������� = mconIntCol��������
    mconIntCol��д���� = 25:  mIntCol��д���� = mconIntCol��д����
    mconIntColʵ������ = 26
    mconIntCol�ɹ��� = 27
    mconIntCol�ɹ���� = 28
    mconIntCol�ۼ� = 29
    mconIntCol�ۼ۽�� = 30
    mconintCol��� = 31
    mconIntCol�ϴι�Ӧ��ID = 32
    
    mconIntColҩƷ��������� = 34
    mconIntColҩƷ���� = 35
    mconIntColҩƷ���� = 36
    intCol����ҩƷ = 37
    
    SetColValue = False
    On Error GoTo errHandle
    intDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ҩƷ������ʾ��ʽ", 0)))
    
    '�����������Ϊ�����˳�
    If IIf(Dbl���� >= num��������, num��������, Dbl����) = 0 And mint�����γ��� = 1 And (int�Ƿ��� = 1 Or lng���� <> 0) Then Exit Function

    gstrSQL = "Select ҩƷ��Դ,����ҩ�� From ҩƷ��� Where ҩƷID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡҩƷ��Դ]", lngҩƷid)
    
    strҩƷ��Դ = zlStr.Nvl(rsTemp!ҩƷ��Դ)
    str����ҩ�� = zlStr.Nvl(rsTemp!����ҩ��)
    
    '�����ȷ����ʱ��ʱ��ҩƷ����������ȡ�ۼ�;
    If mint�����γ��� = 1 And int�Ƿ��� = 1 Then
        num�ۼ� = Get���ۼ�(int�Ƿ��� = 1, lngҩƷid, Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)), lng����)
    End If
    
    With frmRequestDrugCard.mshBill
        intRow = .rows - 1
        .TextMatrix(intRow, 0) = lngҩƷid
        .TextMatrix(intRow, 1) = intRow
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If intDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf intDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = str��Ʒ��
        
        .TextMatrix(intRow, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntColԭ����) = strԭ����
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        
        If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
            '����Ϊ��Ч��
            .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
        End If
        
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ� * num����ϵ��, mintPriceDigit, , True)
        
        If mint�����γ��� <> 1 And int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Getʱ�����ۼ�(lngҩƷid, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), lng����, num����ϵ��), mintPriceDigit, , True)
        End If
        
        .TextMatrix(intRow, mconIntCol��������) = int��������
        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num�������� / num����ϵ��, mintNumberDigit, , True)
        .TextMatrix(intRow, mconIntCol���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = zlStr.FormatEx(numʵ�ʲ��, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = zlStr.FormatEx(numʵ�ʽ��, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntcol�ӳ���) = dbl�ӳ���
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        '�����ʱ��ҩƷ�����ҩƷ,���ܳ�����ǰ�������
        If Me.chk��������.Value = 1 Then
            frmRequestDrugCard.mshBill.ColWidth(mconIntCol��������) = 1100
            frmRequestDrugCard.cmdȫ������.Visible = True
            frmRequestDrugCard.cmdȫ��.Visible = True
            
            If (int�Ƿ��� = 1 Or lng���� <> 0) And mint�����γ��� = 1 Then
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(IIf(Dbl���� >= num��������, num��������, Dbl����) / num����ϵ��, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(Dbl���� / num����ϵ��, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
            End If
        Else
            frmRequestDrugCard.cmdȫ������.Visible = False
            frmRequestDrugCard.cmdȫ��.Visible = False
            
            If (int�Ƿ��� = 1 Or lng���� <> 0) And mint�����γ��� = 1 Then
                .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(IIf(Dbl���� >= num��������, num��������, Dbl����) / num����ϵ��, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(IIf(Dbl���� >= num��������, num��������, Dbl����) / num����ϵ��, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(Dbl���� / num����ϵ��, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(Dbl���� / num����ϵ��, mintNumberDigit, , True)
            End If
        End If
        
        If chk����.Value = 1 Then
            If Val(.TextMatrix(intRow, mconIntCol��д����)) <> Int(Val(.TextMatrix(intRow, mconIntCol��д����))) Then
                .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(Int(Val(.TextMatrix(intRow, mconIntCol��д����))) + 1, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��д����), mintNumberDigit, , True)
            End If
        End If
        
        If .TextMatrix(intRow, mconIntCol�ۼ�) <> "" Then
            .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) * .TextMatrix(intRow, mconIntColʵ������), mintMoneyDigit, , True)
        End If
        
        .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(lngҩƷid, Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)), lng����) * num����ϵ��, mintCostDigit, , True)
        .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(.TextMatrix(intRow, mconIntCol��д����)), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) - Val(.TextMatrix(intRow, mconIntCol�ɹ����)), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = lng�ϴι�Ӧ��ID
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, intCol����ҩƷ) = bln�Ƿ񳣱�
                             
        .rows = .rows + 1
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckStock(ByVal lng�ⷿID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '���ָ���ⷿ��ҩ�⡢ҩ�������Ƽ���(����Ŀⷿ�϶���ҩ�⡢ҩ�����Ƽ����е�һ��)
    On Error GoTo errHandle
    gstrSQL = " Select ����ID From ��������˵�� " & _
              " Where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[�ж��ǲ���ҩ�����Ƽ���]", lng�ⷿID)
              
    If rsCheck.EOF Then
        CheckStock = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPara(ByVal lng�ⷿID As Long)
    Dim rsTemp As New ADODB.Recordset
    '��ȡ������Ĳ�������ֵ��0-�����;1-��飬��������;2-�����ֹ��
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) Value From ҩƷ������ Where �ⷿID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�����Ĳ���]", lng�ⷿID)
    
    If Not rsTemp.EOF Then
        mintCheck = rsTemp!Value
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optMode_Click(index As Integer)
    mintAutoType = index + 1
    
    chk��������С��������.Enabled = False
    chk����.Enabled = False
    chkLowerLimit(0).Enabled = False
    chkLowerLimit(1).Enabled = False
    
    Select Case index
    Case 0
        chk��������С��������.Enabled = True
    Case 1
        chk����.Enabled = True
    Case 2
        chkLowerLimit(0).Enabled = True
        chkLowerLimit(1).Enabled = True
    End Select
    
End Sub

Private Sub tvw��;_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
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

Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
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


