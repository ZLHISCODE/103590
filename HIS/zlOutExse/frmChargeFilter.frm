VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab sst1 
      Height          =   4452
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   6276
      _ExtentX        =   11060
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmChargeFilter.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�շ���Ŀ(&1)"
      TabPicture(1)   =   "frmChargeFilter.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtInput(0)"
      Tab(1).Control(1)=   "ListFeeItem(0)"
      Tab(1).Control(2)=   "tlbOpt(0)"
      Tab(1).Control(3)=   "ils16"
      Tab(1).Control(4)=   "lbl������Ŀ(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "����"
      TabPicture(2)   =   "frmChargeFilter.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmAuditStatus"
      Tab(2).Control(1)=   "cboAudit"
      Tab(2).Control(2)=   "cboApply"
      Tab(2).Control(3)=   "chkDate(1)"
      Tab(2).Control(4)=   "chkDate(0)"
      Tab(2).Control(5)=   "dtpApplyB"
      Tab(2).Control(6)=   "dtpAuditB"
      Tab(2).Control(7)=   "dtpApplyE"
      Tab(2).Control(8)=   "dtpAuditE"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "lblAuditDate"
      Tab(2).Control(12)=   "lblAudit"
      Tab(2).Control(13)=   "lblApplyDate"
      Tab(2).Control(14)=   "lblApply"
      Tab(2).ControlCount=   15
      Begin VB.Frame frmAuditStatus 
         Caption         =   "���״̬"
         Height          =   615
         Left            =   -74085
         TabIndex        =   61
         Top             =   1320
         Width           =   3225
         Begin VB.CheckBox chkAudit 
            Caption         =   "�ܾ�"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   64
            Top             =   240
            Width           =   915
         End
         Begin VB.CheckBox chkAudit 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   63
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkAudit 
            Caption         =   "ͨ��"
            Height          =   255
            Index           =   1
            Left            =   1215
            TabIndex        =   62
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.ComboBox cboAudit 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74085
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2070
         Width           =   2055
      End
      Begin VB.ComboBox cboApply 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74085
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chkDate 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -69480
         TabIndex        =   47
         Top             =   2505
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkDate 
         Height          =   255
         Index           =   0
         Left            =   -69480
         TabIndex        =   46
         Top             =   923
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   18
         ToolTipText     =   "���ƥ��100���������"
         Top             =   540
         Width           =   2160
      End
      Begin VB.ListBox ListFeeItem 
         Height          =   3210
         Index           =   0
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   19
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��,���һ����δѡ���ʾ������"
         Top             =   900
         Width           =   4725
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3972
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   5925
         Begin VB.TextBox txt��ʶ�� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3825
            MaxLength       =   18
            TabIndex        =   9
            Top             =   1350
            Width           =   1830
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   975
            MaxLength       =   64
            TabIndex        =   8
            Top             =   1350
            Width           =   1830
         End
         Begin VB.TextBox txtPatient 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   17
            Top             =   3240
            Width           =   3495
         End
         Begin VB.Frame fra��Դ 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   120
            TabIndex        =   52
            Top             =   3720
            Width           =   5535
            Begin VB.OptionButton opt���� 
               Caption         =   "���ﲡ��"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   55
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "סԺ����"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   2370
               TabIndex        =   54
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "���ﲡ�˺�סԺ����"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   3675
               TabIndex        =   53
               Top             =   0
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.Label lblFil 
               Alignment       =   1  'Right Justify
               Caption         =   "������Դ"
               Height          =   180
               Left            =   0
               TabIndex        =   56
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.TextBox txtFactEnd 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3828
            TabIndex        =   13
            Top             =   2100
            Width           =   1830
         End
         Begin VB.TextBox txtFactBegin 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   972
            TabIndex        =   12
            Top             =   2100
            Width           =   1830
         End
         Begin VB.TextBox txtNoEnd 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3828
            MaxLength       =   8
            TabIndex        =   11
            Top             =   1728
            Width           =   1830
         End
         Begin VB.TextBox txtNOBegin 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   972
            MaxLength       =   8
            TabIndex        =   10
            Top             =   1728
            Width           =   1830
         End
         Begin VB.ComboBox cbo����Ա 
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   972
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2832
            Width           =   1830
         End
         Begin VB.CheckBox chk�˷� 
            Caption         =   "�˷Ѽ�¼"
            Height          =   210
            Left            =   4695
            TabIndex        =   5
            Top             =   555
            Width           =   1020
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   972
            TabIndex        =   14
            Text            =   "cbo����"
            Top             =   2472
            Width           =   1830
         End
         Begin VB.CheckBox chk�շ� 
            Caption         =   "�շѼ�¼"
            Height          =   210
            Left            =   4695
            TabIndex        =   3
            Top             =   270
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkҽ�� 
            Caption         =   "ҽ���շ�"
            Height          =   195
            Left            =   3480
            TabIndex        =   2
            Top             =   278
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk��ͨ 
            Caption         =   "��ͨ�շ�"
            Height          =   195
            Left            =   3480
            TabIndex        =   4
            ToolTipText     =   "ָ������ҽ���շѵ����������շ�"
            Top             =   563
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.ComboBox cbo���ʽ 
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   3825
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1020
            Width           =   1830
         End
         Begin VB.ComboBox cbo�ѱ� 
            Height          =   276
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1020
            Width           =   1830
         End
         Begin VB.TextBox txt������ 
            Height          =   300
            Left            =   3828
            TabIndex        =   15
            Top             =   2472
            Width           =   1830
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   975
            TabIndex        =   1
            Top             =   570
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   173342723
            CurrentDate     =   36588
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   975
            TabIndex        =   0
            Top             =   150
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   173342723
            CurrentDate     =   36588
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   960
            TabIndex        =   58
            Top             =   3240
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   529
            Appearance      =   2
            IDKindStr       =   "ҽ|ҽ����|0|0|0|0|0|;��|���֤��|0|0|0|0|0|;IC|IC����|1|0|0|0|0|;��|���￨|0|0|0|0|0|"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "����"
            IDKind          =   -1
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl��ʶ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����"
            Height          =   180
            Left            =   3024
            TabIndex        =   60
            Top             =   1410
            Width           =   768
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   540
            TabIndex        =   59
            Top             =   1410
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ʶ��"
            Height          =   180
            Left            =   120
            TabIndex        =   57
            Top             =   3312
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ�ݺ�"
            Height          =   180
            Left            =   360
            TabIndex        =   35
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   3288
            TabIndex        =   34
            Top             =   2160
            Width           =   180
         End
         Begin VB.Label lbl����Ա 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ա"
            Height          =   180
            Left            =   360
            TabIndex        =   33
            Top             =   2892
            Width           =   540
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   180
            TabIndex        =   32
            Top             =   2532
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ݺ�"
            Height          =   180
            Left            =   360
            TabIndex        =   31
            Top             =   1788
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   3288
            TabIndex        =   30
            Top             =   1788
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   180
            TabIndex        =   29
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   180
            TabIndex        =   28
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lbl�ѱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�"
            Height          =   180
            Left            =   540
            TabIndex        =   27
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ�Ƹ���"
            Height          =   180
            Left            =   3075
            TabIndex        =   26
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   3252
            TabIndex        =   25
            Top             =   2532
            Width           =   540
         End
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   600
         Index           =   0
         Left            =   -74760
         TabIndex        =   36
         Top             =   1140
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1058
         ButtonWidth     =   1614
         ButtonHeight    =   1058
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ƴ�(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "�Ƴ���ǰѡ����б���"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "����б���Ŀ"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "����ѡ����б���Ŀ"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   -69960
         Top             =   660
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
               Picture         =   "frmChargeFilter.frx":0054
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":03EE
               Key             =   "Insert"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":0788
               Key             =   "Clear"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":0B22
               Key             =   "Delete"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpApplyB 
         Height          =   300
         Left            =   -74085
         TabIndex        =   39
         Top             =   900
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpAuditB 
         Height          =   300
         Left            =   -74085
         TabIndex        =   44
         Top             =   2475
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpApplyE 
         Height          =   300
         Left            =   -71640
         TabIndex        =   40
         Top             =   900
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpAuditE 
         Height          =   300
         Left            =   -71640
         TabIndex        =   45
         Top             =   2475
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   -71880
         TabIndex        =   51
         Top             =   2535
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   -71880
         TabIndex        =   50
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblAuditDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
         Height          =   180
         Left            =   -74880
         TabIndex        =   49
         Top             =   2535
         Width           =   720
      End
      Begin VB.Label lblAudit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   -74700
         TabIndex        =   48
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label lblApplyDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   -74880
         TabIndex        =   42
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   -74700
         TabIndex        =   41
         Top             =   555
         Width           =   540
      End
      Begin VB.Label lbl������Ŀ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շ���Ŀ(&F)"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   6525
      TabIndex        =   23
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6525
      TabIndex        =   21
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6525
      TabIndex        =   20
      Top             =   270
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String 'IN
Public mblnApply As Boolean '�鿴�˷����뵥
Public mstrFilter As String 'IN/Out
Public mblnDateMoved As Boolean 'Out
Public mstrFeeItems As String 'out

Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnNotClick As Boolean
Private mblnUnChange  As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mrsDept As ADODB.Recordset

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo����Ա.hWnd, KeyAscii)
        If lngIdx = -1 And cbo����Ա.ListCount > 0 Then lngIdx = 0
        cbo����Ա.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
'    If KeyAscii >= 32 Then
'        lngIdx = zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo����.ListCount > 0 Then lngIdx = 0
'        cbo����.ListIndex = lngIdx
'    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo����.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Set mrsDept = GetDepartments("'�ٴ�','����'", gint������Դ & ",3")
    If zlSelectDept(Me, 1120, cbo����, mrsDept, cbo����.Text, True, "���п���") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub chkAudit_Click(Index As Integer)
    If Not Visible Then Exit Sub
    
    If chkAudit(0).Value = 0 And chkAudit(1).Value = 0 And chkAudit(2).Value = 0 Then
        chkAudit(Index).Value = 1   '�ݹ����
        Exit Sub
    End If
    
    If Index = 1 Or Index = 2 Then
        cboAudit.Enabled = chkAudit(1).Value = 1 Or chkAudit(2).Value = 1
        dtpAuditB.Enabled = cboAudit.Enabled
        dtpAuditE.Enabled = cboAudit.Enabled
        chkDate(1).Enabled = cboAudit.Enabled
        
        If cboAudit.Enabled Then chkDate(1).Value = 1
    End If
End Sub

Private Sub chkDate_Click(Index As Integer)
    If Not Visible Then Exit Sub
    
    If chkDate(0).Value = 0 And (chkDate(1).Value = 0 Or chkDate(1).Enabled = False) Then
        If chkDate(1).Enabled = False Then
            chkDate(0).Value = 1
            Exit Sub
        Else
            chkDate(1 - Index).Value = 1 '�ݹ����
        End If
    End If
        
    If Index = 0 Then
        dtpApplyB.Enabled = chkDate(Index).Value = 1
        dtpApplyE.Enabled = dtpApplyB.Enabled
    Else
        dtpAuditB.Enabled = chkDate(Index).Value = 1
        dtpAuditE.Enabled = dtpAuditB.Enabled
    End If
End Sub

Private Sub chk��ͨ_Click()
    If chkҽ��.Value = 0 And chk��ͨ.Value = 0 Then
        chk��ͨ.Value = 1
    End If
End Sub

Private Sub chk�˷�_Click()
    If chk�շ�.Value = 0 And chk�˷�.Value = 0 Then
        chk�˷�.Value = 1
    End If
End Sub

Private Sub chk�շ�_Click()
    If chk�շ�.Value = 0 And chk�˷�.Value = 0 Then
        chk�շ�.Value = 1
    End If
End Sub

Private Sub chkҽ��_Click()
    If chkҽ��.Value = 0 And chk��ͨ.Value = 0 Then
        chkҽ��.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If gbln�˷�����ģʽ And mblnApply Then
        If chkDate(0).Value = 0 And chkDate(1).Value = 0 Then
            MsgBox "��ѡ��ʱ�䷶Χ��", vbInformation, gstrSysName
            chkDate(0).SetFocus: Exit Sub
        End If
        fra��Դ.Visible = False
    Else
        If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
            If txtNoEnd.Text < txtNOBegin.Text Then
                MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
                txtNoEnd.SetFocus: Exit Sub
            End If
        End If
        fra��Դ.Visible = True
        If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
            If txtFactEnd.Text < txtFactBegin.Text Then
                MsgBox "����Ʊ�ݺŲ���С�ڿ�ʼƱ�ݺţ�", vbInformation, gstrSysName
                txtFactEnd.SetFocus: Exit Sub
            End If
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    If gbln�˷�����ģʽ And mblnApply Then
        fra��Դ.Visible = False '33789
        cboApply.SetFocus
    Else
        fra��Դ.Visible = True '33789
        dtpBegin.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If sst1.Tab = 1 Then
            txtInput(sst1.Tab - 1).SetFocus
        Else
            KeyCode = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf Shift = 2 Then
        If sst1.Tab = 1 Then
            Dim i As Integer, Index As Integer
            
            Index = sst1.Tab - 1
            If UCase(Chr(KeyCode)) = "A" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = True
                Next
            ElseIf UCase(Chr(KeyCode)) = "C" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = False
                Next
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
End Sub



Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngOldID As Long, strListFeeItem As String
    Dim Curdate As Date, Index As Integer, arrItem As Variant
    
    On Error GoTo errH
    gblnOK = False
    
    Curdate = zlDatabase.Currentdate
    '47928
    InitIDKind
    
    If gbln�˷�����ģʽ And mblnApply Then
        sst1.TabVisible(0) = False
        sst1.TabVisible(1) = False
        sst1.TabVisible(2) = True
        
        cboApply.Clear
        
        cboApply.AddItem "����������"
        Set rsTmp = GetPersonnel("�����շ�Ա", True)
        For i = 1 To rsTmp.RecordCount
            cboApply.AddItem rsTmp!���� & "-" & rsTmp!����
            
            If rsTmp!ID = UserInfo.ID Then cboApply.ListIndex = cboApply.NewIndex
            rsTmp.MoveNext
        Next
        
        cboAudit.AddItem "���������"
        strSQL = "Select Distinct D.ID, D.����, D.����" & vbNewLine & _
                "From ��Ա�� D,�ϻ���Ա�� C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
                "Where A.ϵͳ = [1] And A.��� = 1121 And A.���� = '�˷����' And A.��ɫ = B.��ɫ And B.�û� = C.�û��� And C.��Աid = D.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)
        For i = 1 To rsTmp.RecordCount
            cboAudit.AddItem rsTmp!���� & "-" & rsTmp!����
            
            If rsTmp!ID = UserInfo.ID Then cboAudit.ListIndex = cboAudit.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboApply.ListIndex = -1 And cboApply.ListCount > 0 Then cboApply.ListIndex = 0
        If cboAudit.ListIndex = -1 And cboAudit.ListCount > 0 Then cboAudit.ListIndex = 0
        
        
        dtpApplyB.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpApplyE.MaxDate = dtpApplyB.MaxDate
        
        dtpAuditB.MaxDate = dtpApplyB.MaxDate
        dtpAuditE.MaxDate = dtpApplyB.MaxDate
        
        dtpApplyB.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpApplyE.Value = dtpApplyB.MaxDate
        
        dtpAuditB.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpAuditE.Value = dtpApplyB.MaxDate
        
        chkAudit(0).Value = 1
        chkAudit(1).Value = 0
        chkAudit(2).Value = 0
        
        chkDate(0).Value = 1
        chkDate(1).Value = 1
        
    Else
        sst1.TabVisible(0) = True
        sst1.TabVisible(1) = True
        sst1.TabVisible(2) = False
        
        If glngSys Like "8??" Then
            lbl����.Visible = False
            cbo����.Visible = False
        End If
        
        txtNOBegin.Text = ""
        txtNoEnd.Text = ""
        txtFactBegin.Text = ""
        txtFactEnd.Text = ""
        txtPatient.Text = ""
        chk�շ�.Value = 1
        chk�˷�.Value = 0
        
        chkҽ��.Value = 1
        chk��ͨ.Value = 1
        
        lbl��ʶ��.Caption = IIf(gint������Դ = 1, "�����", "סԺ��")
        
        dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = dtpBegin.MaxDate
        
        If InStr(mstrPrivs, "��ʾ������") = 0 Then
            lbl������.Visible = False
            txt������.Visible = False
        Else
            lbl������.Visible = True
            txt������.Visible = True
        End If
        
        '����Ա
        cbo����Ա.Clear
        If InStr(mstrPrivs, "���в���Ա") > 0 Then
            cbo����Ա.AddItem "�����շ�Ա"
            Set rsTmp = GetPersonnel("�����շ�Ա", True)
            For i = 1 To rsTmp.RecordCount
                cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����Ա.ItemData(cbo����Ա.NewIndex) = rsTmp!ID
                If rsTmp!ID = UserInfo.ID Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
                rsTmp.MoveNext
            Next
        Else
            cbo����Ա.AddItem UserInfo.���� & "-" & UserInfo.����
            cbo����Ա.ItemData(cbo����Ա.NewIndex) = UserInfo.ID
        End If
        If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount > 0 Then cbo����Ա.ListIndex = 0
        
        '��������'@
        cbo����.Clear
        cbo����.AddItem "���п���"
        cbo����.ListIndex = 0
        Set mrsDept = GetDepartments("'�ٴ�','����'", gint������Դ & ",3")
        For i = 1 To mrsDept.RecordCount
            If lngOldID <> mrsDept!ID Then
                cbo����.AddItem mrsDept!���� & "-" & mrsDept!����
                cbo����.ItemData(cbo����.NewIndex) = mrsDept!ID
                lngOldID = mrsDept!ID
            End If
            mrsDept.MoveNext
        Next
        
        cbo�ѱ�.Clear
        cbo�ѱ�.AddItem "���зѱ�"
        cbo�ѱ�.ListIndex = 0
        
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Where Nvl(�������,3) IN(1,3) Order by ����"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
        
        'ҽ�Ƹ��ʽ,Ĭ��Ϊ�ձ�ʾ����
        cbo���ʽ.Clear
        cbo���ʽ.AddItem ""
        cbo���ʽ.ListIndex = 0
        strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ Order by ����"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
        
        If InStr(1, mstrPrivs, "��ϸ��Ŀ����") = 0 Then
            sst1.TabVisible(1) = False
        Else
            For Index = 0 To 0  '�������ܻ��������Ŀ����
                strListFeeItem = ""
                ListFeeItem(Index).Clear
                
                Call GetRegisterItem(g˽��ģ��, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "�շ���Ŀ�б�", "������Ŀ�б�"), strListFeeItem)
                If strListFeeItem <> "" Then
                    arrItem = Split(strListFeeItem, ";")
                    
                    For i = 0 To UBound(arrItem)
                        ListFeeItem(Index).AddItem Split(arrItem(i), ",")(0)
                        ListFeeItem(Index).ItemData(ListFeeItem(Index).NewIndex) = Val(Split(arrItem(i), ",")(1))
                        ListFeeItem(Index).Selected(ListFeeItem(Index).NewIndex) = IIf(Val(Split(arrItem(i), ",")(2)) = 1, True, False)
                    Next
                End If
            Next
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If gbln�˷�����ģʽ And mblnApply Then
        sst1.Height = dtpAuditB.Top + dtpAuditB.Height * 2
        Me.Height = sst1.Height + dtpAuditB.Height * 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mblnApply = False
    mstrFilter = ""
    mstrFeeItems = ""
    If Not mrsDept Is Nothing Then Set mrsDept = Nothing
End Sub

Private Sub opt����_Click(Index As Integer)
    lbl��ʶ��.Caption = IIf(opt����(0).Value, "�����", IIf(opt����(1).Value, "סԺ��", "����/סԺ��"))
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    If Me.Visible = False Then Exit Sub
    
    If gbln�˷�����ģʽ And mblnApply Then
        If cboApply.Visible And cboApply.Enabled Then Call cboApply.SetFocus
    Else
        If sst1.Tab = 0 Then
            txtPatient.SetFocus
        Else
            txtInput(0).SetFocus
        End If
    End If
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub


Private Sub tlbOpt_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Delete"
            If ListFeeItem(Index).ListIndex >= 0 Then
                Call ListFeeItem(Index).RemoveItem(ListFeeItem(Index).ListIndex)
            End If
        Case "Clear"
            ListFeeItem(Index).Clear
        Case "Save"
            Dim strTmp As String, i As Long
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    strTmp = strTmp & ";" & .List(i) & "," & .ItemData(i) & "," & IIf(.Selected(i), 1, 0)
                Next
            End With
            strTmp = Mid(strTmp, 2)
            Call SaveRegisterItem(g˽��ģ��, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "�շ���Ŀ�б�", "������Ŀ�б�"), strTmp)
    End Select
End Sub

Private Sub ListFeeItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If ListFeeItem(Index).ListIndex >= 0 Then
            ListFeeItem(Index).RemoveItem ListFeeItem(Index).ListIndex
        End If
    End If
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInput(Index))
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, strInput As String, strMatch As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strInput = UCase(Trim(txtInput(Index).Text))
        If strInput = "" Then Exit Sub
        strMatch = IIf(Len(strInput) < 3, "", gstrLike)
        
        If Index = 0 Then
        '�շ���Ŀ
            If zlCommFun.IsNumOrChar(strInput) Then
                strIF = " And (A.���� like [1] Or B.���� like [1] And B.���� in(3," & gbytCode + 1 & "))"
            Else
                strIF = " And B.���� like [1]"
            End If
            strSQL = "Select Distinct A.ID, A.����, B.���� ,A.���, A.����, A.���㵥λ " & _
                  " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.id=B.�շ�ϸĿID " & strIF & _
                  " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                  " And rownum<101 Order by ����"
        Else
        '������Ŀ
            If zlCommFun.IsNumOrChar(strInput) Then
                If IsNumeric(strInput) Then
                    strIF = " And ���� like [1]"
                Else
                    strIF = " And ���� like [1]"
                End If
            Else
                strIF = " And ���� like [1]"
            End If
            
            strSQL = "Select ID, ����, ���� From ������Ŀ Where ĩ��=1 " & strIF & _
                " And rownum<101 Order by ����"
        End If
        
        On Error GoTo errH
        vRect = zlControl.GetControlRect(txtInput(Index).hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ŀѡ��", 1, "", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtInput(Index).Height, blnCancel, False, True, strMatch & strInput & "%")
        If Not rsTmp Is Nothing Then
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = rsTmp!ID Then
                        txtInput(Index).SetFocus
                        txtInput(Index).SelStart = 0
                        txtInput(Index).SelLength = Len(txtInput(Index).Text)
                        Exit Sub
                    End If
                Next
                If .ListCount < 100 Then
                    If Index = 0 Then
                        .AddItem rsTmp!���� & "-" & rsTmp!���� & "(" & rsTmp!��� & ")"
                    Else
                        .AddItem rsTmp!���� & "-" & rsTmp!����
                    End If
                    .ItemData(.NewIndex) = rsTmp!ID
                    .Selected(.NewIndex) = True
                Else
                    MsgBox "�������ܿ���,������Ŀ���ֻ�������100��!", vbInformation, gstrSysName
                End If
            End With
        End If
        
        txtInput(Index).SetFocus
        txtInput(Index).SelStart = 0
        txtInput(Index).SelLength = Len(txtInput(Index).Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.hWnd, KeyAscii)
        If lngIdx = -1 And cbo�ѱ�.ListCount > 0 Then lngIdx = 0
        cbo�ѱ�.ListIndex = lngIdx
    End If
End Sub

Private Sub MakeFilter()
    Dim strSQL As String, Index As Integer, i As Long, strIDs As String
    Dim strSQLtmp As String
    Dim strNotAudit As String 'δ���
    Dim strAudit As String '��˻�ܾ�
    Dim strAuditStatus As String '״̬
    
    If gbln�˷�����ģʽ And mblnApply Then
    
        mstrFilter = ""
        If cboApply.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.������=[1]"
        If dtpApplyB.Enabled Then mstrFilter = mstrFilter & " And A.����ʱ�� Between [2] And [3]"

        'cboAudit.Enabled:53990
        If cboAudit.ListIndex > 0 And cboAudit.Enabled Then strAudit = " And  A.�����=[4] "
        If dtpAuditB.Enabled Then strAudit = strAudit & " And A.���ʱ�� Between [5] And [6] "

        '54391
        strAuditStatus = ""
        If chkAudit(0).Value = 1 Then '����
             strNotAudit = " And A.����� is Null "
             strAuditStatus = strAuditStatus & "," & "0"
        End If
        If chkAudit(1).Value = 1 Or chkAudit(2).Value = 1 Then '��˻�ܾ�
            strAudit = strAudit & " And A.����� is Not Null"
            If chkAudit(1).Value = 1 Then strAuditStatus = strAuditStatus & "," & "1"
            If chkAudit(2).Value = 1 Then strAuditStatus = strAuditStatus & "," & "2"
        End If
        
        If strAudit <> "" Then
            strAudit = IIf(strNotAudit <> "", " OR (1 = 1 " & strAudit & " )", strAudit)
        End If
        mstrFilter = mstrFilter & strNotAudit & strAudit
        
        '�������״̬
        If strAuditStatus = "" Then
            strAuditStatus = "0": chkAudit(0).Value = 1
        Else
            strAuditStatus = Mid(strAuditStatus, 2)
        End If
        mstrFilter = mstrFilter & " And Nvl(A.״̬,0) in(" & strAuditStatus & ")"
    Else
        mstrFilter = " And a.�Ǽ�ʱ�� Between [1] And [2] "
        
        '����ʾ�˷�ʱ���漰�����ݱ�,���۵�û��ת�������ݱ�
        If chk�շ�.Value = 1 Then
            mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
        Else
            mblnDateMoved = False
            '����Ҫ���,�����˳����½���ʱ,�������ϴε�ֵ
        End If
        
        If cbo�ѱ�.ListIndex > 0 Then
            strSQL = "Select Distinct NO From ������ü�¼ Where ��¼����=1 And ��¼״̬<>0 And �ѱ�=[3]" & mstrFilter
            mstrFilter = mstrFilter & " And a.NO IN(" & strSQL & ")"
            '����һ�ŵ��ݵĶ����ж��ַѱ�,������ɸѡ�ķѱ��NO��Ӧ�ó���,���Բ������������ַ�ʽ
            'mstrFilter = mstrFilter & " And �ѱ�='" & txt�ѱ�.Text & "'"
        End If
        
        strSQL = ""
        If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
            strSQL = " And a.NO Between [4] And [5]"
        ElseIf txtNOBegin.Text <> "" Then
            strSQL = " And a.NO=[4]"
        End If
        If strSQL <> "" Then
            mstrFilter = mstrFilter & _
                " And a.����id In (Select Nvl(c.����id, b.����id)" & vbNewLine & _
                "          From ������ü�¼ A, ����Ԥ����¼ B, ����Ԥ����¼ C" & vbNewLine & _
                "          Where a.����id = b.����id And b.������� = c.�������(+) And Mod(a.��¼����, 10) = 1" & vbNewLine & _
                strSQL & ")"
        End If
        
        '���ﲡ���շѵ��ĸ��ʽ�˴���¼���Ƿѱ����
        If cbo���ʽ.ListIndex <> -1 And cbo���ʽ.Text <> "" Then
            mstrFilter = mstrFilter & " And a.���ʽ=[6]"   '��������:33789ʱ,ȡ����.(����δ�Ǽ�,�������Ա(����˵�˵�))
        End If
        
        If cbo����Ա.ListIndex = -1 Then
            mstrFilter = mstrFilter & " And a.����Ա����||''=[7]"
        ElseIf cbo����Ա.ItemData(cbo����Ա.ListIndex) > 0 Then
            mstrFilter = mstrFilter & " And a.����Ա����||''=[7]"
        End If
        
        
        If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
            If Val(Nvl(mrsInfo!ID)) = mlngPrePatient Then
                mstrFilter = mstrFilter & " And a.����ID=[19]"
            End If
        End If
    
       If txt����.Text <> "" Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtPatient.Text, 1))) > 0 Then
                mstrFilter = mstrFilter & " And Upper(a.����) Like [8]"
            Else
                mstrFilter = mstrFilter & " And a.���� Like [8]"
            End If
        End If
            
        If txt��ʶ��.Text <> "" Then mstrFilter = mstrFilter & " And a.��ʶ��=[9]"
        
        strSQL = ""
        If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
            '�������Ʊ�ݺ��ж�,ֱ�Ӹ��ݵ��ݵĵǼ�ʱ���ж�
            strSQLtmp = IIf(txtFactEnd.Text = "", " =[10] ", " Between [10] And [11] ")
            strSQL = "Select A.NO" & _
            " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
            " Where A.��������=1 And A.ID=B.��ӡID And B.Ʊ��=1 And B.����=1" & _
            " And B.���� " & strSQLtmp
        End If
        If strSQL <> "" Then
            mstrFilter = mstrFilter & _
                " And a.����id In (Select Nvl(c.����id, b.����id)" & vbNewLine & _
                "          From ������ü�¼ A, ����Ԥ����¼ B, ����Ԥ����¼ C" & vbNewLine & _
                "          Where a.����id = b.����id And b.������� = c.�������(+) And Mod(a.��¼����, 10) = 1" & vbNewLine & _
                "                And a.NO IN(" & strSQL & ")" & vbNewLine & _
                            ")"
        End If
        
        'ҩ��̶�Ϊ���п���
        If Not glngSys Like "8??" Then
            If cbo����.ListIndex <> 0 Then
                mstrFilter = mstrFilter & " And a.��������ID+0=[12]"
            End If
            If txt������.Text <> "" Then
                mstrFilter = mstrFilter & " And a.������=[17]"
            End If
        End If
        
        If InStr(1, mstrPrivs, "��ϸ��Ŀ����") > 0 Then
            For Index = 0 To 0      '�������ܻ��������Ŀ����
                strIDs = ""
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    If ListFeeItem(Index).Selected(i) Then
                        strIDs = strIDs & "," & ListFeeItem(Index).ItemData(i)
                    End If
                Next
                If strIDs <> "" Then
                    strIDs = Mid(strIDs, 2)
                    If Index = 0 Then
                        mstrFeeItems = strIDs
                        mstrFilter = mstrFilter & " And Instr(','||[18]||',',','||a.�շ�ϸĿID||',')>0"
                    'Else
                        'mstrIncomeItems = strIDs
                        'mstrFilter = mstrFilter & " And Instr(','||[10]||',',','||a.������ĿID||',')>0"
                    End If
                End If
            Next
        End If
    End If
End Sub
 

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

 
Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_Validate(Cancel As Boolean)
    Dim strDoctor As String
    strDoctor = UCase(Trim(txt������.Text))
    If strDoctor <> "" Then
        If zlCommFun.IsNumOrChar(strDoctor) Then
            strDoctor = GetDoctorName(strDoctor)
        End If
    End If
    txt������.Text = strDoctor
End Sub

Private Function GetDoctorName(ByVal strCode As String) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIF As String, lngDept As Long, blnCancel As Boolean, vRect As RECT
    
    If zlCommFun.IsCharAlpha(strCode) Then
        strIF = " And ���� Like [1]"
        strCode = strCode & "%"
    Else
        strIF = " And (���� = [1] Or ��� = [1])"
    End If
    If cbo����.ListIndex > 0 Then
        strIF = strIF & " And B.����ID = [2]"
        lngDept = cbo����.ItemData(cbo����.ListIndex)
    End If
    
    strSQL = "Select Distinct A.Id,A.����" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� C, ��������˵�� D" & vbNewLine & _
            "Where A.ID = B.��Աid And A.ID = C.��Աid And B.����id = D.����id And C.��Ա���� In ('ҽ��', '��ʿ') And D.������� In (" & gint������Դ & ", 3) And" & vbNewLine & _
            "      D.�������� In ('�ٴ�','����')" & strIF

    vRect = zlControl.GetControlRect(txt������.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ��ҽ��", 1, "", "��ѡ��ҽ��", False, False, True, vRect.Left, vRect.Top, txt������.Height, blnCancel, False, True, strCode, lngDept)
    If Not rsTmp Is Nothing Then
        GetDoctorName = rsTmp!����
    End If
End Function

 

 

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function
'��ȡĬ��IDKind����
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
            End If
     End Select
End Function
                
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    If mlngPrePatient Then txtPatient.PasswordChar = ""
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        Exit Sub
    End If
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If Trim(txtPatient.Text) <> "" Then Call FindPati(objCard, False, Trim(txtPatient.Text))
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    '����:60010
    If txtPatient.Locked Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    If objCard.���� Like "���֤*" And objCard.ϵͳ Then
        txtPatient.Text = objPatiInfor.���֤��
    Else
        txtPatient.Text = objPatiInfor.����
    End If
    If Trim(txtPatient.Text) <> "" Then Call FindPati(objCard, False, Trim(txtPatient.Text))
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���أ����ҳɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSQL = ""
    mlngPrePatient = 0
    If blnCard And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then   '103563
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If lng����ID <= 0 Then lng����ID = 0
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]" & str����Ժ
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And B.סԺ��=[2]" & str����Ժ
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!���� Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtPatient.Text = ""
                        Set mrsInfo = Nothing: Exit Function
                    Else
                       strSQL = strSQL & " And  B.���� Like [3]"
                       
                    End If
                Else
                    strSQL = strSQL & " And B.����ID=[2]"
                    strInput = "-" & Val(mrsInfo!����ID)
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '������ҽ������Ч:������:����:26982
                    strSQL = strSQL & " And B.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.ҽ����=[1]" & str����Ժ
                End If
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                ' strSQL = strSQL & " And B.���֤��=[1] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.�����=[1]" & str����Ժ
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.סԺ��=[1]" & str����Ժ
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    If lng����ID = 0 Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID <= 0 Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select distinct  B.����id As ID, Decode(sign(nvl(ylkxx.����id,0)),0,'','��') as �����˻�, B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSQL = strSQL & vbNewLine & "      A.���� ��������"
    strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A,ҽ�ƿ���� YLK,����ҽ�ƿ���Ϣ YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) and b.����id=ylkxx.����id(+) and ylkxx.״̬(+)=0 and  ylkxx.�����id=ylk.id(+)  and ylk.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
    
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(Nvl(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtPatient.Text = Nvl(mrsInfo!����)
    Me.txtPatient.Tag = Nvl(mrsInfo!ID)
    mlngPrePatient = Val(Nvl(mrsInfo!ID))
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    GetPatient = True
    Exit Function
ClearPati:
    txtPatient.Text = ""
    txtPatient.PasswordChar = ""
    Set mrsInfo = Nothing
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub


Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then IDKind.SetAutoReadCard True
End Sub


Private Sub txtPatient_LostFocus()
    IDKind.SetAutoReadCard False
End Sub

 

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "����") Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(txtPatient.Tag) <> 0 Then    '����
                 zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not GetPatient(objCard, strInput, blnCard) Then Exit Sub
End Sub

