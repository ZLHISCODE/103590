VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm���ŷ�ҩ�������� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5775
   ClientLeft      =   8805
   ClientTop       =   3960
   ClientWidth     =   6735
   Icon            =   "Frm���ŷ�ҩ��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1100
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   5010
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "Frm���ŷ�ҩ��������.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblNote(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl��ҩҩ��"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblNote(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl����ģʽ"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Cbo������"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Cbo��ҩҩ��"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Cbo����ģʽ"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk�����һ�����ʾ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkDetailPage"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra��ҩ���ݼ��"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "����(&2)"
      TabPicture(1)   =   "Frm���ŷ�ҩ��������.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frm��ΣҩƷ����"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbo��ҩ�嵥"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cbo��ҩ�嵥"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fra�豸����"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblName"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbl��ҩ�嵥"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl��ҩ�嵥"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "����(&3)"
      TabPicture(2)   =   "Frm���ŷ�ҩ��������.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "��ҩ����(&4)"
      TabPicture(3)   =   "Frm���ŷ�ҩ��������.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lvw��Դ����"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "��ҩ��(&5)"
      TabPicture(4)   =   "Frm���ŷ�ҩ��������.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "��ӡ����(&6)"
      TabPicture(5)   =   "Frm���ŷ�ҩ��������.frx":1D86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmd��ӡ����"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cboƱ������"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblƱ��"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame fra��ҩ���ݼ�� 
         Caption         =   " ��ҩ���ݼ��"
         Height          =   1335
         Left            =   240
         TabIndex        =   56
         Top             =   3360
         Width           =   5895
         Begin VB.CheckBox chk����������� 
            Caption         =   "��������������������"
            Height          =   180
            Left            =   240
            TabIndex        =   58
            Top             =   720
            Width           =   3825
         End
         Begin VB.CheckBox chk���洢�ⷿ 
            Caption         =   "��鴢��ⷿ"
            Height          =   180
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   1785
         End
      End
      Begin VB.CheckBox chkDetailPage 
         Caption         =   "������һ�δ���ر�ʱ��ҳǩ"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   2880
         Width           =   2745
      End
      Begin VB.CheckBox Chk�����һ�����ʾ 
         Caption         =   "�����һ�����ʾ"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Width           =   1785
      End
      Begin VB.CommandButton cmd��ӡ���� 
         Caption         =   "��ӡ����(&P)"
         Height          =   345
         Left            =   -74760
         TabIndex        =   52
         Top             =   1050
         Width           =   3315
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   300
         Left            =   -74010
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   600
         Width           =   2565
      End
      Begin VB.Frame Frame4 
         Caption         =   " ���Ϳ���  "
         Height          =   615
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   4935
         Begin VB.CheckBox chkStopTrans 
            Caption         =   "��ͣ��ҩƷ��װ�����ͷ�ҩ����"
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �����������ݿ���  "
         Height          =   3855
         Left            =   -74880
         TabIndex        =   43
         Top             =   1080
         Width           =   4935
         Begin VB.Frame Frame6 
            Caption         =   " ��������  "
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   4695
            Begin VB.CheckBox chkType 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   48
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox chkType 
               Caption         =   "����"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   47
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " ����ѡ��"
            Height          =   2775
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   4695
            Begin MSComctlLib.ListView LvwҩƷ���� 
               Height          =   2385
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   4207
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
               Icons           =   "imgLvwSel"
               SmallIcons      =   "imgLvwSel"
               ColHdrIcons     =   "imgLvwSel"
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
         End
      End
      Begin VB.ComboBox cboName 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Frame frm��ΣҩƷ���� 
         Caption         =   " ѡ���ΣҩƷ�������ŵ����"
         Height          =   580
         Left            =   -74880
         TabIndex        =   36
         Top             =   1845
         Width           =   6135
         Begin VB.CheckBox chk��Σ 
            Caption         =   "C��"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   39
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk��Σ 
            Caption         =   "B��"
            Height          =   375
            Index           =   1
            Left            =   1140
            TabIndex        =   38
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk��Σ 
            Caption         =   "A��"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.ComboBox cbo��ҩ�嵥 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   945
         Width           =   2655
      End
      Begin VB.ComboBox cbo��ҩ�嵥 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   555
         Width           =   2655
      End
      Begin VB.Frame fra�豸���� 
         Caption         =   "  ���ܿ��������豸���� "
         Height          =   1095
         Left            =   -71280
         TabIndex        =   28
         Top             =   510
         Width           =   2415
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "�豸����(&S)"
            Height          =   350
            Left            =   480
            TabIndex        =   29
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ��ѯ��ϸ��¼����������ʱ����"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   24
         Top             =   2040
         Width           =   5295
         Begin VB.TextBox txtMaxRecordCount 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            TabIndex        =   25
            Text            =   "3000"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   2160
            TabIndex        =   27
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ϸ��¼"
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " ���ò�ѯ��ҩ����ҩ����ʱ��ʱ�䷶Χ������ʱ����"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtTimeArea_Sended 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "3"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox txtTimeArea_Send 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "7"
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   1920
            TabIndex        =   23
            Top             =   900
            Width           =   180
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ҩ����"
            Height          =   180
            Left            =   240
            TabIndex        =   22
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   1920
            TabIndex        =   20
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ҩ����"
            Height          =   180
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ѡ���ڷ�ҩʱ�Զ����Ϊ�������ҩƷ����"
         Height          =   2340
         Left            =   -74880
         TabIndex        =   12
         Top             =   2640
         Width           =   6135
         Begin MSComctlLib.ListView lvw��ֵ���� 
            Height          =   1755
            Left            =   2160
            TabIndex        =   16
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw������� 
            Height          =   1750
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw��Σ���� 
            Height          =   1755
            Left            =   4200
            TabIndex        =   34
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "��ΣҩƷ�ȼ�����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   4200
            TabIndex        =   35
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ��ֵ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   2160
            TabIndex        =   15
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ�������"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.ComboBox Cbo����ģʽ 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox Cbo��ҩҩ�� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   690
         Width           =   1815
      End
      Begin VB.ComboBox Cbo������ 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2025
         Width           =   1815
      End
      Begin MSComctlLib.ListView Lvw��Դ���� 
         Height          =   4605
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   8123
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
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblƱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74730
         TabIndex        =   53
         Top             =   660
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "ҩ����ʾ"
         Height          =   180
         Left            =   -74880
         TabIndex        =   42
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl��ҩ�嵥 
         Caption         =   "��ҩ�嵥"
         Height          =   195
         Left            =   -74880
         TabIndex        =   33
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label lbl��ҩ�嵥 
         Caption         =   "��ҩ�嵥"
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Lbl����ģʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ҩ��"
         ForeColor       =   &H00000080&
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Lbl��ҩҩ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩҩ��"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ҫ�������Ǵ����������ʱ�������߼���"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1200
         Width           =   4710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   2085
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm���ŷ�ҩ��������.frx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm���ŷ�ҩ��������.frx":20BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm���ŷ�ҩ��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPrivs As String
Private mblnSetPara As Boolean                          '�Ƿ���в�������Ȩ��
Private BlnStart As Boolean
Private intDays As Integer
Private lngҩ��ID As Long
Private Lng����ģʽ As Long
Private Lng������ʾ As Long
Private Lng�Զ���ӡ As Long
Private Lngȱҩ��� As Long
Private Lng��ҩ��ǩ�� As Long
Private Lng��ҩ��ǩ�� As Long
Private str������� As String
Private str��ֵ���� As String
Private RecDrugStore As New ADODB.Recordset             'ҩ��
Private mstrSourceDep As String                         '��Դ����
Private mLng��ӡ��ҩ�嵥 As Long                        '��ҩ�嵥
Public blnStartPacker As Boolean                       '�Ƿ�����ҩƷ�ְ����ӿ�
Private Sub Get������()
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    '���ü�����
    On Error GoTo errHandle
    strsql = "Select Distinct A.����" & _
             " From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
             " Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id And D.��Ա���� = 'ҩ����ҩ��' " & _
             " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
        
    If Cbo��ҩҩ��.ListIndex <> -1 Then
        strsql = strsql & " AND B.����id=[1] "
    End If
    
    strsql = strsql & " ORDER BY A.���� "

    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption, Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex))
    
    Cbo������.Clear
    Cbo������.AddItem "���м�����"
    Do While Not rs.EOF
        Cbo������.AddItem rs!����
        rs.MoveNext
    Loop
    
    rs.Close
    
    Cbo������.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Cbo��ҩҩ��_Click()
    Call Get������
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, 100, 1342)
    
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim n As Integer
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim str��Σ���� As String
    Dim str��Σ���� As String
    
    str������� = ""
    For n = 1 To lvw�������.ListItems.count
        If lvw�������.ListItems(n).Checked = True Then
            str������� = IIf(str������� = "", lvw�������.ListItems(n).Text, str������� & "," & lvw�������.ListItems(n).Text)
        End If
    Next
    
    str��ֵ���� = ""
    For n = 1 To lvw��ֵ����.ListItems.count
        If lvw��ֵ����.ListItems(n).Checked = True Then
            str��ֵ���� = IIf(str��ֵ���� = "", lvw��ֵ����.ListItems(n).Text, str��ֵ���� & "," & lvw��ֵ����.ListItems(n).Text)
        End If
    Next
    
    For n = 1 To lvw��Σ����.ListItems.count
        If lvw��Σ����.ListItems(n).Checked = True Then
            str��Σ���� = IIf(str��Σ���� = "", n, str��Σ���� & "," & n)
        End If
    Next
    
    If chk��Σ(0).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 1, str��Σ���� & "," & 1)
    If chk��Σ(1).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 2, str��Σ���� & "," & 2)
    If chk��Σ(2).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 3, str��Σ���� & "," & 3)
    
    '����˽�в���
    zldatabase.SetPara "�����һ�����ʾ�����嵥", Chk�����һ�����ʾ.Value, glngSys, 1342
    zldatabase.SetPara "����ģʽ", Cbo����ģʽ.ListIndex, glngSys, 1342
    zldatabase.SetPara "������", Cbo������.Text, glngSys, 1342
    zldatabase.SetPara "�������", str�������, glngSys, 1342
    zldatabase.SetPara "��ֵ����", str��ֵ����, glngSys, 1342
    zldatabase.SetPara "��Σ����", str��Σ����, glngSys, 1342
    zldatabase.SetPara "��ΣҩƷ����", str��Σ����, glngSys, 1342
    zldatabase.SetPara "��ҩҩ��", Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex), glngSys, 1342
    zldatabase.SetPara "�Զ���ӡ", Me.cbo��ҩ�嵥.ListIndex, glngSys, 1342
    zldatabase.SetPara "��ѯ��ҩ����", Val(txtTimeArea_Send.Text), glngSys, 1342
    zldatabase.SetPara "��ѯ��ҩ����", Val(txtTimeArea_Sended.Text), glngSys, 1342
    zldatabase.SetPara "��ѯ��ϸ��¼��", Val(txtMaxRecordCount.Text), glngSys, 1342
    zldatabase.SetPara "��ӡ��ҩ�嵥", Me.cbo��ҩ�嵥.ListIndex, glngSys, 1342
    zldatabase.SetPara "��ҩʱ���洢�ⷿ", Me.chk���洢�ⷿ.Value, glngSys, 1342
    zldatabase.SetPara "��ҩʱ���������������", Me.chk�����������.Value, glngSys, 1342
    
    '��Դ����
    mstrSourceDep = ""
    With Me.Lvw��Դ����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With
    zldatabase.SetPara "��Դ����", mstrSourceDep, glngSys, 1342
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", Me.cboName.ListIndex)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ŷ�ҩ����", "������һ�δ���ر�ʱ��ҳǩ", Me.chkDetailPage.Value)
    
    '�����װ������
    If blnStartPacker = True Then
        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��ͣ����", chkStopTrans.Value
        
        str���� = ""
        str���� = str���� & chkType(0).Value
        str���� = str���� & chkType(1).Value

        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��������", str����
        
        
        If LvwҩƷ����.ListItems(1).Checked Then
             str���� = "����"
        Else
            For n = 1 To LvwҩƷ����.ListItems.count
                If LvwҩƷ����.ListItems(n).Checked Then
                    str���� = IIf(str���� = "", "", str���� & ",") & LvwҩƷ����.ListItems(n).Text
                End If
            Next
        End If
        
        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "ѡ�����", str����
    End If
    
'    Frm���ŷ�ҩ����.BlnSetPara = True
    frm���ŷ�ҩ����New.BlnSetPara = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnStart = False Then
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    Dim intTrans As Integer
    Dim str���� As String
    Dim str���� As String
    Dim n As Integer
    
    BlnStart = False
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(strPrivs, "����ҩ��") Then
        strsql = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��' And ������� IN (2,3))"
    Else
        strsql = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��' And B.������� IN (2,3))"
    End If
    gstrSQL = " Select ID,����||'-'||���� ҩ�� From ���ű� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID In " & strsql & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order by ����||'-'||����"
    Set RecDrugStore = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecDrugStore
        If .EOF Then
            MsgBox "���ʼ��ҩ���������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Cbo��ҩҩ��.Clear
        Do While Not .EOF
            Cbo��ҩҩ��.AddItem !ҩ��
            Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.NewIndex) = !Id
            .MoveNext
        Loop
        Cbo��ҩҩ��.ListIndex = 0
    End With
    
    With Cbo����ģʽ
        .Clear
        .AddItem "0-�������е���"
        .AddItem "1-���������ʵ�"
        .AddItem "2-���������ʱ�"
        .ListIndex = 0
    End With
        
    With cbo��ҩ�嵥
        .Clear
        .AddItem "0_��ҩ�󲻴�ӡ"
        .AddItem "1-��ҩ���Զ���ӡ"
        .AddItem "2_��ҩ����ʾ�Ƿ��ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ҩ�嵥
        .Clear
        .AddItem "0_��ҩ�󲻴�ӡ"
        .AddItem "1-��ҩ���Զ���ӡ"
        .AddItem "2_��ҩ����ʾ�Ƿ��ӡ"
        .ListIndex = 0
    End With
    
    With Me.cboName
        .Clear
        .AddItem "0-��ʾҩƷ����������"
        .AddItem "1-����ʾҩƷ����"
        .AddItem "2-����ʾҩƷ����"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-���ܷ�ҩ�嵥"
        .AddItem "2-��ҩ�嵥"
        .ListIndex = 0
    End With
    
    Call Get������
    
    '�������
    gstrSQL = "Select ���� From ҩƷ������� Order By ���� "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-ȡ�������")
    
    With rsTmp
        Do While Not .EOF
            lvw�������.ListItems.Add , "_" & lvw�������.ListItems.count + 1, !����
            .MoveNext
        Loop
    End With
    
    '��ֵ����
    gstrSQL = "Select ���� From ҩƷ��ֵ���� Order By ���� "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-ȡ��ֵ����")
    
    With rsTmp
        Do While Not .EOF
            lvw��ֵ����.ListItems.Add , "_" & lvw��ֵ����.ListItems.count + 1, !����
            .MoveNext
        Loop
    End With
    
    '��ΣҩƷ����
    With lvw��Σ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "A��"
        .ListItems.Add , "_" & .ListItems.count + 1, "B��"
        .ListItems.Add , "_" & .ListItems.count + 1, "C��"
    End With
    
    '�ָ�����
    WriteCons

    '��Դ����
    Call SetSourceDep
    
    '��װ���ӿ��������
    Call LoadҩƷ����(Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex))
    
    tabShow.TabVisible(4) = blnStartPacker
    
    If blnStartPacker = True Then
        intTrans = Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��ͣ����", "0"))
        chkStopTrans.Value = IIf(intTrans = 1, 1, 0)
        
        str���� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��������", "11")
        chkType(0).Value = Val(Mid(str����, 1, 1))
        chkType(1).Value = Val(Mid(str����, 2, 1))
        
        str���� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "ѡ�����", "����")
        
        For n = 1 To LvwҩƷ����.ListItems.count
            LvwҩƷ����.ListItems(n).Checked = False
            If str���� = "����" Then
                LvwҩƷ����.ListItems(n).Checked = True
            Else
                If InStr(1, "," & str���� & ",", "," & LvwҩƷ����.ListItems(n).Text & ",") > 0 Then
                    LvwҩƷ����.ListItems(n).Checked = True
                End If
            End If
        Next
    End If
    
    BlnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ����(ByVal lngҩ��ID As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get����(lngҩ��ID)
    
    With LvwҩƷ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "����ҩƷ����", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, Mid(rsData!����, InStr(1, rsData!����, "-") + 1), 1, 1
            .ListItems(.ListItems.count).Checked = True
            rsData.MoveNext
        Loop
    End With
End Sub
Private Function WriteCons()
    Dim IntLocate As Integer
    Dim str������ As String
    Dim n As Integer
    Dim i As Integer
    Dim int�Զ�ˢ�� As Integer
    Dim strArr
    Dim int��ѯ��ҩ���� As Integer
    Dim int��ѯ��ҩ���� As Integer
    Dim lng����¼�� As Long
    Dim int��˳�Ժ�������� As Integer
    Dim str��Σ���� As String
    Dim str��Σ���� As String
    Dim int�������� As Integer
    Dim int���洢�ⷿ As Integer
    Dim int����������� As Integer
    
    mblnSetPara = zlStr.IsHavePrivs(strPrivs, "��������")
    
    'ȡ������˽�в���
    Lng����ģʽ = Val(zldatabase.GetPara("����ģʽ", glngSys, 1342, 0, Array(Cbo����ģʽ), mblnSetPara))
    Lng������ʾ = Val(zldatabase.GetPara("�����һ�����ʾ�����嵥", glngSys, 1342, 0, Array(Chk�����һ�����ʾ), mblnSetPara))
    str������ = zldatabase.GetPara("������", glngSys, 1342, "���м�����", Array(Label1, Cbo������), mblnSetPara)
    str������� = zldatabase.GetPara("�������", glngSys, 1342, "", Array(Label2, lvw�������), mblnSetPara)
    str��ֵ���� = zldatabase.GetPara("��ֵ����", glngSys, 1342, "", Array(Label3, lvw��ֵ����), mblnSetPara)
    str��Σ���� = zldatabase.GetPara("��Σ����", glngSys, 1342, "", Array(Label11, lvw��Σ����), mblnSetPara)
    str��Σ���� = zldatabase.GetPara("��ΣҩƷ����", glngSys, 1342, "", Array(frm��ΣҩƷ����), mblnSetPara)
    lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1342, 0, Array(Lbl��ҩҩ��, Cbo��ҩҩ��), mblnSetPara))
    Lng�Զ���ӡ = Val(zldatabase.GetPara("�Զ���ӡ", glngSys, 1342, 0, Array(Me.lbl��ҩ�嵥, Me.cbo��ҩ�嵥), mblnSetPara))
    int��ѯ��ҩ���� = Val(zldatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 7, Array(txtTimeArea_Send), mblnSetPara))
    int��ѯ��ҩ���� = Val(zldatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 3, Array(txtTimeArea_Sended), mblnSetPara))
    lng����¼�� = Val(zldatabase.GetPara("��ѯ��ϸ��¼��", glngSys, 1342, 3000, Array(txtMaxRecordCount), mblnSetPara))
    int�������� = Val(zldatabase.GetPara("��ҩʱ������ҩ���ʼ�¼", glngSys, 1342, 0))
    int���洢�ⷿ = Val(zldatabase.GetPara("��ҩʱ���洢�ⷿ", glngSys, 1342, 0))
    int����������� = Val(zldatabase.GetPara("��ҩʱ���������������", glngSys, 1342, 0))
    
    mstrSourceDep = zldatabase.GetPara("��Դ����", glngSys, 1342, "", Array(Lvw��Դ����), mblnSetPara)
    mLng��ӡ��ҩ�嵥 = Val(zldatabase.GetPara("��ӡ��ҩ�嵥", glngSys, 1342, 0, Array(lbl��ҩ�嵥, Me.cbo��ҩ�嵥), mblnSetPara))
    
    '���ݲ���ֵ����
    If lngҩ��ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.Cbo��ҩҩ��.ListCount - 1
            If Me.Cbo��ҩҩ��.ItemData(IntLocate) = lngҩ��ID Then
                Me.Cbo��ҩҩ��.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cbo��ҩҩ��.ListCount - 1) Then
            MsgBox "����������ҩ����ԭ�����õ�ҩ����ʧЧ����", vbInformation, gstrSysName
            If Cbo��ҩҩ��.ListCount >= 1 Then Cbo��ҩҩ��.ListIndex = 0
        End If
    End If
    Me.Cbo����ģʽ.ListIndex = Lng����ģʽ
    Me.cbo��ҩ�嵥.ListIndex = Lng�Զ���ӡ
    Me.Chk�����һ�����ʾ.Value = Lng������ʾ
    Me.cbo��ҩ�嵥.ListIndex = mLng��ӡ��ҩ�嵥
    Me.cboName.ListIndex = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", 0))
    Me.chkDetailPage.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "������һ�δ���ر�ʱ��ҳǩ", 0))
    Me.chk���洢�ⷿ.Value = int���洢�ⷿ
    Me.chk�����������.Value = int�����������
    
    If int�������� = 1 Then
        Me.Chk�����һ�����ʾ.Value = 1
        Me.Chk�����һ�����ʾ.Enabled = False
    End If

    For n = 0 To Cbo������.ListCount - 1
        If Cbo������.List(n) = str������ Then
            Cbo������.ListIndex = n
            Exit For
        End If
    Next
    
    If str������� <> "" Then
        For n = 1 To lvw�������.ListItems.count
            If InStr("," & str������� & ",", "," & lvw�������.ListItems(n).Text & ",") > 0 Then
                lvw�������.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��ֵ���� <> "" Then
        For n = 1 To lvw��ֵ����.ListItems.count
            If InStr("," & str��ֵ���� & ",", "," & lvw��ֵ����.ListItems(n).Text & ",") > 0 Then
                lvw��ֵ����.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��Σ���� <> "" Then
        For n = 1 To lvw��Σ����.ListItems.count
            If InStr("," & str��Σ���� & ",", "," & n & ",") > 0 Then
                lvw��Σ����.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��Σ���� <> "" Then
        If InStr(1, str��Σ����, "1") Then chk��Σ(0).Value = 1
        If InStr(1, str��Σ����, "2") Then chk��Σ(1).Value = 1
        If InStr(1, str��Σ����, "3") Then chk��Σ(2).Value = 1
    End If
    
    If int��ѯ��ҩ���� <= 0 Or int��ѯ��ҩ���� > 99 Then
        int��ѯ��ҩ���� = 7
    End If
    txtTimeArea_Send.Text = int��ѯ��ҩ����
        
    If int��ѯ��ҩ���� <= 0 Or int��ѯ��ҩ���� > 99 Then
        int��ѯ��ҩ���� = 3
    End If
    txtTimeArea_Sended.Text = int��ѯ��ҩ����
    
    If lng����¼�� <= 0 Then
        lng����¼�� = 3000
    End If
    txtMaxRecordCount.Text = lng����¼��
    
End Function

Private Sub LvwҩƷ����_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With LvwҩƷ����
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "����ҩƷ����" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub


Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case tabShow.Tab
    Case 0
        If Cbo��ҩҩ��.Enabled = True Then Cbo��ҩҩ��.SetFocus
    End Select
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    
    Select Case cboƱ������.ListIndex
    Case 0
        '���ܷ�ҩ��
        strBill = "ZL1_BILL_1342"
    Case 1
        '��ҩ�嵥
        strBill = "ZL1_BILL_1342_1"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
Private Sub txtMaxRecordCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMaxRecordCount_Validate(Cancel As Boolean)
    If Val(txtMaxRecordCount.Text) <= 0 Then
        txtMaxRecordCount.Text = 3000
    End If
End Sub


Private Sub txtTimeArea_Send_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Send_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Send.Text) <= 0 Then
        txtTimeArea_Send.Text = 7
    End If
End Sub


Private Sub txtTimeArea_Sended_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Sended_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Sended.Text) <= 0 Then
        txtTimeArea_Sended.Text = 3
    End If
End Sub


Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select distinct A.���� || '-' || A.���� ����, A.Id " & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.Id =B.����id and B.�������� in ('���','����','����','����','Ӫ��', '�ٴ�','����') And B.������� In (2,3)  And " & _
            " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.���� || '-' || A.����"

    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "SetSourceDep")
    Call SQLTest

    With rs
        If .EOF Then
            MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw��Դ����.ListItems.Clear
        Do While Not .EOF
            Lvw��Դ����.ListItems.Add , "_" & !Id, !����, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw��Դ����.ListItems("_" & !Id).Checked = True
                End If
            End If
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



