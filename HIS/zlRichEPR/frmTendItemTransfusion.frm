VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendItemTransfusion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ƶ���ʿվ��������"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmTendItemTransfusion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   150
      Top             =   420
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5475
      Left            =   90
      TabIndex        =   47
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "��������"
      TabPicture(0)   =   "frmTendItemTransfusion.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblҺ������"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboҺ������"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboҺ����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lvw����"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "���±�����"
      TabPicture(1)   =   "frmTendItemTransfusion.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstColumnUsed"
      Tab(1).Control(1)=   "cmdColumn(1)"
      Tab(1).Control(2)=   "cmdColumn(0)"
      Tab(1).Control(3)=   "lstColumnItems"
      Tab(1).Control(4)=   "cmdMove(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdMove(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblColumnItems(1)"
      Tab(1).Control(7)=   "lblColumnItems(0)"
      Tab(1).Control(8)=   "Label1(2)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "��Ŀ����"
      TabPicture(2)   =   "frmTendItemTransfusion.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd�޸�"
      Tab(2).Control(1)=   "cmdɾ��"
      Tab(2).Control(2)=   "cmd����"
      Tab(2).Control(3)=   "txt������"
      Tab(2).Control(4)=   "lstClass"
      Tab(2).Control(5)=   "lstItems"
      Tab(2).Control(6)=   "Label3"
      Tab(2).Control(7)=   "Label1(3)"
      Tab(2).Control(8)=   "lblColumnItems(3)"
      Tab(2).Control(9)=   "lblColumnItems(2)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "����������"
      TabPicture(3)   =   "frmTendItemTransfusion.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic����"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdģ��"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "picDraw"
      Tab(3).Control(4)=   "SSTab1"
      Tab(3).ControlCount=   5
      Begin MSComctlLib.ListView lvw���� 
         Height          =   2085
         Left            =   450
         TabIndex        =   54
         Top             =   3240
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3678
         View            =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   240
         TabIndex        =   53
         Top             =   2610
         Width           =   6525
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -74550
         ScaleHeight     =   315
         ScaleWidth      =   4035
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   390
         Width           =   4035
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   465
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   3555
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
            Left            =   30
            TabIndex        =   26
            Top             =   60
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdģ�� 
         Caption         =   "��ģ�������ǰ��������"
         Height          =   315
         Left            =   -70380
         TabIndex        =   28
         Top             =   390
         Width           =   2325
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ҫ������"
         Height          =   4635
         Left            =   -70380
         TabIndex        =   29
         Top             =   780
         Width           =   2325
         Begin VB.CommandButton cmd������_���� 
            Caption         =   "����"
            Height          =   350
            Left            =   180
            TabIndex        =   43
            Top             =   4170
            Width           =   945
         End
         Begin VB.CommandButton cmd������_ɾ�� 
            Caption         =   "ɾ��"
            Height          =   350
            Left            =   1230
            TabIndex        =   44
            Top             =   4170
            Width           =   945
         End
         Begin VB.TextBox txt��������Ŀ 
            Appearance      =   0  'Flat
            Height          =   1275
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   2820
            Width           =   2025
         End
         Begin VB.CommandButton cmd��������Ŀ 
            Caption         =   "��������Ŀ"
            Height          =   345
            Left            =   180
            TabIndex        =   41
            Top             =   2460
            Width           =   2025
         End
         Begin VB.CheckBox chk������ʱ���� 
            Caption         =   "������ʱ���ظ���"
            Height          =   225
            Left            =   180
            TabIndex        =   40
            Top             =   2160
            Width           =   1905
         End
         Begin VB.ComboBox cboλ�� 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1770
            Width           =   1575
         End
         Begin VB.TextBox txt�к� 
            Height          =   300
            Left            =   600
            TabIndex        =   37
            Top             =   1410
            Width           =   1575
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   600
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   690
            Width           =   1575
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   600
            TabIndex        =   35
            Top             =   1050
            Width           =   1575
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label lblλ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "λ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   38
            Top             =   1830
            Width           =   360
         End
         Begin VB.Label lbl�к� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�к�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   36
            Top             =   1470
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
            Left            =   195
            TabIndex        =   34
            Top             =   1110
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
            Left            =   195
            TabIndex        =   32
            Top             =   750
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
            Left            =   195
            TabIndex        =   30
            Top             =   390
            Width           =   360
         End
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   -74610
         Picture         =   "frmTendItemTransfusion.frx":007C
         ScaleHeight     =   4665
         ScaleWidth      =   4065
         TabIndex        =   25
         Top             =   720
         Width           =   4095
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   7
            Left            =   -30
            Top             =   120
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   6
            Left            =   -30
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   5
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   4
            Left            =   720
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   3
            Left            =   720
            Top             =   120
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   2
            Left            =   720
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   1
            Left            =   330
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   0
            Left            =   -30
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label lblҪ������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ҫ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   840
            TabIndex        =   50
            Top             =   45
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblҪ���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ҫ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5100
         Left            =   -74970
         TabIndex        =   24
         Top             =   330
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8996
         _Version        =   393216
         TabOrientation  =   2
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "�����ſ�"
         TabPicture(0)   =   "frmTendItemTransfusion.frx":46526
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "��������"
         TabPicture(1)   =   "frmTendItemTransfusion.frx":46542
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "ע������"
         TabPicture(2)   =   "frmTendItemTransfusion.frx":4655E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin VB.CommandButton cmd�޸� 
         Caption         =   "�޸�"
         Height          =   315
         Left            =   -72930
         TabIndex        =   48
         Top             =   4410
         Width           =   585
      End
      Begin VB.CommandButton cmdɾ�� 
         Caption         =   "ɾ��"
         Height          =   350
         Left            =   -73410
         TabIndex        =   21
         Top             =   4800
         Width           =   1065
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "����"
         Height          =   350
         Left            =   -74460
         TabIndex        =   20
         Top             =   4800
         Width           =   1065
      End
      Begin VB.TextBox txt������ 
         Height          =   315
         Left            =   -74040
         TabIndex        =   19
         Top             =   4410
         Width           =   1095
      End
      Begin VB.ListBox lstClass 
         Height          =   3120
         Left            =   -74430
         TabIndex        =   17
         Top             =   1215
         Width           =   2100
      End
      Begin VB.ListBox lstItems 
         Height          =   4020
         Left            =   -71850
         MultiSelect     =   2  'Extended
         TabIndex        =   23
         Top             =   1200
         Width           =   2100
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   4020
         Left            =   -71070
         TabIndex        =   10
         Top             =   1200
         Width           =   2100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -72180
         TabIndex        =   12
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   -72180
         TabIndex        =   11
         Top             =   2145
         Width           =   975
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   4020
         Left            =   -74430
         TabIndex        =   9
         Top             =   1215
         Width           =   2100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&U)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -72180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3270
         Width           =   975
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -72180
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3570
         Width           =   975
      End
      Begin VB.ComboBox cboҺ���� 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2010
         Width           =   2535
      End
      Begin VB.ComboBox cboҺ������ 
         Height          =   300
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "    ҽԺ��ʹ�����°滤ʿ����վ���빴ѡʹ�����°滤ʿվ�Ĳ���"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   52
         Top             =   3000
         Width           =   6135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   -74415
         TabIndex        =   18
         Top             =   4470
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "    ����ҽԺ�ڲ�������ʱ��ʹ�õ����±�������Ŀ���á�"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   15
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "������Ŀ����"
         Height          =   180
         Index           =   3
         Left            =   -74400
         TabIndex        =   16
         Top             =   990
         Width           =   2040
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "������Ŀ�б�"
         Height          =   180
         Index           =   2
         Left            =   -71820
         TabIndex        =   22
         Top             =   990
         Width           =   2070
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "���±���Ŀ"
         Height          =   180
         Index           =   1
         Left            =   -71040
         TabIndex        =   8
         Top             =   990
         Width           =   2070
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ѡ�����¼��Ŀ"
         Height          =   180
         Index           =   0
         Left            =   -74400
         TabIndex        =   7
         Top             =   990
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "    ����ҽԺ�ڲ�������ʱ��ʹ�õ����±�������Ŀ���á�"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   6
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "    �ƶ���ʿ����վ��ִ����Һ��ҽ����ʱ�����ҽ���´��˼ǳ�����ҽ����������Զ����������ݸ����ݡ�"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label lblҺ������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1)������Һ�����ƹ����Ļ�����Ŀ"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   630
         TabIndex        =   2
         Top             =   1740
         Width           =   2700
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2)������Һ���������Ļ�����Ŀ"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   3720
         TabIndex        =   4
         Top             =   1740
         Width           =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "    ֻ����Һ������ʱ������һ�����ܺ�������ݸ����ݣ����ͬʱ������Һ�����ƣ��������ϸ�������ݸ����ݡ�"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   330
         TabIndex        =   1
         Top             =   1080
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5430
      TabIndex        =   46
      Top             =   5700
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   45
      Top             =   5700
      Width           =   1100
   End
End
Attribute VB_Name = "frmTendItemTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnALLOW As Boolean        '�Ƿ��б༭��Ȩ��
Private mblnStart As Boolean
Private mblnClick As Boolean
Private mblnEdit As Boolean         '�Ƿ�����˱༭
Private Enum emnuPage
    �����ſ�
    ��������
End Enum
Private mrsBoard As New ADODB.Recordset

Private Sub cbo����_Click()
    Frame1.Enabled = (cbo����.ListCount > 0) And InStr(1, mstrPrivs, "�༭") <> 0
    Timer1.Enabled = True
End Sub

Private Sub cbo����_Change()
    If mblnClick Then Exit Sub
    cmd������_����.Caption = "����"
    Call SetShape
    Frame1.Tag = ""
End Sub

Private Sub cboҺ����_Click()
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cboҺ������_Click()
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnClear As Boolean, blnTrans As Boolean
    Dim intRow As Integer, intCount As Integer
    Dim sqlText As String
    
    On Error GoTo errHand
    
    If cboҺ������.ItemData(cboҺ������.ListIndex) > 0 And cboҺ����.ItemData(cboҺ����.ListIndex) = 0 Then
        MsgBox "������Һ������Ӧ�Ļ����¼��Ŀ��", vbInformation, gstrSysName
        cboҺ����.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    '����������Ŀ
    blnClear = True
    If cboҺ������.ItemData(cboҺ������.ListIndex) > 0 Then
        gstrSQL = "ZL_�����¼��Ŀ_Transfusion(" & cboҺ������.ItemData(cboҺ������.ListIndex) & ",'11'," & IIf(blnClear, "1", "0") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Һ������")
        blnClear = False
    End If
    If cboҺ����.ItemData(cboҺ����.ListIndex) > 0 Then
        gstrSQL = "ZL_�����¼��Ŀ_Transfusion(" & cboҺ����.ItemData(cboҺ����.ListIndex) & ",'12'," & IIf(blnClear, "1", "0") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Һ����")
        blnClear = False
    End If
    
    '�������±�����
    intCount = lstColumnUsed.ListCount
    For intRow = 1 To intCount
        sqlText = sqlText & "," & Me.lstColumnUsed.ItemData(intRow - 1)
    Next
    sqlText = Mid(sqlText, 2)
    Call zlDatabase.SetPara("���±���Ŀ", sqlText, 100)
    
    '�����°��Ե㲡��
    sqlText = ""
    intCount = lvw����.ListItems.Count
    For intRow = 1 To intCount
        If lvw����.ListItems(intRow).Checked Then
            sqlText = sqlText & "," & Mid(lvw����.ListItems(intRow).Key, 2)
        End If
    Next
    If sqlText <> "" Then
        sqlText = Mid(sqlText, 2)
        Call zlDatabase.SetPara("�ƶ���ʿվ�°没���б�", sqlText, 100)
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    mblnEdit = False
    
    Unload Me
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd��ӷ���_Click()
    
End Sub

Private Sub cmd��������Ŀ_Click()
    Dim strIDs As String, strNames As String
    Call frmClinicSelect.ShowMe(Me, strIDs, strNames)
    txt��������Ŀ.Tag = strIDs
    txt��������Ŀ.Text = strNames
End Sub

Private Sub cmd������_����_Click()
    Dim lngID As Long
    Dim intPos As Integer
    Dim intCount As Integer
    Dim blnTrans As Boolean
    Dim strIDs As String, strItems As String
    
    If Trim(txt����.Text) = "" Then
        MsgBox "��������Ϊ�գ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    If Trim(txt�к�.Text) = "" Then
        MsgBox "�кŲ���Ϊ�գ�", vbInformation, gstrSysName
        txt�к�.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txt�к�.Text) Then
        MsgBox "�к��в��ܺ��зǷ��ַ���", vbInformation, gstrSysName
        txt�к�.SetFocus
        Exit Sub
    End If
    If Val(txt�к�.Text) < 0 Or Val(txt�к�.Text) > 13 Then
        MsgBox "�кŲ���С��������13��", vbInformation, gstrSysName
        txt�к�.SetFocus
        Exit Sub
    End If
    
    Me.Caption = "�ƶ���ʿվ��������" & "(���ڱ�������,���Ժ�......)"
    gcnOracle.BeginTrans
    blnTrans = True
    
    If txt��������Ŀ.Tag <> "" Then
        strIDs = txt��������Ŀ.Tag
    End If
    lngID = Val(Frame1.Tag)
    If lngID = 0 Then lngID = zlDatabase.GetNextId("��������ʽ")
    gstrSQL = "ZL_��������ʽ_APPENDITEM(" & lngID & "," & Me.cbo����.ItemData(Me.cbo����.ListIndex) & "," & Me.cbo����.ListIndex + 1 & "," & _
        "'" & Me.cbo����.Text & "','" & Me.txt����.Text & "'," & Me.txt�к�.Text & "," & Me.cboλ��.ListIndex + 1 & "," & _
        IIf(Me.cbo����.ListIndex = -1, 0, 1) & "," & chk������ʱ����.Value & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '�����ʽ:<ITEMLIST><ITEM><XH/><MC/></ITEM></ITEMLIST>
    intCount = 0
    Do While strIDs <> ""
        If Len(strIDs) > 3800 Then
            '������Ѱ����
            intPos = GetSplit(Mid(strIDs, 1, 3800))
            strItems = Mid(strIDs, 1, intPos)
            strIDs = Mid(strIDs, intPos + 1)
        Else
            strItems = strIDs
            strIDs = ""
        End If
        
        gstrSQL = "ZL_��������ʽ_UPDATEZLXM(" & lngID & ",'" & strItems & "'," & IIf(intCount = 0, "1", "0") & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        intCount = intCount + 1
    Loop
    Me.Caption = "�ƶ���ʿվ��������"
    gcnOracle.CommitTrans
    blnTrans = False
    
    Timer1.Enabled = True
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Me.Caption = "�ƶ���ʿվ��������"
End Sub

Private Function GetSplit(ByVal strInput As String) As Integer
    Dim intPos As Integer
    '������Ѱ����,���ض��ŵ�λ��
    
    intPos = 3800
    Do While True
        If Mid(strInput, intPos, 1) = "," Then
            intPos = intPos - 1
            GetSplit = intPos
            Exit Function
        End If
        intPos = intPos - 1
    Loop
End Function

Private Sub cmd������_ɾ��_Click()
    On Error GoTo errHand
    
    If Val(Frame1.Tag) = 0 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ��Ҫ�أ�" & cbo����.Text & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gcnOracle.Execute "ZL_��������ʽ_DELETEITEM(" & Val(Frame1.Tag) & ")", , adCmdStoredProc
    
    Timer1.Enabled = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdģ��_Click()
    If MsgBox("��ɾ����ǰ�����������ݺ����ݹ�����ģ���������ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlDatabase.ExecuteProcedure("ZL_��������ʽ_BUILD(" & Me.cbo����.ItemData(Me.cbo����.ListIndex) & ")", "ģ�������ǰ��������")
    Timer1.Enabled = True
End Sub

Private Sub cmdɾ��_Click()
    Dim intIndex As Integer
    Dim strXH As String
    On Error GoTo errHand
    
    strXH = GetSelItems()
    If strXH = "" Then
        MsgBox "����Ҫѡ��һ����Ŀ��", vbInformation, gstrSysName
        lstItems.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_�����¼��Ŀ_MOBILE('','" & strXH & "')", "ɾ����������")
    
    intIndex = lstClass.ListIndex
    If intIndex = -1 Then Exit Sub
    lstClass.RemoveItem intIndex
    If intIndex < lstClass.ListCount Then
        lstClass.ListIndex = intIndex
    Else
        If lstClass.ListCount >= 1 Then
            lstClass.ListIndex = intIndex - 1
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd�޸�_Click()
    Dim strXH As String
    On Error GoTo errHand
    
    If Trim(txt������.Text) = "" Then
        MsgBox "����������Ϊ�գ�", vbInformation, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    strXH = GetSelItems
    Call zlDatabase.ExecuteProcedure("ZL_�����¼��Ŀ_MOBILE('" & txt������.Text & "','" & strXH & "')", "���·�����")
    lstClass.List(lstClass.ListIndex) = txt������.Text
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd����_Click()
    Dim strXH As String
    On Error GoTo errHand
    
    If Trim(txt������.Text) = "" Then
        MsgBox "����������Ϊ�գ�", vbInformation, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    strXH = GetSelItems(True)
    If strXH = "" Then
        MsgBox "����Ҫѡ��һ����Ŀ��", vbInformation, gstrSysName
        lstItems.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_�����¼��Ŀ_MOBILE('" & txt������.Text & "','" & strXH & "')", "���·�������")
    lstClass.AddItem txt������.Text
    lstClass.ListIndex = 0
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    Dim str���±� As String, str�Ե� As String
    Dim rsTemp As New ADODB.Recordset
    
    mblnClick = False
    mblnStart = False
    mblnEdit = False
    mstrPrivs = gstrPrivs
    mblnALLOW = (InStr(1, gstrPrivs, "�༭") > 0)
    
    '1)������Ŀ
    '��ȡ�ı�����Ŀ
    gstrSQL = "Select ��Ŀ����,��Ŀ���,��Ŀ���� From �����¼��Ŀ Where (��Ŀ����=1 And ��Ŀ����>10) or ��Ŀ��ʾ=4 Order by ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ")
    rsTemp.Filter = "��Ŀ����=1"
    cboҺ������.AddItem "δ����"
    Call zlControl.CboAddData(cboҺ������, rsTemp, False)
    rsTemp.Filter = "��Ŀ����<>1"
    cboҺ����.AddItem "δ����"
    Call zlControl.CboAddData(cboҺ����, rsTemp, False)
    rsTemp.Filter = 0
    '��ȡ���õ�����
    gstrSQL = " Select ��Ŀ���,�������� From �����¼��Ŀ Where �������� IN ('11','12')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ")
    
    rsTemp.Filter = "��������='11'"
    If rsTemp.RecordCount <> 0 Then
        Call zlControl.CboLocate(cboҺ������, rsTemp!��Ŀ���, True)
    Else
        cboҺ������.ListIndex = 0
    End If
    rsTemp.Filter = "��������='12'"
    If rsTemp.RecordCount <> 0 Then
        Call zlControl.CboLocate(cboҺ����, rsTemp!��Ŀ���, True)
    Else
        cboҺ����.ListIndex = 0
    End If
    rsTemp.Filter = 0
    
    '��ȡ�°��Ե㲡��
    str�Ե� = "," & zlDatabase.GetPara("�ƶ���ʿվ�°没���б�", 100) & ","
    gstrSQL = "" & _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����б�")
    With lvw����
        .ListItems.Clear
        Do While Not rsTemp.EOF
            .ListItems.Add , "K" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����
            If InStr(1, str�Ե�, "," & rsTemp!ID & ",") <> 0 Then .ListItems("K" & rsTemp!ID).Checked = True
            rsTemp.MoveNext
        Loop
    End With
    
    '2)���±�
    str���±� = zlDatabase.GetPara("���±���Ŀ", 100)
    gstrSQL = " Select A.��Ŀ���,A.��Ŀ���� From �����¼��Ŀ A" & _
              " Where A.Ӧ�÷�ʽ<>0 AND ��Ŀ��� NOT IN (SELECT * FROM TABLE(F_NUM2LIST([1])))  " & _
              " Order by ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���±�)
    With rsTemp
        Me.lstColumnItems.Clear
        Do While Not .EOF
            Me.lstColumnItems.AddItem !��Ŀ����
            Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With
    '��ȡ��ѡ�����Ŀ�嵥
    gstrSQL = " Select B.��Ŀ���,B.��Ŀ���� From �����¼��Ŀ B " & _
              " Where B.��Ŀ��� IN (SELECT * FROM TABLE(F_NUM2LIST([1])))" & _
              " Order by B.��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���±�)
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.lstColumnUsed.AddItem !��Ŀ����
            Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    
    '3)������Ŀ����
    gstrSQL = " Select distinct �ƶ����� From �����¼��Ŀ B where �ƶ����� Is Not NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstClass.Clear
        Me.lstClass.AddItem "δ�趨"
        Me.lstClass.AddItem "ȫ��"
        Do While Not .EOF
            Me.lstClass.AddItem !�ƶ�����
            .MoveNext
        Loop
        If lstClass.ListCount > 0 Then Me.lstClass.ListIndex = 0
    End With
    
    '4)������
    Call InitUnits
    With cbo����
        .Clear
        .AddItem "����ԭ������"
        .ItemData(.NewIndex) = 1
        .AddItem "������������"
        .ItemData(.NewIndex) = 2
        .AddItem "һ�������б�"
        .ItemData(.NewIndex) = 3
        .AddItem "�ؼ������б�"
        .ItemData(.NewIndex) = 4
        .AddItem "��Σ�б�"
        .ItemData(.NewIndex) = 5
        .AddItem "��Ժ�б�"
        .ItemData(.NewIndex) = 6
        .AddItem "��Ժ�б�"
        .ItemData(.NewIndex) = 7
        .AddItem "Ԥ��Ժ�б�"
        .ItemData(.NewIndex) = 8
        .AddItem "�����б�"
        .ItemData(.NewIndex) = 9
        .AddItem "Ԥ�����б�"
        .ItemData(.NewIndex) = 10
        .AddItem "ת���б�"
        .ItemData(.NewIndex) = 11
    End With
    
    cbo����.Clear
    cbo����.AddItem "�����ſ�"
    cbo����.AddItem "��������"
    cbo����.AddItem "ע������"
    cbo����.ListIndex = 0
    
    cboλ��.Clear
    cboλ��.AddItem "��"
    cboλ��.AddItem "��"
    cboλ��.ListIndex = 0
    
    Call SetEnabled
    mblnStart = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim intIndex As Integer
    Dim objlst As ListBox
    If Index = 0 Then
        If Me.lstColumnItems.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnItems.ListIndex
        Me.lstColumnUsed.AddItem Me.lstColumnItems.Text
        Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = Me.lstColumnItems.ItemData(Me.lstColumnItems.ListIndex)
        Me.lstColumnItems.RemoveItem Me.lstColumnItems.ListIndex
        Set objlst = lstColumnItems
    Else
        If Me.lstColumnUsed.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnUsed.ListIndex
        Me.lstColumnItems.AddItem Me.lstColumnUsed.Text
        Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = Me.lstColumnUsed.ItemData(Me.lstColumnUsed.ListIndex)
        Me.lstColumnUsed.RemoveItem Me.lstColumnUsed.ListIndex
        Set objlst = lstColumnUsed
    End If
    If objlst.ListCount >= intIndex + 1 Then
        objlst.ListIndex = intIndex
    Else
        objlst.ListIndex = objlst.ListCount - 1
    End If
    
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0) And mblnALLOW
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0) And mblnALLOW
    
    Call SetMoveState
    
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Dim arrData
    Dim strCopy As String
    Dim lngDo As Long, lngMAX As Long
    Dim lngSelIndex As Long, lngTarIndex As Long
    
    '��ǰ����
    lngSelIndex = lstColumnUsed.ListIndex
    'Ŀ������
    lngTarIndex = lngSelIndex + IIf(Index = 0, -1, 1)
    lngMAX = lstColumnUsed.ListCount - 1
    For lngDo = 0 To lngMAX
        If lngDo = lngTarIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngSelIndex) & "," & lstColumnUsed.ItemData(lngSelIndex)
        ElseIf lngDo = lngSelIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngTarIndex) & "," & lstColumnUsed.ItemData(lngTarIndex)
        Else
            strCopy = strCopy & "|" & lstColumnUsed.List(lngDo) & "," & lstColumnUsed.ItemData(lngDo)
        End If
    Next
    strCopy = Mid(strCopy, 2)
    Debug.Print strCopy
    
    lstColumnUsed.Clear
    arrData = Split(strCopy, "|")
    For lngDo = 0 To lngMAX
        lstColumnUsed.AddItem Split(arrData(lngDo), ",")(0)
        lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Split(arrData(lngDo), ",")(1))
    Next
    lstColumnUsed.ListIndex = lngTarIndex
    
    Call SetMoveState
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub lblҪ����_Click(Index As Integer)
    Dim intDo As Integer, intCount As Integer
    
    mblnClick = True
    intCount = lblҪ����.Count - 1
    For intDo = 1 To intCount
        lblҪ����(intDo).BackStyle = 0
    Next
    Frame1.Tag = lblҪ����(Index).Tag
    cmd������_����.Caption = "�޸�"
    lblҪ����(Index).BackStyle = 1
    Call SetShape(Index)
    
    '��λ��Ҫ�أ���ʾ��Ӧ������
    mrsBoard.Filter = "ID=" & Val(lblҪ����(Index).Tag)
    If mrsBoard.RecordCount = 0 Then Exit Sub
    
    cbo����.ListIndex = mrsBoard!���� - 1
    If Not zlControl.CboLocate(cbo����, mrsBoard!����) Then cbo����.Text = mrsBoard!����
    txt����.Text = IIf(IsNull(mrsBoard!����), "", mrsBoard!����)
    txt�к�.Text = mrsBoard!�к�
    cboλ��.ListIndex = mrsBoard!λ�� - 1
    chk������ʱ����.Value = mrsBoard!�Ƿ�����
    
    txt��������Ŀ.Text = Get������ĿNAME(Val(lblҪ����(Index).Tag))
    txt��������Ŀ.Tag = Get������ĿID(Val(lblҪ����(Index).Tag))
    mblnClick = False
End Sub

Private Function Get������ĿID(ByVal lngID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "" & _
        " SELECT a.XH" & _
        " FROM ��������ʽ p," & _
        " XMLTable('/ITEMLIST/ITEM/XH' PASSING p.������Ŀ" & _
        " COLUMNS XH VARCHAR2(256) PATH '/XH') a" & _
        " Where p.id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ����", lngID)
    With rsTemp
        Do While Not rsTemp.EOF
            Get������ĿID = Get������ĿID & "," & rsTemp!xh
            rsTemp.MoveNext
        Loop
    End With
    
    If Get������ĿID <> "" Then Get������ĿID = Mid(Get������ĿID, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get������ĿNAME(ByVal lngID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "" & _
        " SELECT a.MC" & _
        " FROM ��������ʽ p," & _
        " XMLTable('/ITEMLIST/ITEM/MC' PASSING p.������Ŀ" & _
        " COLUMNS MC VARCHAR2(256) PATH '/MC') a" & _
        " Where p.id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ����", lngID)
    With rsTemp
        Do While Not rsTemp.EOF
            Get������ĿNAME = Get������ĿNAME & "," & rsTemp!MC
            rsTemp.MoveNext
        Loop
    End With
    
    If Get������ĿNAME <> "" Then Get������ĿNAME = Mid(Get������ĿNAME, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lstClass_Click()
    Dim strCond As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    cmd�޸�.Enabled = False
    cmdɾ��.Enabled = False
    txt������.Text = lstClass.Text
    gstrSQL = " Select ��Ŀ���,��Ŀ���� From �����¼��Ŀ"
    If lstClass.Text = "ȫ��" Then
        
    ElseIf lstClass.Text = "δ�趨" Then
        strCond = " Where �ƶ����� Is NULL"
    Else
        cmd�޸�.Enabled = True And mblnALLOW
        cmdɾ��.Enabled = True And mblnALLOW
        strCond = " Where �ƶ����� =[1]"
    End If
    gstrSQL = gstrSQL & strCond & " Order by ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", lstClass.Text)
    
    With rsTemp
        lstItems.Clear
        Do While Not .EOF
            lstItems.AddItem !��Ŀ����
            lstItems.ItemData(lstItems.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub lstColumnItems_DblClick()
    If lstColumnItems.ListCount = 0 Then Exit Sub
        If Not mblnALLOW Then Exit Sub
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    Call SetMoveState
End Sub

Private Sub lstColumnUsed_DblClick()
    If lstColumnUsed.ListCount = 0 Then Exit Sub
        If Not mblnALLOW Then Exit Sub
    Call cmdColumn_Click(1)
End Sub

Private Sub SetMoveState()
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    
    If lstColumnUsed.ListIndex < 0 Then Exit Sub
    If lstColumnUsed.SelCount < 0 Then Exit Sub
    cmdMove(0).Enabled = (lstColumnUsed.ListIndex > 0) And mblnALLOW
    cmdMove(1).Enabled = (lstColumnUsed.ListIndex < lstColumnUsed.ListCount - 1) And mblnALLOW
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("ȷ��Ҫ�˳������������޸Ļ�δ���棡", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Function GetSelItems(Optional ByVal blnSel As Boolean = False) As String
    Dim i As Integer, j As Integer
    
    j = lstItems.ListCount
    For i = 1 To j
        If blnSel Then
            If lstItems.Selected(i - 1) Then
                GetSelItems = GetSelItems & "," & lstItems.ItemData(i - 1)
            End If
        Else
            GetSelItems = GetSelItems & "," & lstItems.ItemData(i - 1)
        End If
    Next
    If GetSelItems <> "" Then GetSelItems = Mid(GetSelItems, 2)
End Function

Private Sub SetEnabled()
    Me.cboҺ����.Enabled = mblnALLOW
    Me.cboҺ������.Enabled = mblnALLOW
    lvw����.Enabled = mblnALLOW
    cmdColumn(0).Enabled = mblnALLOW
    cmdColumn(1).Enabled = mblnALLOW
    cmdMove(0).Enabled = mblnALLOW
    cmdMove(1).Enabled = mblnALLOW
    cmd����.Enabled = mblnALLOW
    cmd�޸�.Enabled = mblnALLOW
    cmdɾ��.Enabled = mblnALLOW
    Frame1.Enabled = (cbo����.ListCount > 0) And InStr(1, mstrPrivs, "�༭") <> 0
    cmd������_����.Enabled = InStr(1, mstrPrivs, "�༭") <> 0
    cmd������_ɾ��.Enabled = cmd������_����.Enabled
End Sub

Private Function GetUser����IDs() As String
'���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSQL = _
            "Select Distinct ����ID From (" & _
            " Select A.����ID as ����ID" & _
            " From ��������˵�� A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
            " And A.������� in(1,2,3) And A.��������='����'" & _
            " Union" & _
            " Select A.����ID From �������Ҷ�Ӧ A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngUserId)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strUnits As String, i As Long

    On Error GoTo errH
    strUnits = GetUser����IDs
    
    '�����Ź۲���
    If InStr(mstrPrivs, "���в���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If

    cbo����.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserId)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "���в���") > 0 Then
                If rsTmp!ID = glngDeptId Then  'ֱ����������
                    Call zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
                End If
                If InStr("," & strUnits & ",", "," & rsTmp!ID & ",") > 0 And cbo����.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
                End If
            Else '����ȱʡ���������Ŀ����ж��
                If rsTmp!ȱʡ = 1 And cbo����.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then
        Call zlControl.CboSetIndex(cbo����.hwnd, 0)
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefreshBoard()
    Dim lng����ID As Long
    Dim intDel As Integer, intCount As Integer
    On Error GoTo errHand
    'ˢ�¹�����
    
    '��ɾ�����пؼ�
    intCount = lblҪ����.Count - 1
    For intDel = 1 To intCount
        Unload lblҪ����(intDel)
        Unload lblҪ������(intDel)
    Next
    '����ֵ
    cbo����.Text = ""
    txt����.Text = ""
    txt�к�.Text = ""
    cboλ��.ListIndex = 0
    chk������ʱ����.Value = 0
    txt��������Ŀ.Text = ""
    txt��������Ŀ.Tag = ""
    
    
    '��ȡ����
    Frame1.Tag = ""
    lng����ID = Me.cbo����.ItemData(Me.cbo����.ListIndex)
    gstrSQL = " Select ID,����,����,����,�к�,λ��,�Ƿ�̶�,�Ƿ�����,����" & _
              " From ��������ʽ " & _
              " Where ����ID=[1] " & _
              " Order by �к�,λ��"
    Set mrsBoard = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", lng����ID)
    
    '���μ��ؿؼ�
    With mrsBoard
        .Filter = "����=" & SSTab1.Tab + 1
        Do While Not .EOF
            Load lblҪ����(.AbsolutePosition)
            lblҪ����(.AbsolutePosition).Tag = !ID
            lblҪ����(.AbsolutePosition).Caption = !����
            lblҪ����(.AbsolutePosition).Top = lblҪ����(0).Top + (!�к� - 1) * 360
            lblҪ����(.AbsolutePosition).Left = IIf(!λ�� = 1, 60, 2580)
            lblҪ����(.AbsolutePosition).Visible = True
            
            Load lblҪ������(.AbsolutePosition)
            lblҪ������(.AbsolutePosition).Caption = IIf(IsNull(!����), "", !����)
            lblҪ������(.AbsolutePosition).Top = lblҪ������(0).Top + (!�к� - 1) * 360
            lblҪ������(.AbsolutePosition).Left = lblҪ����(.AbsolutePosition).Left + lblҪ����(.AbsolutePosition).Width + 60
            lblҪ������(.AbsolutePosition).AutoSize = False
            lblҪ������(.AbsolutePosition).WordWrap = False
            lblҪ������(.AbsolutePosition).Height = 240
            lblҪ������(.AbsolutePosition).Visible = True
            
            .MoveNext
        Loop
        .Filter = 0
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call RefreshBoard
    Call SetShape
End Sub

Private Sub SetShape(Optional ByVal intIndex As Integer = 0)
    Dim blnShow As Boolean
    blnShow = (intIndex > 0)
    
    If blnShow Then
        Shape(0).Left = lblҪ����(intIndex).Left - Shape(0).Width
        Shape(0).Top = lblҪ����(intIndex).Top - Shape(0).Height
        Shape(1).Left = lblҪ����(intIndex).Left + (lblҪ����(intIndex).Width - Shape(0).Width) / 2
        Shape(1).Top = Shape(0).Top
        Shape(2).Left = lblҪ����(intIndex).Left + lblҪ����(intIndex).Width
        Shape(2).Top = Shape(0).Top
        Shape(3).Left = Shape(2).Left
        Shape(3).Top = lblҪ����(intIndex).Top + (lblҪ����(intIndex).Height - Shape(3).Height) / 2
        Shape(4).Left = Shape(2).Left
        Shape(4).Top = lblҪ����(intIndex).Top + lblҪ����(intIndex).Height
        Shape(5).Left = Shape(1).Left
        Shape(5).Top = Shape(4).Top
        Shape(6).Left = Shape(0).Left
        Shape(6).Top = Shape(4).Top
        Shape(7).Left = Shape(0).Left
        Shape(7).Top = Shape(3).Top
    End If
    
    Shape(0).Visible = blnShow
    Shape(1).Visible = blnShow
    Shape(2).Visible = blnShow
    Shape(3).Visible = blnShow
    Shape(4).Visible = blnShow
    Shape(5).Visible = blnShow
    Shape(6).Visible = blnShow
    Shape(7).Visible = blnShow
End Sub




