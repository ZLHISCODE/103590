VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���µ�ѡ��"
   ClientHeight    =   6840
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8865
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame FraSplit 
      Height          =   135
      Left            =   30
      TabIndex        =   60
      Top             =   6075
      Width           =   8775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6300
      TabIndex        =   61
      Top             =   6345
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7500
      TabIndex        =   62
      Top             =   6345
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   5910
      Index           =   0
      Left            =   120
      TabIndex        =   63
      Top             =   135
      Width           =   8625
      Begin VB.CheckBox chk 
         Caption         =   "���µ���ӡʱ�������������(�������ʵ���ʹ����Ч)"
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   35
         Top             =   5475
         Width           =   4770
      End
      Begin VB.PictureBox picPoly 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6885
         ScaleHeight     =   240
         ScaleWidth      =   1545
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4875
         Width           =   1545
         Begin VB.OptionButton optPloy 
            Caption         =   "ֱ��"
            Height          =   180
            Index           =   1
            Left            =   765
            TabIndex        =   59
            Top             =   15
            Width           =   720
         End
         Begin VB.OptionButton optPloy 
            Caption         =   "б��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   58
            Top             =   15
            Value           =   -1  'True
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "���µ��ļ��Ŀ�ʼʱ��"
         Height          =   795
         Left            =   5355
         TabIndex        =   54
         Top             =   3960
         Width           =   3135
         Begin VB.OptionButton opt���µ���ʼʱ�� 
            Caption         =   "��Ժʱ��"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   55
            Top             =   330
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt���µ���ʼʱ�� 
            Caption         =   "���ʱ��"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   56
            Top             =   330
            Width           =   1125
         End
      End
      Begin VB.Frame fra 
         Caption         =   "�����Զ���־"
         Height          =   3675
         Index           =   15
         Left            =   5355
         TabIndex        =   36
         Top             =   165
         Width           =   3135
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   5
            ItemData        =   "frmCaseTendBodyPara.frx":000C
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2100
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   4
            ItemData        =   "frmCaseTendBodyPara.frx":0010
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1734
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   3
            ItemData        =   "frmCaseTendBodyPara.frx":0014
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1368
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   2
            ItemData        =   "frmCaseTendBodyPara.frx":0018
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":001A
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1002
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   1
            ItemData        =   "frmCaseTendBodyPara.frx":001C
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":001E
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   636
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   0
            ItemData        =   "frmCaseTendBodyPara.frx":0020
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   270
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   6
            ItemData        =   "frmCaseTendBodyPara.frx":0024
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   2466
            Width           =   2430
         End
         Begin VB.ComboBox cboBody 
            Height          =   300
            Index           =   7
            ItemData        =   "frmCaseTendBodyPara.frx":0028
            Left            =   525
            List            =   "frmCaseTendBodyPara.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   2835
            Width           =   2430
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����Զ���־��˳���ڵ�������"
            Height          =   180
            Index           =   12
            Left            =   135
            TabIndex        =   53
            Top             =   3255
            Width           =   2895
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ"
            Height          =   180
            Index           =   50
            Left            =   135
            TabIndex        =   47
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
            TabIndex        =   45
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
            TabIndex        =   43
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
            TabIndex        =   41
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
            TabIndex        =   39
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
            TabIndex        =   37
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
            TabIndex        =   49
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
            TabIndex        =   51
            Top             =   2895
            Width           =   360
         End
      End
      Begin VB.ComboBox cboNote 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":002C
         Left            =   1695
         List            =   "frmCaseTendBodyPara.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2475
         Width           =   1905
      End
      Begin VB.ComboBox cboStyle 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0030
         Left            =   1695
         List            =   "frmCaseTendBodyPara.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2835
         Width           =   1905
      End
      Begin VB.PictureBox PicValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1740
         Left            =   2100
         ScaleHeight     =   1740
         ScaleWidth      =   2175
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdUnVisible 
            Height          =   315
            Left            =   1665
            Picture         =   "frmCaseTendBodyPara.frx":0034
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "ȡ��"
            Top             =   1395
            Width           =   450
         End
         Begin zl9TemperatureChartGX.ColorPicker usrColor 
            Height          =   2190
            Left            =   -15
            TabIndex        =   66
            Top             =   -450
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.ComboBox cboOper 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":05BE
         Left            =   1710
         List            =   "frmCaseTendBodyPara.frx":05C0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1890
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   0
         Left            =   2130
         ScaleHeight     =   165
         ScaleWidth      =   270
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "���ѡ����ɫ"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chk 
         Caption         =   "�೦�����Է��ӷ�ĸ��ʾ"
         Height          =   180
         Index           =   15
         Left            =   195
         TabIndex        =   34
         Top             =   5205
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ӥ�����µ���ʾ��Ժ��Ϣ"
         Height          =   180
         Index           =   5
         Left            =   195
         TabIndex        =   32
         Top             =   4635
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ�Ƿ��ӡҽԺ����"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   33
         Top             =   4920
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����������ݴ�ӡ���ʱ������ʾ�������ݼ̳�)"
         Height          =   180
         Index           =   4
         Left            =   195
         TabIndex        =   30
         Top             =   4110
         Width           =   4400
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ܡ�������Ŀ��ʾ�������ݣ�������ʾ���죩"
         Height          =   180
         Index           =   6
         Left            =   195
         TabIndex        =   31
         Top             =   4380
         Width           =   4215
      End
      Begin VB.ComboBox cboSplit 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":05C2
         Left            =   2400
         List            =   "frmCaseTendBodyPara.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2145
         Width           =   900
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ����ʾƤ�Խ��"
         Height          =   315
         Index           =   8
         Left            =   195
         TabIndex        =   29
         Top             =   3765
         Width           =   2790
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ��Ե�����ʾ(����������˫����ʾ)"
         Height          =   315
         Index           =   7
         Left            =   195
         TabIndex        =   28
         Top             =   3465
         Width           =   3630
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ�����ʾ���˵������Ϣ"
         Height          =   315
         Index           =   3
         Left            =   195
         TabIndex        =   27
         Top             =   3180
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
         TabIndex        =   19
         Text            =   "1"
         Top             =   1830
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   1185
         Width           =   420
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   3
         Text            =   "14"
         Top             =   585
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������ע�������ٴ�����ʱ,ֹͣǰһ��������ע"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   885
         Width           =   4500
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   6
         Left            =   2085
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1185
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1500
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1830
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
         Left            =   4635
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   8
         BuddyControl    =   "txt(4)"
         BuddyDispid     =   196627
         BuddyIndex      =   4
         OrigLeft        =   4635
         OrigTop         =   1170
         OrigRight       =   4890
         OrigBottom      =   1440
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   18
         BuddyControl    =   "txt(3)"
         BuddyDispid     =   196627
         BuddyIndex      =   3
         OrigLeft        =   3285
         OrigTop         =   1170
         OrigRight       =   3540
         OrigBottom      =   1440
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   2955
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "18"
         Top             =   1185
         Width           =   330
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   4305
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "6"
         Top             =   1185
         Width           =   330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������䷽ʽ��"
         Height          =   180
         Index           =   5
         Left            =   5340
         TabIndex        =   57
         Top             =   4875
         Width           =   1620
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��˵����ʾλ��"
         Height          =   180
         Index           =   11
         Left            =   210
         TabIndex        =   23
         Top             =   2535
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���²�����ʾ��ʽ"
         Height          =   180
         Index           =   13
         Left            =   210
         TabIndex        =   25
         Top             =   2895
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����       ��"
         Height          =   180
         Index           =   8
         Left            =   3930
         TabIndex        =   12
         Top             =   1230
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҹ���       ����"
         Height          =   180
         Index           =   7
         Left            =   2400
         TabIndex        =   9
         Top             =   1230
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ȱʡ��ʽ"
         Height          =   180
         Index           =   10
         Left            =   210
         TabIndex        =   0
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Զ���־��ʱ��֮����           ����"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   21
         Top             =   2205
         Width           =   3510
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����¼�볬����ǰ        ��Ļ����¼����"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   1875
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���±����ʱ��������ݹ̶�        ��"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   15
         Top             =   1545
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���¿�ʼ��¼ʱ��"
         Height          =   180
         Index           =   31
         Left            =   210
         TabIndex        =   6
         Top             =   1230
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   1410
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�������ע    ��,��ɫ"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   585
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
    Dim intIndex As Integer
    
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
    
    cboOper.Clear
    cboOper.AddItem "0-����ʾ"
    cboOper.AddItem "1-��ʾ0"
    cboOper.AddItem "2-��ʾ��������"
    
    cboNote.Clear
    cboNote.AddItem "0-��ʾ������"
    cboNote.AddItem "1-��ʾ������"
    cboNote.AddItem "2-����ʾ"
    
    cboStyle.Clear
    cboStyle.AddItem "0-��ͷ"
    cboStyle.AddItem "1-����"
    cboStyle.AddItem "2-����+��ͷ"
    cboStyle.AddItem "3-����+����"
    
    cboSplit.Clear
    cboSplit.AddItem "����"
    cboSplit.AddItem "��"
    
    intIndex = zlDatabase.GetPara("���µ��ļ���ʼʱ��", glngSys, 1255, 1, Array(opt���µ���ʼʱ��(0), opt���µ���ʼʱ��(1)), InStr(mstrPrivs, "����ѡ������") > 0)
    If intIndex > 0 And intIndex < 2 Then
        opt���µ���ʼʱ��(intIndex).Value = True
    Else
        opt���µ���ʼʱ��(0).Value = True
    End If
    
    txt(6).Text = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4, Array(txt(6), ud(6), lbl(31)), InStr(mstrPrivs, "����ѡ������") > 0)
    txt(1).Text = zlDatabase.GetPara("���±������", glngSys, 1255, 8, Array(txt(1), ud(1), lbl(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = zlDatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1:1", Array(cboBody(0), cboBody(1), cboBody(2), cboBody(3), cboBody(4), cboBody(5), cboBody(6), cboBody(7)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    For intLoop = 0 To 7
        If UBound(Split(strTmp, ";")) >= intLoop Then
            cboBody(intLoop).ListIndex = Val(Split(strTmp, ";")(intLoop))
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    
    strTmp = zlDatabase.GetPara("���±�־�ָ���", glngSys, 1255, "0", Array(cboSplit, lbl(6)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
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
    
    strTmp = zlDatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2", Array(cboOper, lbl(10)), InStr(mstrPrivs, "����ѡ������") > 0)
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboOper.ListIndex = Val(strTmp)
    Else
        cboOper.ListIndex = 0
    End If
    
    'δ��˵����ʾλ��
    strTmp = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0", Array(cboNote, lbl(11)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboNote.ListIndex = CInt(Val(strTmp))
    Else
        cboNote.ListIndex = 0
    End If
    
     '���²�����ʾ��ʽ
    strTmp = Val(zlDatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0", Array(cboStyle, lbl(13)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        cboStyle.ListIndex = CInt(Val(strTmp))
    Else
        cboStyle.ListIndex = 0
    End If
    
    picColor(0).BackColor = Val(zlDatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "0", Array(picColor(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    
    txt(0).Text = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10", Array(txt(0), lbl(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("��ӡҽԺ����", glngSys, 1255, "1", Array(chk(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Value = Val(zlDatabase.GetPara("���µ���ʾ���", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    txt(2).Text = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1", Array(txt(2), lbl(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(4).Value = Val(zlDatabase.GetPara("����������", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(5).Value = Val(zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, "1", Array(chk(5)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(6).Value = Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, "1", Array(chk(6)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(7).Value = Val(zlDatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, "0", Array(chk(7)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(8).Value = Val(zlDatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0", Array(chk(8))))
    chk(12).Value = Val(zlDatabase.GetPara("���±�־��˳��������", glngSys, 1255, "0", Array(chk(12)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(14).Value = Val(zlDatabase.GetPara("���µ�����ӡ������", glngSys, 1255, "0", Array(chk(14)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(15).Value = Val(zlDatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, "0", Array(chk(15)), InStr(mstrPrivs, "����ѡ������") > 0))
    
    strTmp = Val(zlDatabase.GetPara("���������䷽ʽ", glngSys, 1255, "0", Array(lbl(5), optPloy(0), optPloy(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) = 1 Then
        optPloy(1).Value = True
    Else
        optPloy(0).Value = True
    End If
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function


Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboOper_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboSplit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
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
    Dim strTmp As String
    Dim intIndex As Integer
    Dim lngColor As Long
        
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex & ";" & cboBody(2).ListIndex & ";" & cboBody(3).ListIndex & ";" & cboBody(4).ListIndex & ";" & cboBody(5).ListIndex & ";" & cboBody(6).ListIndex & ";" & cboBody(7).ListIndex
    Call zlDatabase.SetPara("���¿�ʼʱ��", Val(txt(6).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����ʱ��ҹ���־", txt(3).Text & ";" & txt(4).Text, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���±������", Val(txt(1).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ����", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�������ע����", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��������ȱʡ��ʽ", Val(cboOper.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    lngColor = picColor(0).BackColor
    Call zlDatabase.SetPara("����������ʾ��ɫ", lngColor, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    If opt���µ���ʼʱ��(0).Value Then
        intIndex = 0
    Else
        intIndex = 1
    End If
    Call zlDatabase.SetPara("���µ��ļ���ʼʱ��", intIndex, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����¼�뻤����������", Val(txt(2).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�ٴ�����ֹͣǰ�α�ע", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("δ��˵����ʾλ��", Val(cboNote.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾ���", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��ӡҽԺ����", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����������", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", chk(5).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���ܲ�����ʾ��������", chk(6).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾ��ʽ", chk(7).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ���ʾƤ�Խ��", chk(8).Value, glngSys, 1255)
    Call zlDatabase.SetPara("���±�־�ָ���", Val(cboSplit.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���±�־��˳��������", chk(12).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�೦������ʾ��ʽ", chk(15).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���²�����ʾ��ʽ", Val(cboStyle.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("���µ�����ӡ������", chk(14).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    If optPloy(1).Value Then
        intIndex = 1
    Else
        intIndex = 0
    End If
    Call zlDatabase.SetPara("���������䷽ʽ", intIndex, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    mblnOK = True
    
    Unload Me
End Sub

Private Sub optPloy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub opt���µ���ʼʱ��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub picColor_Click(Index As Integer)
    On Error Resume Next
    
    With picValue
        If Index = 0 Or Index = 2 Then
            .Left = picColor(Index).Left + fra(0).Left
            .Top = picColor(Index).Top + picColor(Index).Height + fra(0).Top
        Else
            .Left = fra(15).Left + 200
            .Top = picColor(Index).Top + picColor(Index).Height
        End If
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    
    usrColor.Tag = Index
    usrColor.Color = picColor(Index).BackColor
End Sub

Private Sub picColor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub PicValue_GotFocus()
    If usrColor.Visible Then usrColor.SetFocus
End Sub

Private Sub usrColor_LostFocus()
   If Not Me.ActiveControl Is usrColor _
        And Not Me.ActiveControl Is picValue _
        And Not Me.ActiveControl Is cmdUnVisible _
    Then picValue.Visible = False
End Sub

Private Sub usrColor_pOK()
    picColor(Val(usrColor.Tag)).BackColor = usrColor.Color
    picValue.Visible = False
    If picColor(Val(usrColor.Tag)).Enabled And picColor(Val(usrColor.Tag)).Visible Then picColor(Val(usrColor.Tag)).SetFocus
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

