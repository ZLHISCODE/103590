VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmTechnicExpense 
   AutoRedraw      =   -1  'True
   Caption         =   "���˼ƷѴ���"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTechnicExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   7275
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmTechnicExpense.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   0
      ScaleHeight     =   2160
      ScaleWidth      =   11880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5115
      Width           =   11880
      Begin MSComctlLib.ImageList imgList 
         Left            =   7335
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicExpense.frx":1A92
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   9780
         TabIndex        =   20
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1125
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   7965
         TabIndex        =   19
         ToolTipText     =   "�ȼ���F2"
         Top             =   1125
         Width           =   1680
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   35
         ToolTipText     =   "���:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�������"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   11
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6555
            TabIndex        =   15
            Top             =   180
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   16
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            HideSelection   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBaby 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ӥ����(&B)"
            Height          =   240
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   240
            Left            =   5790
            TabIndex        =   37
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   8820
            TabIndex        =   36
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   38
         Top             =   390
         Width           =   3675
         Begin VB.TextBox txtʵ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1020
            Width           =   1845
         End
         Begin VB.TextBox txtӦ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   405
            Width           =   1845
         End
         Begin VB.Label lblʵ�� 
            AutoSize        =   -1  'True
            Caption         =   "ʵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   270
            TabIndex        =   40
            Top             =   1095
            Width           =   690
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   270
            TabIndex        =   39
            Top             =   480
            Width           =   690
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   525
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         RowHeightMin    =   320
         BackColorBkg    =   15466495
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "^         ��Ŀ|^          ���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   30
      TabIndex        =   23
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   11865
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1425
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   18000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���˼Ʒѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   225
         TabIndex        =   27
         ToolTipText     =   "���:F6"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9540
         TabIndex        =   24
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   8520
      TabIndex        =   22
      Top             =   855
      Width           =   3375
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   405
         Width           =   2175
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   465
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   21
      Top             =   855
      Width           =   8490
      Begin VB.TextBox txt�ѱ� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   1590
      End
      Begin VB.TextBox txt���ʽ 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   1590
      End
      Begin VB.TextBox txt�Ա� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   870
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   870
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   6510
         TabIndex        =   44
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   4740
         TabIndex        =   43
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl���ʽ 
         Caption         =   "���� ��ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2445
         TabIndex        =   42
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   6750
         TabIndex        =   33
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   2415
         TabIndex        =   30
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   4980
         TabIndex        =   29
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3195
      Left            =   15
      TabIndex        =   10
      Top             =   1920
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5636
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmTechnicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��ڲ���
'����������������������������������������������������������������������������������������������������������������������������������������
Public mlngҽ��ID As Long '��������ʱ��
Public mlng���ͺ� As Long '��������ʱ��
Public mlng����ID As Long 'ȷ��Ҫ�ƷѵĲ���ID
Public mlng��ҳID As Long 'ȷ��Ҫ�Ʒѵ���ҳID

Public mint������Դ As Integer '1-���ﲡ��,2-סԺ����
Public mint��¼���� As Integer '1-�շ�(����),2-����(��/ס)

Public mbln���õǼ� As Boolean '���Ǽ�,����ʵ�ս��
Public mlng��������ID As Long 'Ϊ��ǰ������ҽ������
Public mlng���˿���id As Long '��Ҫ������ȷ�����ﲡ�˵Ŀ���ID

Public mlng��������ID As Long
Public mstr����ҽ�� As String

Public mbytInState As Byte '0-ִ��,1-����,2-����(��֧��),3-ɾ��
Public mstrInNO As String '�������ĵ��ݺ�(ִ��ʱΪ�޸�)

Public mstrTime As String '�����������ݵĵǼ�ʱ��
Public mblnDelete As Boolean '�Ƿ����˷ѵ���(����)

Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum
Private Enum BillCol
    �� = 0
    ��� = 1
    ��Ŀ = 2
    ��� = 3
    ��λ = 4
    ���� = 5
    ���� = 6
    ���� = 7
    Ӧ�ս�� = 8
    ʵ�ս�� = 9
    ִ�п��� = 10
    ��־ = 11
    ���� = 12
End Enum

Public mstrPrivs As String
'ҽ������վ���ط��ò���
'����������������������������������������������������������������������������������������������������������������������������������������
Private mstrLike As String '����ƥ�䷽ʽ
Private mblnPay As Boolean '��ҩ�Ƿ����븶��
Private mblnTime As Boolean '����Ƿ����븶��
Private mbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ�����
Private mbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ����
Private mstr�շ���� As String '��������շ����
Private mblnҩ����λ As Boolean '�Ƿ������ﵥλ��סԺ��λ��ʾҩƷ
Private mstrҩ����λ As String '���ݲ�����Դ������"���ﵥλ"��"סԺ��λ"
Private mstrҩ����װ As String '���ݲ�����Դ������"�����װ"��"סԺ��װ"
Private mlngPreRow As Long '��¼��ǰ��,�����ı���ʱ
Private mlng��ҩ�� As Long, mlng��ҩ�� As Long, mlng��ҩ�� As Long
'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID
'�������
Private mobjBill As ExpenseBill '���õ��ݶ���
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes '������Ŀ���ܼ���

'�������
Private mbytWarn As Byte '���ʱ����ķ���ֵ
Private mintWarn As Integer '���ʱ�����ʾ�ļ���ѡ��
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mrsWarn As New ADODB.Recordset '����������
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ

Private mcurModiMoney As Currency '�޸ĵ���ʱԭ���ݵĽ��
Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnNewRow As Boolean
Private mblnOne As Boolean '�Ƿ�ֻ��һ�������շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mdblItemNum As Double '���ݿ��е�ǰ�����Ŀ������
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private marrDr() As String '��¼ҽ����"ID|����ID|���|����|����"
Private mblnEnterCell As Boolean '�����Ƿ�ִ��Entercell�¼�

Private Const STR_HEAD = "��,450,4;���,750,1;��Ŀ,2175,1;���,1105,1;��λ,520,4;����,520,1;����,570,1;����,1055,7;" & "Ӧ�ս��,1030,7;ʵ�ս��,1080,7;ִ�п���,1255,1;��־,520,4;����,520,1"

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln��������ۿ� As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Then Cancel = True: Exit Sub
    
    If mobjBill.Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("����Ŀ���� " & bytsubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Details(mobjBill.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln��������ۿ� = gbln��������ۿ�
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If

        If bln��������ۿ� Then lngMainRow = mobjBill.Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        '���¼��㲢ˢ��
        If bln��������ۿ� Then
            If ItemHaveSub(lngMainRow) Then
                Call Calc��������ʵ��(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) 'ֻ��һ��������,����ȫ����ɾ��ʱ,������ͨ���������
            End If
        End If
        
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ�������ɾ��
        
        mlngPreRow = 0  '��ʾ�иı���
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, i As Long
    
    'ҩƷ�����
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If mblnҩ����λ Then
                            dblStock = dblStock / .Detail.ҩ����װ
                        End If
                        .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                        sta.Panels(2) = "[" & .Detail.���� & "]���ÿ����:" & dblStock
                        
                        'ҩ���ı�,ʵ��ҩƷ���¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf .�շ���� = "4" And .Detail.�������� Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = dblStock
                        sta.Panels(2) = "[" & .Detail.���� & "]���ÿ����:" & dblStock
                        
                        '���ϲ��Ÿı�,ʱ���������¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                        If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub bill_CellCheck(Row As Long, Col As Long)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte
    
    If Bill.TextMatrix(Row, 2) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mbytInState = 3 Then Exit Sub
    
    '������δ��������Ч
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" And mobjBill.Details(i).���ӱ�־ = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    If bytTime > 0 Then
        mobjBill.Details(Row).���ӱ�־ = IIF(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
    End If
End Sub

Private Function SelectIsNurse() As Boolean
'���ܣ��жϵ�ǰ�������Ƿ�ʿ
    Dim str���� As String
    
    If cbo������.ListIndex <> -1 Then
        If cbo������.ItemData(cbo������.ListIndex) = 0 Then Exit Function
        
        If cbo������.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                str���� = Split(marrDr(cbo������.ListIndex), "|")(6)
                SelectIsNurse = str���� = "��ʿ"
            End If
        End If
    End If
End Function

Private Sub bill_CommandClick()
    Dim lng��ĿID As Long, blnCancel As Boolean
    Dim str��� As String, str��׼��Ŀ As String
    
    If gbln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIF(SelectIsNurse, "'E','M','4'", mstr�շ����)
        End If
    Else
        str��� = IIF(SelectIsNurse, "'E','M','4'", mstr�շ����)
    End If
    If Not IsNull(mrsInfo!����) Then
        str��׼��Ŀ = Get������׼��Ŀ(mrsInfo!����ID, "A.ID")
    End If
    
    lng��ĿID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint������Դ, True, str���, , , str��׼��Ŀ)
    If lng��ĿID <> 0 Then
        Bill.Text = lng��ĿID
        mblnSelect = True
        Call bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency, blnҽ�� As Boolean
    Dim blnStock As Boolean, lngDoUnit As Long, strժҪ As String
    Dim lng��ĿID As Long, str��׼��Ŀ As String, str��� As String
    Dim blnInput As Boolean, cur��� As Currency, lng���˿���ID As Long
    Dim colStock As Collection
    On Error GoTo errH
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        '�Ƿ�ҽ������
        blnҽ�� = Val(txt���ʽ.Tag) = 1 Or Not IsNull(mrsInfo!����)
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "���"
                If Bill.ListIndex <> -1 Then '�������������򲻻���������
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                        For i = 2 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .�շ�ϸĿID = 0: .�շ���� = ""
                            End With
                            Call CalcMoneys
                            Call ShowMoney
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
                End If
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, 2) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set mobjDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If gbln�շ���� Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "û��ȷ���������,�����������"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            str��� = IIF(SelectIsNurse, "'E','M','4'", mstr�շ����)
                        End If
                        If Not IsNull(mrsInfo!����) Then
                            str��׼��Ŀ = Get������׼��Ŀ(mrsInfo!����ID, "A.ID")
                        End If
                        lng��ĿID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint������Դ, True, str���, Bill.Text, Bill.TxtHwnd, str��׼��Ŀ)
                        If lng��ĿID <> 0 Then
                            Set mobjDetail = GetInputDetail(lng��ĿID)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(���Ӳ���)
                    
                    'ҽ��������Ŀ�Ƿ��������
                    If mint������Դ = 2 And mint��¼���� = 2 And Not IsNull(mrsInfo!����) Then
                        If mobjDetail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                            mrsMedAudit.Filter = "��ĿID=" & mobjDetail.ID
                            If mrsMedAudit.RecordCount = 0 Then
                                MsgBox "��ǰ����δ����׼ʹ�ø���Ŀ��", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    
                    '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                    If InStr(",5,6,7,", mobjDetail.���) > 0 _
                        Or (mobjDetail.��� = "4" And mobjDetail.��������) Then
                        If PhysicExist(mobjDetail, Bill.Row) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��鴦��ְ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                        'ҽ���򹫷Ѳ���
                        If InStr(",1,2,", txt���ʽ.Tag) > 0 Then
                            If CheckDuty(mobjDetail, False) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '���в���
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '���˿���ID
                    lng���˿���ID = mobjBill.����ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    sta.Panels(2) = ""
                    lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mobjDetail.���, mobjDetail.ID, _
                        mobjDetail.ִ�п���, lng���˿���ID, Get��������ID, mint������Դ, Nvl(mrsInfo!����ID, 0)) '����ȱʡ�벡�˲���(��������)��ͬ
                    
                    
                    '��ȡҩƷ�����Ϣ
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        '��ǰ��ҩƷ���
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit, blnStock)
                        If mblnҩ����λ Then
                            dblStock = dblStock / mobjDetail.ҩ����װ
                        End If
                        mobjDetail.��� = dblStock
                        sta.Panels(2) = "[" & mobjDetail.���� & "]���ÿ����:" & mobjDetail.���

                        '��������
                        mobjDetail.�������� = Get��������(mobjDetail.ID)
                    ElseIf mobjDetail.��� = "4" And mobjDetail.�������� Then
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                        mobjDetail.��� = dblStock
                        sta.Panels(2).Text = "[" & mobjDetail.���� & "]���ÿ����:" & mobjDetail.���
                    End If
                    
                    '����֧����Ŀ��Ӧ���
                    If Not IsNull(mrsInfo!����) Then
                        If Not ItemExistInsure(mobjDetail.ID, mrsInfo!����) Then
                            If gintҽ������ = 1 Then
                                If MsgBox("��Ŀ""" & mobjDetail.���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            ElseIf gintҽ������ = 2 Then
                                MsgBox "��Ŀ""" & mobjDetail.���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Details(Bill.Row).ժҪ
                        End If
                    End If
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney��ҽ�����ܷ���ժҪ
                    If mobjBill.Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Details(Bill.Row).ժҪ
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mint��¼���� = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        If curTotal > 0 Then
                            cur��� = Val(txtʵ��.Tag)
                            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn, blnҽ��)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    ElseIf mint������Դ = 2 And Not IsNull(mrsInfo!����) Then
                        strժҪ = gclsInsure.GetItemInfo(mrsInfo!����, mrsInfo!����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        '��һ�е�����ȷ��
                        If .�շ���� = "7" And mblnPay Then Bill.ColData(5) = 4 '����
                        If .�շ���� = "F" Then Bill.ColData(11) = -1 '���ӱ�־
                        
                        '���������������
                        If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                            And Not (.�շ���� = "4" And .Detail.��������) Then
                            Bill.ColData(6) = IIF(mblnTime, 4, 5) '����
                            Bill.ColData(7) = 4 '����
                        Else
                            Bill.ColData(6) = 4 '����
                            Bill.ColData(7) = 5 '����
                        End If
                        
                        'ִ�п���
                        mblnEnterCell = False: Bill.Col = BillCol.ִ�п���: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, 10, Not blnInput) 'ֱ�ӻس�ʱ����ִ�п���
                        mblnEnterCell = False: Bill.Col = BillCol.��Ŀ: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",5,6,7,", .�շ����) > 0 Then
                            'ָ���˹̶�ҩ��ʱ,��������ѡ��
                            Select Case .�շ����
                                Case "5"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                                Case "6"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                                Case "7"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(10) = 5: .Key = 1
                        Else
                            Bill.ColData(10) = 3: .Key = Bill.ListCount
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                                                
                         '������Ŀ����,�������շ���Ŀ�д�����Ŀ����δȡ��ȡ,ҩƷ�����ж�,ҩƷ��������������
                        If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                            If (gbln��������ۿ� And mobjBill.Details(Bill.Row).�������� = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                                End If
                            End If
                        End If
                        
                    End With
                End If
                
                'ֻ����һ�θ���
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).�շ���� = "7" Then
                                '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                                'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                                Bill.Col = 5: Exit For
                            End If
                        Next
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                
                    '����ҩ���Ǵ�����Ŀ�ſɸ��ĸ���(������ı�,����Ҳ��)
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then 'And mobjBill.Details(Bill.Row).�������� = 0 Then
                        '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                        If mobjBill.Details(Bill.Row).Detail.���� Or mobjBill.Details(Bill.Row).Detail.��� Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� > mobjBill.Details(Bill.Row).Detail.��� Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" _
                                And (mobjBill.Details(i).Detail.��� Or mobjBill.Details(i).Detail.����) Then
                                If Val(Bill.Text) * mobjBill.Details(i).���� > mobjBill.Details(i).Detail.��� Then
                                    MsgBox "�� " & i & " ��ҩƷ""" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,�޸ĸ�������ÿ�治�㣡", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '���㲢ˢ�¸���
                        mobjBill.Details(Bill.Row).���� = Bill.Text
                        Call CalcMoneys(Bill.Row)
                        Call ShowDetails(Bill.Row)
                                               
                         '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" And mobjBill.Details(i).�������� = mobjBill.Details(Bill.Row).�������� Then
                                If mobjBill.Details(i).�������� = 0 Or (mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                    mobjBill.Details(i).���� = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                                                
                        Call ShowMoney
                    Else
                        sta.Panels(2) = "������Ŀ�ĸ������ܸ��ģ�"
                        Bill.Text = mobjBill.Details(Bill.Row).����: Beep '�ָ�ԭ�и���ֵ
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '�����Ϸ��Լ��
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� < 0 Then
                        'Ȩ��
                        If Not ((InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 And InStr(mstrPrivs, "ҩƷ��������") > 0) _
                             Or (InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) = 0 And InStr(mstrPrivs, "���Ƹ�������") > 0)) Then
                            MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.���� Then
                                MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint��¼���� = 2 Then
                                If Not IsNull(mrsInfo!����) Then
                                    If Not gclsInsure.GetCapability(support��������, , mrsInfo!����) Then
                                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    'ҩƷ�����
                    With mobjBill.Details(Bill.Row)
                        If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                            If .Detail.���� Or .Detail.��� Then
                                '������ʱ��ҩƷ�����ֹ����
                                If .���� * CSng(Bill.Text) > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIF(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 And Bill.ColData(10) = 5 Then
                                    '����ҩƷ�������
                                    If .���� * CSng(Bill.Text) > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ��治�㵱ǰ������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .����: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox """" & .Detail.���� & """�ĵ�ǰ��治�㵱ǰ���븶��������", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                    
                    dblPreTime = mobjBill.Details(Bill.Row).����
                    mobjBill.Details(Bill.Row).���� = Bill.Text
                    
                    '�����������
                    If Not CheckLimit(mobjBill, Bill.Row, mblnҩ����λ) Then
                        mobjBill.Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                        Cancel = True: Exit Sub
                    End If
                    If mobjBill.Details(Bill.Row).Detail.¼������ > 0 And mobjBill.Details(Bill.Row).���� > mobjBill.Details(Bill.Row).Detail.¼������ Then
                        If MsgBox("��������γ�����¼������" & mobjBill.Details(Bill.Row).Detail.¼������ & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                            mobjBill.Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                    If mobjBill.Details(Bill.Row).�������� <> 0 And mobjBill.Details(Bill.Row).Detail.���д��� <> 0 Then
                        sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                        mobjBill.Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                
                    Call CalcMoneys(Bill.Row)
                    
                    '����������(���Ѿ�������з��õ�δ��ʾǰ)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                        mobjBill.Details(Bill.Row).���� = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mint��¼���� = 2 And mrsWarn.State = 1 Then
                        curTotal = GetBillTotal(mobjBill)
                        If curTotal > 0 Then
                            cur��� = Val(txtʵ��.Tag)
                            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn, blnҽ��)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).���� = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '��������д���������
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).�������� = Bill.Row And mobjBill.Details(i).Detail.���д��� = 2 Then
                            mobjBill.Details(i).���� = Bill.Text * mobjBill.Details(i).Detail.��������
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
                        End If
                    Next
                    Call ShowMoney

                 ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                    If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                        If ItemHaveSub(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                    End If
               End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫɾ�����ã������븺��������ʵ�֣�", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If

                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '���û�ж�Ӧ��������Ŀ,���޷�����
                    If mobjBill.Details(Bill.Row).Detail.��� And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).ԭ��, mobjBill.Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).��׼����
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).��׼����
                        
                        mobjBill.Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                        Call CalcMoneys(Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        If mint��¼���� = 2 And mrsWarn.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            If curTotal > 0 Then
                                cur��� = Val(txtʵ��.Tag)
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn, blnҽ��)
                                If mbytWarn = 2 Or mbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney
                    Else
                        Bill.Text = "0"
                        sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                        Beep
                    End If
                End If
            Case "ִ�п���"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                            If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                                .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                                If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                            End If
                    
                            'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                            If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                                If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                    If .���� * .���� > .Detail.��� Then
                                        If .�շ���� = "4" Then
                                            MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                        Else
                                            MsgBox "[" & .Detail.���� & "]Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                        End If
                                        Cancel = True
                                    End If
                                Else
                                    Set colStock = IIF(.�շ���� = "4", mcolStock2, mcolStock1)
                                    If colStock("_" & .ִ�в���ID) <> 0 Then
                                        If .���� * .���� > .Detail.��� Then
                                            If colStock("_" & .ִ�в���ID) = 1 Then
                                                If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ��治�㵱ǰ������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                    Cancel = True
                                                End If
                                            ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                                MsgBox "[" & .Detail.���� & "]�ĵ�ǰ��治�㵱ǰ���븶��������", vbInformation, gstrSysName
                                                Cancel = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                            '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                            If .�շ���� = "4" And .Detail.�������� Then
                                Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                            End If
                        
                            If ItemHaveSub(Bill.Row) Then
                                KeyCode = 0
                                Call LocateMainItemNextRow(Bill.Row)
                            End If
                    End With
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            If mobjBill.Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = 1
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = 1
    End If
End Sub

Private Sub SetSubItem()
'����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
'����:
'������:Bill_KeyDown��������Ŀ��
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long               'ִ�п���ID
Dim bln��������ۿ� As Boolean
Dim strժҪ As String

lngMainRow = Bill.Row               '�������
If gbln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
    bln��������ۿ� = Not mobjBill.Details(lngMainRow).Detail.���ηѱ�
End If

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)
            mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, 1) = "" '�б�Ҫ����
        
        'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
             If mcolDetails(i).��� = .�շ���� Then
                '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                lngDoUnit = .ִ�в���ID
             Else
                If mcolDetails(i).ִ�п��� = 0 Then
                    '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                    lngDoUnit = .ִ�в���ID
                Else
                    '������ҩ��Ŀ��ִ�п���
                    lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(i).���, _
                        mcolDetails(i).ID, mcolDetails(i).ִ�п���, lngDoUnit, Get��������ID, mint������Դ)
                End If
             End If
        End If
        
        'b.������ĿΪҩƷ,���ĵ�ִ�п���(��������ִ�п���Ϊ��,Ҳ��ִ�е�����)
        If lngDoUnit = 0 Then
            lngDoUnit = mobjBill.����ID
            If lngDoUnit = 0 And cbo��������.ListIndex <> -1 Then
                lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
            End If
            lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(i).���, mcolDetails(i).ID, _
                mcolDetails(i).ִ�п���, lngDoUnit, Get��������ID, mint������Դ, .ִ�в���ID) '���Ĵ���ȱʡ������ִ�п�����ͬ
        End If
            
        '����֧����Ŀ��Ӧ���
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                If Not ItemExistInsure(mcolDetails(i).ID, mrsInfo!����) Then
                    If gintҽ������ = 1 Then
                        If MsgBox("��Ŀ""" & mcolDetails(i).���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        End If
                    ElseIf gintҽ������ = 2 Then
                        MsgBox "��Ŀ""" & mcolDetails(i).���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln��������ۿ�)
        Call ShowDetails(Bill.Rows - 1)
        
        If mrsInfo.State = 1 Then
             If Not IsNull(mrsInfo!����) Then
                'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
                strժҪ = mobjBill.Details(Bill.Rows - 1).ժҪ
                
                strժҪ = gclsInsure.GetItemInfo(mrsInfo!����, mrsInfo!����ID, mcolDetails(i).ID, strժҪ, 1)
                mobjBill.Details(Bill.Rows - 1).ժҪ = strժҪ
             End If
        End If
    Next
    
    If bln��������ۿ� And Not mbln���õǼ� Then
        Call CalcMoney(lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ�����ǰ������������.
        
        Call Calc��������ʵ��(lngMainRow)
    End If
    
    Call ShowMoney
End With

End Sub

Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function

Private Sub Calc��������ʵ��(ByVal lngMainRow As Long)
'����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
'����:  lngMainRow-������ID

Dim i As Long, j As Long
Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
Dim cur���ۺ�ʵ�� As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).�������� <> lngMainRow Then Exit For    '��ȻĿǰ�����˲������ڴ����м������������,����һ�ŵ�����������,Ϊ�˽������ܵ�����,����ȫ��ɨ��
        
        If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
            Next
        End If
    Next
   
    cur���ۺ�ʵ�� = CCur(Format(ActualMoney(.�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�), gstrDec))
    
    cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
    
    .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gstrDec)
    .Details(lngMainRow).InComes(1).Key = "_" & Format(cur���ۺ�ʵ��, gstrDec)
    
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubDept(ByVal lngRow As Long)
Dim i As Long, j As Long
    With mobjBill
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID) '������ȡ
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                If cbo��������.ListIndex <> -1 Then
                                    .Details(i).ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                                End If
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(j).���, _
                                    mcolDetails(j).ID, mcolDetails(j).ִ�п���, .Details(i).ִ�в���ID, Get��������ID, mint������Դ)
                            End If
                        End If
                    End If
                    
                    '��ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.Details(i).ִ�в���ID)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.Details(i).ִ�в���ID)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function ItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                ItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim strStock As String, i As Long
    
    If Not Bill.Active Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If Not mblnEnterCell Then Exit Sub
    
    If mbytInState = 3 Then
        '����б༭����������ɫ
        Bill.SetColColor 1, &HE7CFBA '��ȻҪ�ɰ�ɫ
        Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                If mbln����ҩ�� Or mbln����ҩ�� Then
                    strStock = GetStockInfo(.�շ�ϸĿID, mbln����ҩ��, mbln����ҩ��, mblnҩ����λ, mstrҩ����װ)
                    If strStock <> "" Then sta.Panels(2) = "��" & Bill.Row & "�п��:" & strStock
                End If
                If strStock = "" Then
                    '��ʱ���¿����ʾ
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If mblnҩ����λ Then
                        .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                    End If
                    sta.Panels(2) = "[" & .Detail.���� & "]���ÿ��:" & .Detail.���
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                sta.Panels(2) = "[" & .Detail.���� & "]���ÿ��:" & .Detail.���
            ElseIf .Detail.��� And .InComes.Count > 0 And Bill.TextMatrix(0, Bill.Col) = "����" Then
                sta.Panels(2) = "�۸�Χ:" & FormatEx(.InComes(1).ԭ��, 5) & "-" & FormatEx(.InComes(1).�ּ�, 5)
            Else
                sta.Panels(2) = ""
            End If
            
            Bill.ColData(1) = IIF(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(2) = BillColType.CommandButton
            
             '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If ItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(1) = BillColType.Text_UnModify
                Bill.ColData(2) = BillColType.Text_UnModify
            End If
            
            '����Ƿǵ���״̬
            If mbytInState <> 2 Then
                If .�շ���� = "7" And mblnPay Then
                    Bill.ColData(5) = 4
                Else
                    Bill.ColData(5) = 5
                End If
                
                '���������������
                If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                    And Not (.�շ���� = "4" And .Detail.��������) Then
                    Bill.ColData(6) = IIF(mblnTime, 4, 5) '����
                    Bill.ColData(7) = 4 '���
                Else
                    Bill.ColData(6) = 4
                    Bill.ColData(7) = 5
                End If
                
                If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                    Bill.ColData(10) = BillColType.UnFocus
                Else
                    Bill.ColData(10) = BillColType.ComboBox
                End If
                
                If .�շ���� = "F" Then
                    Bill.ColData(11) = -1
                Else
                    Bill.ColData(11) = 5
                End If
                
                 'ֻ����һ�����
                If mblnOne Then Bill.ColData(1) = 5
            End If
        End With
    End If
   
    '������δ�������,��ָ��е�����
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(1) = IIF(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
        Bill.ColData(2) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı��������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '�������
    End If
    
    If gbln�շ���� And Bill.TextMatrix(Row, 1) = "" And mblnOne Then
        mrsClass.Filter = "����=" & mstr�շ����
        Bill.TextMatrix(Row, 1) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "���" '�������շ����ʱ������������
            Call zlControl.CboSetWidth(Bill.CboHwnd, 1000)
            '������Ϊ��,���Զ�Ĭ��Ϊ��һ�շ�ϸĿ�����
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & mstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���"
            Call zlControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                '�ɷ�����С��
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivs, "ҩƷС������") > 0 Then
                        Bill.TextMask = "." & Bill.TextMask
                    End If
                Else
                    Bill.TextMask = "." & Bill.TextMask
                End If
                
                '�ɷ����븺��
                If Not mobjBill.Details(Bill.Row).Detail.���� Then
                    If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivs, "ҩƷ��������") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivs, "���Ƹ�������") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint��¼���� = 2 Then
                            If Not IsNull(mrsInfo!����) Then
                                If Not gclsInsure.GetCapability(support��������, , mrsInfo!����) Then
                                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case "����"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
   
    
    '��ʾժҪ
    If mobjBill.Details.Count >= Bill.Row Then
        If mobjBill.Details(Bill.Row).ժҪ <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  ժҪ:" & mobjBill.Details(Bill.Row).ժҪ
        End If
    End If
    
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'bill.ToolTipText = bill.TextMatrix(bill.MouseRow, bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ListIndex
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, strDoctor As String
    
    If mbytInState <> 0 Then Exit Sub
    
    '��λҽ��
    cbo������.Clear
    If cbo��������.ListIndex <> -1 Then
        Call Load������(cbo��������.ItemData(cbo��������.ListIndex))
    End If
    
    '���ݶ���
    If mbytInState = 0 Then
        If cbo��������.ListIndex = -1 Then
            mobjBill.��������ID = 0
        Else
            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
    End If
    
    If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0 '����click�¼�
    cboBaby.Enabled = DeptIsWoman(mobjBill.��������ID)
    
    '�������������Ŀ��ִ�п���
    If mbytInState = 0 And cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '�������շ���Ŀ
                If InStr(",4,5,6,7,", .Detail.���) = 0 And .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                    '����8113���޸�
'                ElseIf InStr(",4,5,6,7,", .Detail.���) > 0 Then
'                '�������ҩ��Ϊ�洢�ⷿ�����õķ����ڲ��˿���(��������)��ִ�п���
'                    If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
'                        Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
'                    End If
'                    If .Key = "1" Then .Key = "0"        '1��ʾִ�п��Ҳ���ѡ��
'                    mlngPreRow = 0      '������Entercell�¼���������ִ�п��ҵĿ�ѡ��
                End If
            End With
        Next
    End If
End Sub

Private Sub cbo������_Click()
    Dim arrDepts As Variant, i As Long, k As Long
    
    If mbytInState = 0 Then
        mobjBill.������ = IIF(cbo������.ListIndex = -1, "", NeedName(cbo������.Text))
                        
        '��ʿ���
        If Bill.Active Then
            If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
                And Bill.RowData(Bill.Rows - 1) <> 0 Then
                '�����Ч����
                Bill.TextMatrix(Bill.Rows - 1, 1) = ""
                Bill.RowData(Bill.Rows - 1) = 0
            ElseIf Bill.Col = 1 Then
                Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
            End If
        End If
        
        '��ʿ���:�жϷǷ�����
'        If HaveStopClass > 0 Then
'            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
'        End If
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�Ӱ�_Click()
    If mbytInState = 1 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIF(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 Or gblnOK Then
        If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckNegative() As Boolean
'���ܣ���鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strItems As String, str���� As String
    Dim str��λ As String, dbl���� As Double
    
    CheckNegative = True
    If mobjBill.����ID = 0 Then Exit Function
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                strItems = strItems & ",(" & .�շ�ϸĿID & "," & .ִ�в���ID & ")"
                strSQL = strSQL & " Union ALL Select " & .�շ�ϸĿID & "," & .ִ�в���ID & ",0 From Dual"
            End If
        End With
    Next
    strItems = Mid(strItems, 2)
    If strItems = "" Then Exit Function
    
    strSQL = _
        " Select �շ�ϸĿID,ִ�в���ID,Sum(Nvl(����,1)*����) as ����" & _
        " From ���˷��ü�¼" & _
        " Where (�շ�ϸĿID+0,ִ�в���ID+0) IN(" & strItems & ")" & _
        " And ��¼״̬<>0 And ���ʷ���=1 And �۸񸸺� is NULL" & _
        " And ����ID=" & mobjBill.����ID & " And Nvl(��ҳID,0)=" & mobjBill.��ҳID & _
        " Group by �շ�ϸĿID,ִ�в���ID" & strSQL
    strSQL = "Select �շ�ϸĿID,ִ�в���ID,Sum(����) as ���� From (" & strSQL & ") Group by �շ�ϸĿID,ִ�в���ID"
    
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'Union:��������
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID & " And ִ�в���ID=" & .ִ�в���ID
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        str��λ = .Detail.ҩ����λ
                        dbl���� = Nvl(rsTmp!����, 0) / .Detail.ҩ����װ
                    Else
                        str��λ = .Detail.���㵥λ
                        dbl���� = Nvl(rsTmp!����, 0)
                    End If
                    str���� = Get��������(.ִ�в���ID)
                    If Abs(.����) * .���� > dbl���� Then
                        MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(Abs(.����) * .����, 5) & str��λ & _
                            " �����ѼƷ����� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                        CheckNegative = False: Exit Function
                    End If
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, strTmp As String
    Dim i As Long, j As Long, lng����ID As Long
    Dim blnҽ�� As Boolean, cur���ն� As Currency
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur��� As Currency
    Dim colStock As Collection
    
    If mbytInState = 3 Then
        If mint��¼���� <> 1 And False Then '������ȫ��ɾ��
            For i = 1 To Bill.Rows - 1
                If Bill.TextMatrix(i, Bill.Cols - 1) = "��" And Bill.RowData(i) > 0 Then
                    strSQL = strSQL & "," & Bill.RowData(i)
                End If
            Next
            If strSQL = "" Then
                MsgBox "������ѡ��һ��Ҫɾ���ķ��ã�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '������ѡ����
            strSQL = Mid(strSQL, 2)
            i = GetBillRows(mstrInNO, mint��¼����)
            If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        Else
            '��ΪҪ����Ϊȫ�ˣ�������ʺ��������ʣ����ݽ��ʺ��Ҫ���
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint��¼����)
            If j < i Then
                MsgBox "�����еĲ�����Ŀ��ǰ�Ѳ���������(�����ѽ��ʵ���Ŀ)��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        'ҽ�����������ϴ�(ע���ж�˳��)
        If mint������Դ = 2 Then
            intInsure = BillExistInsure(mstrInNO) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, , intInsure) Then
                    'ȥ����ҽ������ƥ����
                    If strSQL <> "" Then '���ܲ�������
                        MsgBox "��Ϊҽ��������Ҫ,�õ����е���Ŀ����ȫ��ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mint������Դ = 2 Then
            strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            If mint��¼���� = 2 Then
                strSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Else
                strSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & mstrInNO & "')"
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        'ҽ�����������ϴ�
        If mint������Դ = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        'ҽ�����������ϴ�
        If mint������Դ = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ɾ��������ҽ������ʧ�ܣ��õ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0
        
        gblnOK = True: Unload Me: Exit Sub
    Else '�������뵥��״̬
        If mobjBill.����ID = 0 Or mrsInfo.State = 0 Then
            MsgBox "û�з��ֲ�����Ϣ�����ݲ��ܱ��档", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mobjBill.Details.Count = 0 Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Checkִ�п���
        If i <> 0 Then
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If cbo��������.ListIndex = -1 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus: Exit Sub
        End If
        
        If cbo������.ListIndex = -1 Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��ʿ���:�жϷǷ�����
'        If HaveStopClass > 0 Then
'            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
'            Exit Sub
'        End If
                
        '����ʱ����
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        If mint������Դ = 2 Then
            If Not PatiCanBilling(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), mstrPrivs) Then Exit Sub
        End If
                
        '����ʱ����
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "���õķ���ʱ�䲻��С�ڲ��˵���Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '�Ƿ���
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
             '8407
'            ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
'                '�ռ�ҩƷ�ķ�ҩҩ��
'                strTmp = strTmp & "," & mobjBill.Details(i).�շ�ϸĿID
            End If
        Next
        
'        '���ҩƷ�ķ�ҩҩ����Ӧ�ķ������(�洢�ⷿ)
'        If strTmp <> "" Then
'            strTmp = Mid(strTmp, 2)
'            Set rsTmp = GetServiceDept(strTmp)
'            If Not rsTmp Is Nothing Then
'                strTmp = ""
'                For i = 1 To mobjBill.Details.Count
'                    If InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
'                        strInfo = mobjBill.Details(i).�շ�ϸĿID
'                        '�ȼ���Ƿ�������Ĵ洢�ⷿ
'                        rsTmp.Filter = "�շ�ϸĿID=" & strInfo & " And ִ�п���id=" & mobjBill.Details(i).ִ�в���ID
'                        If rsTmp.RecordCount = 0 Then
'                            strTmp = strTmp & "," & i
'                        Else
'                            '�ټ���Ƿ�������ķ������(û�����÷�����ҵ�,��������IDΪ��)
'                            rsTmp.Filter = "(" & rsTmp.Filter & " And ��������ID=" & mobjBill.��������ID & ") Or (" & rsTmp.Filter & " And ��������ID=0)"
'                            If rsTmp.RecordCount = 0 Then
'                                strTmp = strTmp & "," & i
'                            End If
'                        End If
'                    End If
'                Next
'                If strTmp <> "" Then
'                    strTmp = Mid(strTmp, 2)
'                    MsgBox "����,��" & strTmp & "��ҩƷ�Ƿ�Υ�����¹���:" & vbCrLf & vbCrLf & _
'                        "A.ѡ���ִ�п��Ҳ���ҩƷ�Ĵ洢�ⷿ" & vbCrLf & _
'                        "B.��������[" & NeedName(cbo��������.Text) & "]������ҩƷ�ڴ˴洢�ⷿ�ķ������.", _
'                        vbInformation, gstrSysName
'                    Exit Sub
'                End If
'            End If
'        End If
        
        
        '*********��ģ�鲻�������䵥����ȷ�����˵����**********
        'ҽ���������ʼ��    ��Ϊ����Ա�������䵥��,��ȷ������,����Ҫ�ټ��һ��(�˴������ж�Ȩ��,��Ϊ��Ȩ�޲ſ����Ǹ���)
'        If InStr(mstrPrivs, "��������") > 0 And mint��¼���� = 2 Then    '����������һ�ָ�������Ȩ��,�ſ����и���
'            If Not IsNull(mrsInfo!����) Then
'                If Not gclsInsure.GetCapability(support��������, , mrsInfo!����) Then
'                    For i = 1 To mobjBill.Details.Count
'                        If mobjBill.Details(i).���� * mobjBill.Details(i).���� < 0 Then
'                            MsgBox "�����е� " & i & " ���Ǹ���,����ҽ����֧�ָ������ʣ�", vbInformation, gstrSysName
'                            Bill.SetFocus: Exit Sub
'                        End If
'                    Next
'                End If
'            End If
'        End If
                
        '����ְ����
        If InStr(",1,2,", txt���ʽ.Tag) > 0 Then '���ѻ�ҽ������
            i = CheckDuty(, False)
            If i > 0 Then
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = 2: Bill.SetFocus
                Exit Sub
            End If
        End If

        '���в�����Ŀ
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = 2: Bill.SetFocus
            Exit Sub
        End If
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
        
        'Ҫ������,ҽ��������Ŀ�������
        If mint������Դ = 2 And mint��¼���� = 2 Then
            If Not IsNull(mrsInfo!����) And Not mrsMedAudit Is Nothing Then
                If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!����) Then Exit Sub
            End If
        End If
        
        '���ʷ��౨��
        If mint��¼���� = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '���ݷ���
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                'ˢ�²��˷���״��
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, mcurModiMoney, True)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!Ԥ�����
                    cmdCancel.Tag = rsTmp!�������
                    txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                End If
                sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
                sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - curTotal, "0.00")
                
                '���¶�ȡ���ն�
                cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                
                '�Ƿ�ҽ������
                blnҽ�� = txt���ʽ.Tag = "1" Or Not IsNull(mrsInfo!����)
                cur��� = Val(txtʵ��.Tag)
                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID)
                        
                For i = 1 To mobjBill.Details.Count
                    mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, cur���ն� - mcurModiMoney, curTotal, IIF(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, mintWarn, blnҽ��)
                    If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                Next
            End If
        End If
        
        'ҩƷ���ɼ��
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(�������)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
                    
        '�����������
        If Not CheckLimit(mobjBill, , mblnҩ����λ) Then Exit Sub
        
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.���� Or .Detail.���) _
                    And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Details(j).ִ�в���ID Then
                            If .�շ���� = "4" Then
                                MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����룬��ϲ���", vbInformation, gstrSysName
                            End If
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        'ҩƷ�����(�������ֹʱ�����ʱ��ҩƷ)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIF(.�շ���� = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If mblnҩ����λ Then
                            .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        End If
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���""" & .Detail.��� & """������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If mblnҩ����λ Then
                            .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        End If
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���""" & .Detail.��� & """������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���""" & .Detail.��� & """������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���""" & .Detail.��� & """������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Next
        
        '����������ϵ����Ч��
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .�շ���� = "4" And .Detail.�������� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblTotal) Then Exit Sub
                End If
            End With
        Next
        
        '�����˷Ѽ��
        If mint��¼���� = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        If Not SaveBill Then Exit Sub
        
        If mstrInNO <> "" Then
            gblnOK = True: Unload Me: Exit Sub
        Else
            sta.Panels(2) = "��һ�ŵ���:" & mobjBill.NO
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call SetMoneyList
            Call NewBill
            
            '���¶�ȡ������Ϣ
            Call GetPatient(mlng����ID, mlng��ҳID)
            
            Bill.SetFocus
        End If
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub Form_Activate()
    If mbytInState <> 0 Then
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
        ElseIf cmdCancel.Visible And cmdCancel.Enabled Then
            cmdCancel.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',;|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill, i As Long
    
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    
    gblnOK = False
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    Call InitLocPar
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Then
        If Not InitData Then
            Unload Me: Exit Sub
        End If
    End If
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then
        If Not ReadBill(mstrInNO, mbytInState = 3) Then
            Unload Me: Exit Sub
        End If
    Else
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportBill(mint������Դ, mstrInNO, mint��¼����, mbln���õǼ�)
            If mobjBill.NO = "" Then
                MsgBox "������ȷ��ȡ�Ʒѵ��ݵ����ݣ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '����б༭����������ɫ
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 2, &HE7CFBA
                Bill.SetColColor 6, &HE7CFBA
                Bill.SetColColor 10, &HE7CFBA
                Bill.SetColColor 5, &HE0E0E0
                Bill.SetColColor 7, &HE0E0E0
                Bill.SetColColor 11, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                
                mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
                Call GetCboIndex(cbo������, mobjBill.������, True)
                mobjBill.������ = NeedName(cbo������.Text)
                
                Call zlControl.CboSetIndex(cboBaby.Hwnd, mobjBill.Ӥ����)
                If cbo��������.ListIndex <> -1 Then cboBaby.Enabled = DeptIsWoman(mobjBill.��������ID)
                
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                If mint��¼���� = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '�ڶ�ȡ����ǰȡ
                End If
                
                '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
                Call GetPatient(mlng����ID, mlng��ҳID)
                If mrsInfo.State = 0 Then
                    MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                If gbln��������ۿ� Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '�������:�޸�ʱ���Ͻ�Ҫ�˻صĿ��
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc(.�շ����) '���⴦��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            .Detail.��� = .Detail.��� + .���� * .����
                        ElseIf .�շ���� = "4" And .Detail.�������� Then
                            .Detail.��� = .Detail.��� + .���� * .����
                        End If
                    End With
                Next
                
                Call SetColNum
            End If
        Else
            '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
            Call GetPatient(mlng����ID, mlng��ҳID)
            If mrsInfo.State = 0 Then
                MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        End If
        
        If mstrInNO <> "" And mint��¼���� = 2 And mint������Դ = 2 Then
            Call ReCalcInsure '���¼���ͳ����
        End If
        
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - fraTitle.Height - fraInfo.Height + 230
    
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    
    cboNO.Left = fraTitle.Width - cboNO.Width - 90
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraUnit.Left = Me.ScaleWidth - fraUnit.Width
    fraInfo.Width = Me.ScaleWidth - fraUnit.Width - fraInfo.Left
    
    Bill.Width = Me.ScaleWidth - Bill.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
        
    cbo������.Left = lblDate.Left - cbo������.Width - 200
    lbl������.Left = cbo������.Left - lbl������.Width - 45
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 500
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mbytInState)
    
    mlngҽ��ID = 0
    mlng���ͺ� = 0
    mlng����ID = 0
    mlng��ҳID = 0
    mint������Դ = 0
    mint��¼���� = 0
    mbln���õǼ� = False
    mlng��������ID = 0
    mlng���˿���id = 0
    
    mlng��������ID = 0
    mstr����ҽ�� = ""
    
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    
    mbytInState = 0
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mstrPrivs = ""
    
    Set mrsInfo = Nothing
    Set mrsUnit = Nothing
    Set mrsClass = Nothing
    Set mrsWork = Nothing
    Set mrsMedAudit = Nothing
    
    If Not InDesign Then
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", _
            IIF(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIF(sta.Panels("WB").Bevel = sbrInset, 1, 0))
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    If mbytInState = 3 Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        If mbytInState <> 2 Then
            .ColData(1) = IIF(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
            .ColData(2) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
            .ColData(5) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(7) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(11) = 5 '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        End If
        
        '����б༭����������ɫ
        .SetColColor 1, &HE7CFBA
        .SetColColor 2, &HE7CFBA
        .SetColColor 6, &HE7CFBA
        .SetColColor 10, &HE7CFBA
        .SetColColor 5, &HE0E0E0
        .SetColColor 7, &HE0E0E0
        .SetColColor 11, &HE0E0E0
        
        .TextMatrix(Row, 0) = Row
        
        '����ط��ֶ����ò�ִ��
        If Row > 0 And .ColData(1) <> 5 And Me.Visible And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo��������.ListIndex <> -1 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 And Not cbo��������.Locked Then
        lngIdx = zlControl.CboMatchIndex(cbo��������.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo��������.ListCount > 0 Then lngIdx = 0
        cbo��������.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        strText = cbo������.Text
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            If IsNumeric(strText) Then
                For i = 0 To cbo������.ListCount - 1
                    If i > UBound(marrDr) Then Exit For
                    If CStr(Split(marrDr(i), "|")(2)) = strText Then
                        If intIdx = -1 Then cbo������.ListIndex = i
                        intIdx = i
                    End If
                Next
                If intIdx = -1 Then
                    For i = 0 To cbo������.ListCount - 1
                        If i > UBound(marrDr) Then Exit For
                        If Val(Split(marrDr(i), "|")(2)) = Val(strText) Then
                            If intIdx = -1 Then cbo������.ListIndex = i
                            intIdx = i
                        End If
                    Next
                End If
            Else
                For i = 0 To cbo������.ListCount - 1
                    If UCase(cbo������.List(i)) Like UCase(strText) & "*" Then
                        If intIdx = -1 Then cbo������.ListIndex = i
                        intIdx = i
                    End If
                Next
            End If
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = ""
        Else
            mobjBill.������ = NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp "zl9InExse", Me.Hwnd, "frmCharge"
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF6 '�����ǰ��������,�����µ�״̬
            If mbytInState = 0 Then
                txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
                Call ClearRows: Call Bill.ClearBill
                Call SetColNum: Call ClearMoney
                Call NewBill
                Bill.SetFocus
            End If
        Case vbKeyF7 '�л����뷨
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape, vbKeyX
            If KeyCode = vbKeyX And Shift <> 4 Then Exit Sub
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Sub SetMoneyList()
'����:���ݵ�ǰ������Ŀ�����������п�
    Dim lngW As Long
    lngW = mshMoney.Width - 60
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    mshMoney.ColWidth(0) = lngW * 0.5
    mshMoney.ColWidth(1) = lngW * 0.5
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    
    mshMoney.TextMatrix(0, 0) = "��Ŀ"
    mshMoney.TextMatrix(0, 1) = "���"
    mshMoney.Row = 0
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '��������
    strSQL = "Select ��������ID,����ҽ�� From ����ҽ����¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
    If Not rsTmp.EOF Then
        mlng��������ID = Nvl(rsTmp!��������ID, 0)
        mstr����ҽ�� = Nvl(rsTmp!����ҽ��)
    End If
    If mlng��������ID = 0 Or mstr����ҽ�� = "" Then
        MsgBox "û�з���Դҽ����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    If mbln���õǼ� Then
        '��Ϊ��ǰѡ���ҽ������
        strSQL = "Select ID,����,����,���� From ���ű� Where ID=[1]"
    Else
        '��Ϊ��ǰѡ���ҽ�����һ�������
        strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([1],[2]) Order by ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��������ID, mlng��������ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��������.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo��������.ItemData(cbo��������.ListCount - 1) = rsTmp!ID
            If rsTmp!ID = mlng��������ID Then
                cbo��������.ListIndex = cbo��������.NewIndex
            End If
            rsTmp.MoveNext
        Next
        If cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    Else
        MsgBox "����ȷ���������ң����ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����շ����:"'5','E','Z'"
    If mstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where Instr([1],����)>0 Order by ���"
    End If
    'Set mrsClass = New ADODB.Recordset
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�շ����)
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(mstr�շ����, "'5'") > 0 Or InStr(mstr�շ����, "'6'") > 0 Or InStr(mstr�շ����, "'7'") > 0 Or mstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(IIF(mint��¼���� = 1, 8, 9))
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ��ҩƷ���ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(mstr�շ����, "'4'") > 0 Or mstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(IIF(mint��¼���� = 1, 40, 41))
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ִ�в���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID and B.������� IN([1],3) " & _
        " Order by B.�������,A.����"
    'Set mrsUnit = New ADODB.Recordset
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint������Դ)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str��� As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'���ܣ���ȡ����������ͬ�����Ŀ��ִ�п���ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).�շ���� = str��� _
            And mobjBill.Details(i).ִ�в���ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                Exit Function
            End If
        End If
    Next
    If str��� = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'���ܣ����ݵ��������������б������
'������blnEnter=�Ƿ񰴽�����д���,����ִ�п��ұ��ֲ���
    Dim rsTmp As New ADODB.Recordset
    Dim str��Ա���� As String, strTmp As String
    Dim lng����ID As Long, strIDs As String
    Dim strSQL As String, i As Long
    
    Bill.Clear
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "���"
            If cbo������.ListIndex <> -1 Then
                If cbo������.ListIndex <= UBound(marrDr) Then
                    If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                        str��Ա���� = Split(marrDr(cbo������.ListIndex), "|")(6)
                    End If
                End If
            End If
        
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    '��ʿ���:����
'                    If Not (str��Ա���� = "��ʿ" And InStr(",E,M,4,", mrsClass!����) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!���
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
'                    End If
                    mrsClass.MoveNext
                Next
            End If
        Case "ִ�п���"
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����ID = GetLastDeptID(.�շ����, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����ID = 0 Then lng����ID = .ִ�в���ID
                            
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem mrsWork!���� & "-" & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        lng����ID = Get��������ID
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        '0-����ȷ,1-���˿���,2-���˲���,3-�����˿���,4-ָ������
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!����ID, 0) & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!����ID, 0) & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSQL = "Select Nvl(��������ID,0) as ��������ID,ִ�п���ID" & _
                                    " From �շ�ִ�п���" & _
                                    " Where �շ�ϸĿID=[1]" & _
                                    " And (������Դ is NULL Or ������Դ=[2])" & _
                                    " And (��������ID is NULL Or ��������ID=[3])" & _
                                    " Order by Decode(������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, mint������Դ, Val(Nvl(mrsInfo!����ID, 0)))
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                            Case 6 '�����˿���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = mrsUnit!���� & "-" & mrsUnit!����
                                If Not (SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = mrsUnit!ID
                                    
                                    '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).�շ����) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.ִ�п��� = 4 Then 'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
End Sub

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '���õ��ݱ��ʽ
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = 2
        .PrimaryCol = 2
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, 0) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(0) = 5
            
            .ColData(1) = IIF(gbln�շ����, 3, 5)
            If mblnOne Then .ColData(1) = 5
            
            .ColData(2) = 1 '��Ŀ����,��Ť��ѡ
            .ColData(6) = 4 '��/������
            .ColData(3) = 5 '�������
            .ColData(4) = 5 '��λ����
            .ColData(5) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(7) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(8) = 5 'Ӧ�ս������
            .ColData(9) = 5 'ʵ�ս������
            .ColData(10) = 3 'Ĭ��ȡ�������һ���һ����
            .ColData(11) = 5 '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
            .ColData(12) = 5 '����ȱʡ����
        End If
        .SetColColor 1, &HE7CFBA
        .SetColColor 2, &HE7CFBA
        .SetColColor 6, &HE7CFBA
        .SetColColor 10, &HE7CFBA
        .SetColColor 5, &HE0E0E0
        .SetColColor 7, &HE0E0E0
        .SetColColor 11, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    Call SetMoneyList

    '��ȡ����ƥ�䷽ʽ
    sta.Panels("PY").Visible = mbytInState = 0
    sta.Panels("WB").Visible = mbytInState = 0
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���
        i = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0))
        If i = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf i = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If

    '����
    If mbln���õǼ� Then
        lblTitle.Caption = gstrUnitName & "��Ѻ��õǼ�"
    Else
        If mint��¼���� = 1 Then
            lblTitle.Caption = gstrUnitName & "�����շѵ�"
        ElseIf mint��¼���� = 2 Then
            lblTitle.Caption = gstrUnitName & "���˼��ʵ�"
        End If
    End If
    txtӦ��.Text = gstrDec: txtʵ��.Text = gstrDec
    
    
    arrBaby = Array("0-���˱���", "1-��1��Ӥ��", "2-��2��Ӥ��", "3-��3��Ӥ��", "4-��4��Ӥ��", "5-��5��Ӥ��")
    For i = 0 To UBound(arrBaby)
        cboBaby.AddItem arrBaby(i)
    Next
    cboBaby.ListIndex = 0
    
    Select Case mbytInState
        Case 0 'ִ��
            Call SetShowCol
        Case 1 '����
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Case 3 '����
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            
            '��ʱ��֧�ֲ���ɾ��
            If mint��¼���� <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    txt����.Locked = Not bln
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln
    
    chk�Ӱ�.Enabled = bln
    cboBaby.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���ȡ������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    mintWarn = -1: mstrWarn = ""
    Set mrsWarn = New ADODB.Recordset
    
    txt����.ForeColor = Me.ForeColor
    Set mrsInfo = New ADODB.Recordset
    
    If mint������Դ = 2 Then '��סԺ�����Ƿ����ǿ�Ƽ���Ȩ��
        If InStr(mstrPrivs, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
            strSQL = ""
        ElseIf InStr(mstrPrivs, "��Ժδ��ǿ�Ƽ���") > 0 Then
            strSQL = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
        ElseIf InStr(mstrPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
            strSQL = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
        Else
            strSQL = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
        End If
    End If
    
    '�ֶ���ʹ�ò���ʱ���������ȷ����(��Nullֵ),����ΪadVarChar����
    strSQL = "Select" & _
        " A.����ID,B.��ҳID,To_Number(Nvl(B.��ǰ����ID,[3])) as ����ID," & _
        " Nvl(B.��Ժ����ID,[3]) as ����ID,B.��Ժ����,B.��Ժ����," & _
        " A.�����,A.סԺ��,B.��Ժ���� as ����,A.����,A.�Ա�,A.����,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
        " A.������," & IIF(mint������Դ = 2 And mint��¼���� = 2, "Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,", "A.������,") & _
        " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Y.���� as ������," & _
        " zl_PatiDayCharge(A.����ID) as ���ն�,Nvl(B.����,A.����) as ����,Nvl(B.��������,0) as ��������" & _
        " From ������Ϣ A,������ҳ B,������� X,ҽ�Ƹ��ʽ Y" & _
        " Where A.����ID=B.����ID(+) And A.����ID=X.����ID(+)" & strSQL & _
        " And A.����ID=[1] And B.��ҳID(+)=[2] And A.ҽ�Ƹ��ʽ=Y.����(+)"
        
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, mlng���˿���id)
    If Not mrsInfo.EOF Then
        If Not IsNull(mrsInfo!����) Then
            txt����.ForeColor = vbRed
        End If
        cboBaby.ListIndex = 0
        cboBaby.Enabled = DeptIsWoman(Val("" & mrsInfo!����ID))
        
        '�������ﻮ������Ҫ���������
        If mint��¼���� = 2 Then
            'ˢ�²��˷���״��
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, mcurModiMoney, True)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!Ԥ�����
                cmdCancel.Tag = rsTmp!�������
                txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
            End If
            sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag), gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag), "0.00")
            
            'ˢ�±�����Ϣ
            strSQL = "Select Nvl(���ò���,1) as ���ò���,Nvl(��������,1) as ��������," & _
                " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
                " Where " & IIF(mint������Դ = 1, "����ID is NULL", "����ID=[1]")
            Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!����ID, 0)))
        End If
                            
        'סԺ���ʲŴ��������
        If mint������Դ = 2 Then
            '�������
            If Not IsNull(mrsInfo!����) Then
                chk����.Value = 0: chk����.Visible = True
            Else
                chk����.Value = 0: chk����.Visible = False
            End If
            
            '����ʱ��
            If Not IsNull(mrsInfo!��Ժ����) Then
                txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        '��ʾ������Ϣ
        txt����.Text = Nvl(mrsInfo!����)
        txt�Ա�.Text = Nvl(mrsInfo!�Ա�)
        txt����.Text = Nvl(mrsInfo!����)
        txt�ѱ�.Text = Nvl(mrsInfo!�ѱ�)
        txt���ʽ.Text = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
        txt���ʽ.Tag = Nvl(mrsInfo!������, 0) '��Ҫ��дΪ��
        txt����.Text = Nvl(mrsInfo!����)
        txt������.Text = Nvl(mrsInfo!������)
        txt������.Text = Format(Nvl(mrsInfo!������), "0.00")
        
        With mobjBill
            .����ID = Nvl(mrsInfo!����ID, 0)
            .��ҳID = Nvl(mrsInfo!��ҳID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .���� = Nvl(mrsInfo!����)
            .��ʶ�� = IIF(mint������Դ = 1, Nvl(mrsInfo!�����, 0), Nvl(mrsInfo!סԺ��, 0))
            .���� = Nvl(mrsInfo!����)
            .�Ա� = Nvl(mrsInfo!�Ա�)
            .���� = Nvl(mrsInfo!����)
            .�ѱ� = Nvl(mrsInfo!�ѱ�)
        End With
        
        '�ڵ�һ�ν���ʱ��ȡ��������������Ŀ��Ϣ
        If Not Visible And mint������Դ = 2 And mint��¼���� = 2 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    Dim strMainRows As String
    Dim bln��������ۿ� As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIF(lngRow = 0, 1, lngRow) To IIF(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln��������ۿ� = False
        If gbln��������ۿ� And Not mbln���õǼ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
            If mobjBill.Details(i).�������� > 0 Then    '����
                bln��������ۿ� = Not mobjBill.Details(mobjBill.Details(i).��������).Detail.���ηѱ�
                If bln��������ۿ� And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).��������      '��������һ�е�ʱ��
            Else
                If ItemHaveSub(i) Then                          '����������
                     bln��������ۿ� = Not mobjBill.Details(i).Detail.���ηѱ�
                     If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln��������ۿ�)
    Next
    
    '������������,������bln��������ۿ۱���,��Ϊ�������������Ǵ������ʱ�Ѹı�
    If gbln��������ۿ� And Not mbln���õǼ� Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc��������ʵ��(Split(strMainRows, ",")(i))
        Next
    End If
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln��������ۿ� As Boolean)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim rsTmp As New ADODB.Recordset
    Dim strInfo As String, i As Long
    Dim dblMoney As Double '�û�����ı�۽��
    Dim dbl�Ӱ�Ӽ��� As Double
    
    Dim rsPrice As New ADODB.Recordset '���ڼ���ʱ��
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dblPrice As Double
    
    On Error GoTo errH
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).�շ�ϸĿID)
    End If
    
    gstrSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ��� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL)) " & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            '�Ȼ�ȡ����Ա��ǰ����ı�۽��
            If .Detail.��� Then
                If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
                    '����ҩƷʱ��(�����򲻷���)
                    '��Ȼ�м�¼(�������Ŀʱ���ж�)
                    dblAllTime = .���� * .����
                    If mblnҩ����λ And InStr(",5,6,7,", .�շ����) > 0 Then
                        dblAllTime = dblAllTime * .Detail.ҩ����װ '���ʱ�۰��ۼ��������м���
                    End If
                    If dblAllTime <> 0 Then
                        'ҩ��������ҩƷ����Ч��(����Ŀⷿһ����ҩ��)
                        gstrSQL = "Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
                            " Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0) as ʱ�� From ҩƷ���" & _
                            " Where �ⷿID=[1] And ҩƷID=[2] And ����=1 And Nvl(��������,0)>0" & _
                            " And (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
                            " Order by Nvl(����,0)"
                        Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .ִ�в���ID, .�շ�ϸĿID)
                        
                        'ʱ��=�ܽ��/������
                        dblPrice = 0 '������Ӧ�ս��
                        For i = 1 To rsPrice.RecordCount
                            If dblAllTime = 0 Then Exit For
                            'ȡС��
                            If dblAllTime <= rsPrice!��� Then
                                dblCurTime = dblAllTime
                            Else
                                dblCurTime = rsPrice!���
                            End If
                            dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!ʱ��, "0.00000"), gstrDec)
                            dblAllTime = Val(dblAllTime) - Val(dblCurTime)
                            rsPrice.MoveNext
                        Next
                        If dblAllTime <> 0 Then
                            '����δ�ֽ����
                            If InStr(",5,6,7,", .�շ����) > 0 Then
                                MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        Else
                            'ע�⣺���������ֻ�ܱ���4λС��,�Ҳ���������,������Ҫ�ֹ�����;�����������ڼ��㾫������������
                            dblAllTime = .���� * .����
                            If mblnҩ����λ And InStr(",5,6,7,", .�շ����) > 0 Then
                                dblAllTime = dblAllTime * .Detail.ҩ����װ '���ۼ���������ʵ��
                            End If
                            dblMoney = Format(dblPrice / dblAllTime, "0.00000") '�������ǰ��ۼ۵�λ
                        End If
                    Else
                        dblMoney = 0
                    End If
                Else
                    If .InComes.Count = 0 Then
                        '�����һ�μ�����,���Ĭ��ȡԭ��
                        dblMoney = 0 'IIf(IsNull(rsTmp!ԭ��), 0, rsTmp!ԭ��)
                    Else
                        dblMoney = .InComes(1).��׼����
                        '����û�����ı�۲������۷�Χ����ȡĬ��ֵ
                        If Abs(dblMoney) > Abs(IIF(IsNull(rsTmp!�ּ�), 0, rsTmp!�ּ�)) Then
                            dblMoney = IIF(IsNull(rsTmp!ԭ��), 0, rsTmp!ԭ��)
                        End If
                    End If
                End If
            End If
        End With
        
        '�����ԭ�м�¼
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        
        '��д���з��ü�¼
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .������ĿID = rsTmp!������ĿID
                .������Ŀ = rsTmp!����
                .�վݷ�Ŀ = IIF(IsNull(rsTmp!�վݷ�Ŀ), "", rsTmp!�վݷ�Ŀ)
                .ԭ�� = IIF(IsNull(rsTmp!ԭ��), 0, rsTmp!ԭ��)
                .�ּ� = IIF(IsNull(rsTmp!�ּ�), 0, rsTmp!�ּ�)
                If mobjBill.Details(lngRow).Detail.��� Then
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And mblnҩ����λ Then
                        .��׼���� = Format(dblMoney * mobjBill.Details(lngRow).Detail.ҩ����װ, "0.00000")
                    Else
                        .��׼���� = Format(dblMoney, "0.00000")
                    End If
                Else
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And mblnҩ����λ Then
                        .��׼���� = Format(Nvl(rsTmp!�ּ�, 0) * mobjBill.Details(lngRow).Detail.ҩ����װ, "0.00000")
                    Else
                        .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), "0.00000")
                    End If
                End If
                'Ӧ�ս��=���� * ���� * ����
                If mobjBill.Details(lngRow).Detail.��� And (InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 _
                        Or mobjBill.Details(lngRow).�շ���� = "4" And mobjBill.Details(lngRow).Detail.��������) Then
                    .Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                Else
                    .Ӧ�ս�� = .��׼���� * mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
                End If
                
                '�������������ü���(����������Ŀ)
                If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                    .Ӧ�ս�� = .Ӧ�ս�� * IIF(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
                End If
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100
                    .Ӧ�ս�� = .Ӧ�ս�� * (1 + dbl�Ӱ�Ӽ���)
                End If
                
                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
                
                If mbln���õǼ� Then
                    .ʵ�ս�� = 0
                Else
                    If mobjBill.Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Then
                        .ʵ�ս�� = .Ӧ�ս��
                    Else
                        .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, _
                            mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, _
                            mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����, dbl�Ӱ�Ӽ���), gstrDec))
                    End If
                End If
                
                '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����)
                    If strInfo <> "" Then
                        mobjBill.Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        mobjBill.Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
                
                'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
                mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
            End With
            rsTmp.MoveNext
        Next
    Else
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, curTotal As Currency
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    curTotal = GetBillTotal(mobjBill)
    
    If IsNumeric(cmdOK.Tag) Then
        sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '���������
    For i = 1 To Bill.Cols - 1
        '����ʱ�շ�������
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).�շ���� <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).�շ����)
    End If
    
    'ˢ�µ�����
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And mblnҩ����λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.ҩ����λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIF(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).����
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, "0.00000")
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ִ�п���"
                If mobjBill.Details(lngRow).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            Bill.TextMatrix(lngRow, i) = Get��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(lngRow, i) = Get��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Details(lngRow).�շ���� = "F" And mobjBill.Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim blnExist As Boolean
    Dim curʵ�� As Currency, curӦ�� As Currency
    Dim i As Long, j As Long, k As Long
    
    mshMoney.Redraw = False
    
    '�������ܷ�Ŀ
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '�����Ƿ��Ѿ��������������Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).������ĿID = mobjBill.Details(i).InComes(j).������ĿID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Details(i).InComes(j).ʵ�ս��
                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + mobjBill.Details(i).InComes(j).Ӧ�ս��
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
    
    'ˢ����ʾ
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).������Ŀ
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).ʵ�ս��, gstrDec)
        curʵ�� = curʵ�� + mcolMoneys(i).ʵ�ս��
        curӦ�� = curӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
    
    txtӦ��.Text = Format(curӦ��, gstrDec)
    txtʵ��.Text = Format(curʵ��, gstrDec)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    mshMoney.Redraw = True
End Sub

Private Function GetCurӦ��() As Currency
'���ܣ���ȡ���˵�ǰ���ݺϼƽ��(�շѲ����ۼӵ���ʱ��)
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCurӦ�� = GetCurӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
End Function

Private Function GetInputDetail(ByVal lng��ĿID As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select" & _
            " A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,F.Ҫ������," & _
            " Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
            " Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
            " Decode(A.���,'4',1,C." & mstrҩ����װ & ") as ҩ����װ," & _
            " Decode(A.���,'4',A.���㵥λ,C." & mstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,����֧����Ŀ F" & _
            " Where A.ID=C.ҩƷID(+) And A.ID=D.����ID(+) And B.����=A.���" & _
            " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] And A.ID=[1] And A.ID=F.�շ�ϸĿID(+) And F.����(+)=[3]"

    Else
        strSQL = _
            " Select" & _
            " A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,0 as Ҫ������," & _
            " Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
            " Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
            " Decode(A.���,'4',1,C." & mstrҩ����װ & ") as ҩ����װ," & _
            " Decode(A.���,'4',A.���㵥λ,C." & mstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E" & _
            " Where A.ID=C.ҩƷID(+) And A.ID=D.����ID(+) And B.����=A.���" & _
            " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ĿID, IIF(gbln��Ʒ��, 3, 1), lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
        .ҩ����λ = Nvl(rsTmp!ҩ����λ)
        .���� = Nvl(rsTmp!����, 0) = 1
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .���� = Nvl(rsTmp!��������)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = GetPay(lngRow)
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIF(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .���, .���㵥λ, intPay, dblTime, 0, lngDoUnit, ""
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    strSQL = "Select count(����ID) as NUM from �շѴ�����Ŀ where ����ID=" & mobjBill.Details(lngRow).�շ�ϸĿID
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
End Function

Private Function GetSubDetails(ByVal lng��ĿID As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
            
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO > 0 Then
        strSQL = _
            "Select" & _
            " A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
            " A.��������,A.����,Nvl(F.����,A.����) as ����,A.���,A.���㵥λ,A.���ηѱ�," & _
            " Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
            " Decode(A.���,'4',1,D." & mstrҩ����װ & ") as ҩ����װ,A.�������," & _
            " Decode(A.���,'4',A.���㵥λ,D." & mstrҩ����λ & ") as ҩ����λ," & _
            " A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,G.Ҫ������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,����֧����Ŀ G" & _
            " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[2] And C.����ID=[1] And A.ID=G.�շ�ϸĿID(+) And G.����(+)=[3]"
    Else
        strSQL = _
            "Select" & _
            " A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
            " A.��������,A.����,Nvl(F.����,A.����) as ����,A.���,A.���㵥λ,A.���ηѱ�," & _
            " Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
            " Decode(A.���,'4',1,D." & mstrҩ����װ & ") as ҩ����װ,A.�������," & _
            " Decode(A.���,'4',A.���㵥λ,D." & mstrҩ����λ & ") as ҩ����λ," & _
            " A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,0 as Ҫ������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F" & _
            " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[2] And C.����ID=[1]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ĿID, IIF(gbln��Ʒ��, 3, 1), lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
                .ID = rsTmp!ID
                .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
                .���� = rsTmp!����
                .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
                .��� = Nvl(rsTmp!���)
                .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
                .ҩ����λ = Nvl(rsTmp!ҩ����λ)
                .���㵥λ = Nvl(rsTmp!���㵥λ)
                .���� = Nvl(rsTmp!����, 0) = 1
                .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
                .��� = rsTmp!���
                .������� = rsTmp!�������
                .���� = rsTmp!����
                .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
                .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
                .������� = Nvl(rsTmp!�������, 0)
                .���д��� = Nvl(rsTmp!���д���, 0)
                .�������� = Nvl(rsTmp!��������, 1)
                .���� = Nvl(rsTmp!��������)
                .�������� = Nvl(rsTmp!��������, 0) = 1
                .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
                GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                    .ҩ����װ, .ҩ����λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , .Ҫ������
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).�������� > lngRow Then
            mobjBill.Details(i).�������� = mobjBill.Details(i).�������� - 1
        End If
        mobjBill.Details(i).��� = mobjBill.Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
    Set mobjBill = New ExpenseBill
    
    mcurModiMoney = 0
    mlngPreRow = 0
    cboNO.Text = ""
    chk�Ӱ�.Value = IIF(OverTime, 1, 0)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            
    'Ӥ���Ѵ���
    Call cbo��������_Click
    With mobjBill
        .�����־ = mint������Դ
        .������ = NeedName(cbo������.Text)
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .������ = UserInfo.����
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
    End With
End Sub

Private Sub ClearMoney()
'���ܣ����������ʾ��
    Dim i As Long, j As Long
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim int�к� As Integer, int��� As Integer, int�۸񸸺� As Integer
    Dim dbl���� As Double, dbl���� As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int���� As Integer, bln�ϴ� As Boolean
    Dim strSQL As String, strStuffDept As String '��¼���Ϸ��ϲ���
    
    If mint��¼���� = 1 Then
        mobjBill.NO = zlDatabase.GetNextNo(13)
    Else
        mobjBill.NO = zlDatabase.GetNextNo(14)
    End If
    mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
    
    int��� = 0
    arrSQL = Array()

    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                
                '��������
                With mobjBill
                    If mint������Դ = 2 Then
                        gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                            ZVal(.��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                            ZVal(.����ID) & "," & ZVal(.����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                    Else
                        If mint��¼���� = 2 Then
                            gstrSQL = "zl_������ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & _
                                ZVal(.��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "'," & _
                                "'" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & "," & _
                                ZVal(.����ID) & "," & ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
                        Else
                            gstrSQL = "zl_���ﻮ�ۼ�¼_Insert('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                                ZVal(.��ʶ��) & ",'" & IIF(Val(txt���ʽ.Tag) = 0, "", txt���ʽ.Tag) & "','" & .���� & "'," & _
                                "'" & .�Ա� & "','" & .���� & "','" & .�ѱ� & "'," & .�Ӱ��־ & "," & _
                                ZVal(.����ID) & "," & ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
                        End If
                    End If
                End With
                
                '�շ�ϸĿ����
                With mobjBillDetail
                    '�����������
                    If .��� <> int�к� Then
                        int�к� = .���
                        int�۸񸸺� = int���
                        
                        '���´����������
                        For i = .��� + 1 To mobjBill.Details.Count
                            If mobjBill.Details(i).�������� = .��� Then
                                mobjBill.Details(i).�������� = int���
                            End If
                        Next
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                    
                    If mint������Դ = 2 Then
                        gstrSQL = gstrSQL & IIF(.������Ŀ��, 1, 0) & "," & ZVal(.���մ���ID) & ",'" & .���ձ��� & "',"
                    ElseIf mint��¼���� = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl���� = .����
                    If InStr(",5,6,7,", .�շ����) > 0 And mblnҩ����λ Then
                        dbl���� = Format(.���� * .Detail.ҩ����װ, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIF(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & ZVal(.ִ�в���ID) & ","
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    dbl���� = .��׼����
                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And mblnҩ����λ Then
                        dbl���� = Format(.��׼���� / mobjBillDetail.Detail.ҩ����װ, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIF(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & .������ĿID & "," & _
                        "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                    If mint������Դ = 2 Then
                        gstrSQL = gstrSQL & IIF(.ͳ���� = 0, "NULL", .ͳ����) & ","
                    End If
                End With
                                                
                '��������
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                
                '�Ƿ�ֻ���ɻ��۵�
    '                If mbln���õǼ� Then
    '                    int���� = 0 '��ķ��õǼǲ������ɻ��۵�
    '                Else
    '                    '����Ӧ���������,Ӧ��ִ��ҽ����������ж�
    '                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 Then
    '                        int���� = IIF(InStr(gstr���ͻ��۵�, "5") > 0, 1, 0)
    '                    Else
    '                        int���� = IIF(InStr(gstr���ͻ��۵�, mobjBillDetail.�շ����) > 0, 1, 0)
    '                    End If
    '                End If
                If int���� = 0 Then bln�ϴ� = True 'ֻҪ���ڲ��ǻ��۵���Ҫ�ϴ�
                
                '�ռ����Ϸ��ϲ���,�Ա��Զ�����,���ﲡ�˽�����ʱ(����Ϊ����ʱ����),סԺ����ֻ�м���
                With mobjBillDetail
                    If (mint������Դ = 1 And mint��¼���� = 2 And gbln�����Զ����� Or mint������Դ = 2 And gblnסԺ�Զ�����) And int���� = 0 Then
                        If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� Then
                            If InStr("," & strStuffDept, "," & .ִ�в���ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .ִ�в���ID
                            End If
                        End If
                    End If
                End With
                
                If mint������Դ = 2 Then
                    gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                        "0," & IIF(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                        "NULL,'" & mobjBillDetail.ժҪ & "'," & chk����.Value & "," & ZVal(mlngҽ��ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.���� & "')"
                Else
                    If mint��¼���� = 2 Then
                        gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                            IIF(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                            "NULL,'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ")"
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.���� & "'," & _
                            IIF(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                            "'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    '����ҽ��Ժ�ӷ���
    gstrSQL = "ZL_����ҽ������_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mint��¼���� & ",'" & mobjBill.NO & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
    
    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        '���ж��Ƿ�ҽ�����˼ǵ���,�����Ϸ��Լ��(�����޸�ʱ������һ������ж�)
        If mint������Դ = 2 Then
            'ȥ����ҽ������ƥ����
            intInsure = BillExistInsure(mstrInNO)
        End If
        If mint������Դ = 2 Then
            gstrSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            If mint��¼���� = 2 Then
                gstrSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Else
                gstrSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & mstrInNO & "')"
            End If
        End If
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If
    
    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        strTmp = ""
        On Error GoTo errH
        gcnOracle.BeginTrans
        
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            'ִ���Զ�����
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                For i = 0 To UBound(Split(strStuffDept, ","))
                    strSQL = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & ",25,'" & mobjBill.NO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Next
            End If
            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�
            If mint������Դ = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If gclsInsure.GetCapability(support�����ϴ�, , mrsInfo!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsInfo!����) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans
        
        'ҽ���ӿ�
        '1.ҽ�����������ϴ�
        If mint������Դ = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ҽ������ʧ��,�õ��ݵķ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.����ʵʱ�ϴ�
        If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
            'ҽ�����������ϸ
            If gclsInsure.GetCapability(support�����ϴ�, , mrsInfo!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsInfo!����) Then
                strTmp = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!����) Then
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        cboNO.AddItem mobjBill.NO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i 'ֻ��ʾ10��
        Next
        
        'ҽ���ӿ�
        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Integer
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strNO=���ݺ�
'      blnDelete=�Ƿ��ȡҪ�˷ѵĵ���
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String, intSign As Integer
    Dim curTotal As Currency, curӦ��Total As Currency
    Dim strSQL As String, i As Long
    Dim intInsure As Integer, blnDo As Boolean
    Dim blnNOMoved As Boolean
    
    If mbytInState = 1 Then
        blnNOMoved = MovedByNO(strNO, "���˷��ü�¼", "��¼����=" & mint��¼���� & IIF(mstrTime <> "", " And �Ǽ�ʱ��=To_Date('" & mstrTime & "','YYYY-MM-DD HH24:MI:SS')", ""))
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    strSQL = _
        " Select A.����ID,A.��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.����," & _
        " A.���˲���ID,A.��������ID,A.�Ӱ��־,A.Ӥ����,A.������,A.������,A.����Ա����," & _
        " A.��������ID,C.����||'-'||C.���� as ��������,A.����ʱ��," & _
        " B.ҽ�Ƹ��ʽ,B.������,B.������,A.�Ƿ���" & _
        " From ���˷��ü�¼ A,������Ϣ B,���ű� C" & _
        " Where Rownum=1 And NO=[1] And A.��¼����=[2]" & _
        " And A.����ID=B.����ID And Instr([3],A.��¼״̬)>0" & _
        IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
        " And A.��������ID=C.ID"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ��ݡ�", vbInformation, gstrSysName
        Exit Function
    End If

    cboNO.Text = strNO
    txt����.Text = Nvl(rsTmp!����)
    txt�Ա�.Text = Nvl(rsTmp!�Ա�)
    txt����.Text = Nvl(rsTmp!����)
    If Nvl(rsTmp!��ҳID, 0) <> 0 Then
        txt����.Text = Nvl(rsTmp!����)
    End If
    txt�ѱ�.Text = Nvl(rsTmp!�ѱ�)
    txt������.Text = Nvl(rsTmp!������)
    txt������.Text = Format(Nvl(rsTmp!������), "0.00")
    txt���ʽ.Text = Nvl(rsTmp!ҽ�Ƹ��ʽ)
    
    cbo��������.AddItem Nvl(rsTmp!��������)
    cbo��������.ItemData(cbo��������.NewIndex) = Nvl(rsTmp!��������ID, 0)
    cbo��������.ListIndex = cbo��������.NewIndex
    
    If Nvl(rsTmp!�Ƿ���, 0) = 1 Then
        chk����.Value = 1: chk����.Visible = True
    End If
    
    chk�Ӱ�.Value = Nvl(rsTmp!�Ӱ��־, 0)
    cboBaby.ListIndex = IIF(Val("" & rsTmp!Ӥ����) > cboBaby.ListCount - 1, 0, Val("" & rsTmp!Ӥ����))
    
    '������
    Call GetCboIndex(cbo������, Nvl(rsTmp!������))
    If cbo������.ListIndex = -1 And Not IsNull(rsTmp!������) Then
        cbo������.AddItem rsTmp!������
        cbo������.ListIndex = cbo������.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    If mint��¼���� = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID, , True)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "Ԥ��:" & Format(rsPatiMoney!Ԥ�����, "0.00") & _
                "/����:" & Format(rsPatiMoney!�������, gstrDec) & _
                "/ʣ��:" & Format(rsPatiMoney!Ԥ����� - rsPatiMoney!�������, "0.00")
        End If
    End If
    
    '------------------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        
        '��ȡ������ԭʼ��¼�ķ���ID
        strSQL1 = _
            " Select A.ID,A.���,A.�շ�ϸĿID," & _
            " Nvl(A.����,1)*A.����" & IIF(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & " as ԭʼ����" & _
            " From ���˷��ü�¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.��¼����=[2]"
        
        '��ȡҩƷ�շ���¼�е�׼����
        strSQL2 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIF(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.����� is NULL" & _
            " And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        strSQL = "Select Nvl(�۸񸸺�,���) From ���˷��ü�¼" & _
            " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1"
        
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        If mint��¼���� = 2 Then
            If mint������Դ = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure)
            Else
                blnDo = gbytBillOpt = 2
            End If
            If blnDo Then
                strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                    " (" & _
                    " Select Nvl(�۸񸸺�,���) as ���" & _
                    " From ���˷��ü�¼" & _
                    " Where NO=[1] And ��¼���� IN(2,12)" & _
                    " Group by Nvl(�۸񸸺�,���)" & _
                    " Having Sum(Nvl(���ʽ��,0))=0" & _
                    " )"
            End If
        End If
        
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIF(mblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & mstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIF(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIF(mblnҩ����λ, "*Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From ���˷��ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=[2]" & _
            " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),C.����,C.����,A.�շ�ϸĿID,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X." & mstrҩ����λ & ",X." & mstrҩ����װ
            
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = _
            " Select A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������,A.���㵥λ," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            " Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            " A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.���=B.��� And B.ID=C.����ID(+)" & _
            " Group by A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������," & _
            " A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���,A.���ӱ�־" & _
            " Having Sum(A.����*A.����)<>0"
            
        strSQL = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,A.���," & _
            " A.��������,A.���㵥λ,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
            " A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
            " A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
            " A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIF(mblnDelete, -1, 1) '����,�����������
        
        strSQL = _
            "Select A.�շ�ϸĿID,A.�շ����,A.ִ�в���ID,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " A.���㵥λ,A.����,A.����,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��,A.���ӱ�־,A.��������" & _
            " From ���˷��ü�¼ A Where A.��¼����=[2]" & _
            " And Instr([4],A.��¼״̬)>0 And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        End If
        
        strSQL = _
            " Select A.���,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIF(mblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & mstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg([7]*A.����" & IIF(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIF(mblnҩ����λ, "*Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum([7]*A.Ӧ�ս��) as Ӧ�ս��,Sum([7]*A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ����" & _
            " And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            " Group by A.���,C.����,C.����,A.�շ�ϸĿID,B.����,B.���," & _
            " Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X." & mstrҩ����λ
            
        strSQL = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,A.���,A.��������," & _
            " A.���㵥λ,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            " Order by ���"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, IIF(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), IIF(gbln��Ʒ��, 3, 1), intSign)
    If rsTmp.EOF Then
        If blnDelete Then
            MsgBox "�����е�ǰ�޿��Բ����ļ�¼�����ܵ����е���Ŀ�Ѿ�ȫ��ִ�С�", vbInformation, gstrSysName
        Else
            MsgBox "�����е�ǰ�޿��Բ����ļ�¼��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        Bill.RowData(i) = rsTmp!��� '���ڼ�������
        
        Bill.TextMatrix(i, 1) = rsTmp!���
        Bill.TextMatrix(i, 2) = rsTmp!����
        Bill.TextMatrix(i, 3) = Nvl(rsTmp!���)
        Bill.TextMatrix(i, 4) = Nvl(rsTmp!���㵥λ)
        Bill.TextMatrix(i, 5) = Nvl(rsTmp!����)
        Bill.TextMatrix(i, 6) = FormatEx(rsTmp!����, 5)
        Bill.TextMatrix(i, 7) = Format(rsTmp!����, "0.00000")
        Bill.TextMatrix(i, 8) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, 9) = Format(rsTmp!ʵ�ս��, gstrDec)
        Bill.TextMatrix(i, 10) = Nvl(rsTmp!ִ�в���)
        Bill.TextMatrix(i, 11) = IIF(rsTmp!���ӱ�־ = 1, "��", "")
        Bill.TextMatrix(i, 12) = Nvl(rsTmp!��������)
        
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "��"
        End If
        
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 2, &HE7CFBA
    Bill.SetColColor 6, &HE7CFBA
    Bill.SetColColor 10, &HE7CFBA
    Bill.SetColColor 5, &HE0E0E0
    Bill.SetColColor 7, &HE0E0E0
    Bill.SetColColor 11, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))

        '��ȡҩƷ�շ���¼�е�׼����
        strSQL1 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIF(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.����� is NULL" & _
            " And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "Select Nvl(�۸񸸺�,���) From ���˷��ü�¼" & _
            " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1"
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From ���˷��ü�¼" & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
        
        strSQL = _
            " Select Sum(A.ID) as ID,A.���,A.����,A.�շ����," & _
            " Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
            " Sum(A.ʵ�ս��) as ʣ��ʵ�� From (" & _
            " Select Decode(A.��¼״̬,2,0,A.ID) as ID,A.���,B.����,A.�շ����," & _
            " Nvl(A.����,1)*A.����" & IIF(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & " as ����," & _
            " A.Ӧ�ս��,A.ʵ�ս��" & _
            " From ���˷��ü�¼ A,������Ŀ B,ҩƷ��� X" & _
            " Where A.��¼����=[2] And A.NO=[1]" & _
            " And A.������ĿID=B.ID And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+)) A" & _
            " Group by A.���,A.����,A.�շ����" & _
            " Having Sum(A.����)<>0"
                    
        '��������
        strSQL = _
            " Select A.����,Sum(A.ʣ��Ӧ��*(A.׼������/A.ʣ������)) as Ӧ�ս��," & _
            " Sum(ʣ��ʵ��*(A.׼������/A.ʣ������)) as ʵ�ս�� From (" & _
            " Select A.����,A.ʣ������,A.ʣ��Ӧ��,A.ʣ��ʵ��," & _
            " Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������)) as ׼������" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
            " Where A.ID=B.����ID(+)" & _
            " ) A Group by A.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIF(mblnDelete, -1, 1) '����,�����������
        
        strSQL = "Select A.������ĿID,A.Ӧ�ս��,A.ʵ�ս�� From ���˷��ü�¼ A" & _
            " Where Instr([4],A.��¼״̬)>0 And A.��¼����=[2] And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        End If
        
        strSQL = _
            " Select B.����,Sum([6]*A.Ӧ�ս��) as Ӧ�ս��,Sum([6]*A.ʵ�ս��) as ʵ�ս�� " & _
            " From (" & strSQL & ") A,������Ŀ B Where A.������ĿID=B.ID Group By B.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, IIF(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), intSign)
    If rsTmp.EOF Then Exit Function
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    Call SetMoneyList
    
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!����
        mshMoney.TextMatrix(i, 1) = Format(rsTmp!ʵ�ս��, gstrDec)
        curTotal = curTotal + rsTmp!ʵ�ս��
        curӦ��Total = curӦ��Total + rsTmp!Ӧ�ս��
        rsTmp.MoveNext
    Next
    
    txtʵ��.Text = Format(curTotal, gstrDec)
    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'���ܣ������еĿ���(���ʱչ��)
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(5) = 0
    ElseIf Bill.ColWidth(5) = 0 Then
        Bill.ColWidth(5) = 520
    End If
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    Dim i As Long
    'ȡ������ҩ�ĸ���
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).����
            Exit For
        End If
    Next
End Function

Private Function GetDetailNum(lngRow As Long) As Double
'���ܣ���ȡ����ָ��ϸĿ���ܼ�������(����������)
'������lngRow=��ǰ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngNum As Long, i As Long
    
    If lngRow <= mobjBill.Details.Count Then
        '��ǰ�����е�����
        For i = 1 To mobjBill.Details.Count
            If i <> lngRow And mobjBill.Details(i).�շ�ϸĿID = mobjBill.Details(lngRow).�շ�ϸĿID Then
                lngNum = lngNum + mobjBill.Details(i).���� * IIF(mobjBill.Details(i).���� = 0, 1, mobjBill.Details(i).����)
            End If
        Next
        '���ݿ��е�����
        strSQL = _
            "Select Sum(A.����*Nvl(A.����,1)" & IIF(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & ") as NUM" & _
            " From ���˷��ü�¼ A,ҩƷ��� B" & _
            " Where A.�۸񸸺� is Null And A.��¼״̬<>0 And A.���ʷ���=1" & _
            " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2]" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.�շ�ϸĿID+0=[3]"
            
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID), Val(Nvl(mrsInfo!��ҳID, 0)), mobjBill.Details(lngRow).�շ�ϸĿID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + Nvl(rsTmp!Num, 0)
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSQL As String, bytDay As Byte
    Dim strҩ�� As String, lng��������ID As Long
    
    lng��������ID = mrsInfo!����ID    '������������
    If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    
    If str��� = "4" Then
        strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([1],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[1])" & _
            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
            " And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
            
        '�Լ�SQL�����Ĳ�֧�ִ洢�ⷿ����֮ǰ��
'        strSQL = "Select A.ID,A.����,A.����,A.����,B.��������,B.�������" & _
'            " From ���ű� A,��������˵�� B" & _
'            " Where A.ID=B.����ID And B.��������='���ϲ���' And B.������� IN([1],3)" & _
'            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
'            " Order by B.�������,A.����"
    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not Check�ϰల��(True) Then
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
                " And B.������� IN([1],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[1])" & _
                " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                " And A.�շ�ϸĿID=[3]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
                " And B.������� IN([1],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[1])" & _
                " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                " And A.�շ�ϸĿID=[3]" & _
                " Order by B.�������,C.����"
        End If
    End If
    
    On Error GoTo errH
    'Set mrsWork = New ADODB.Recordset
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint������Դ, lng��������ID, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load������(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngOldID As Long
    
    cbo������.Clear
    
    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��','��ʿ') And B.����ID=[1]"
    '��Ϊ����ҽ��
    If lng����ID = mlng��������ID And mlng��������ID <> mlng��������ID Then
        strSQL = strSQL & " And A.����=[2]"
    End If
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mstr����ҽ��)
    
    i = IIF(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo������.AddItem IIF(IsNull(rsTmp!����), "", rsTmp!���� & "-") & rsTmp!����
                cbo������.ItemData(cbo������.ListCount - 1) = rsTmp!����ID
                marrDr(i - 1) = rsTmp!ID & "|" & rsTmp!����ID & "|" & Nvl(rsTmp!���) & "|" & rsTmp!���� & "|" & Nvl(rsTmp!����) & "|" & rsTmp!ְ�� & "|" & Nvl(rsTmp!��Ա����)
                
                If rsTmp!ID = UserInfo.ID And cbo������.ListIndex = -1 Then cbo������.ListIndex = cbo������.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo������.ListCount > 0 Then ReDim Preserve marrDr(cbo������.ListCount - 1)
        
        If cbo������.ListCount = 1 And cbo������.ListIndex = -1 Then cbo������.ListIndex = 0
    End If
End Sub

Private Function CalcGridToTal(Optional blnӦ�� As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
    Else
        For i = 1 To Bill.Cols - 1
            If blnӦ�� Then
                If Bill.TextMatrix(0, i) = "Ӧ�ս��" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "ʵ�ս��" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "ɾ��" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(1) = GetOrigColWidth(1) - 120
            Bill.ColWidth(2) = GetOrigColWidth(2) - 100
            Bill.ColWidth(10) = GetOrigColWidth(10) - 200
            
            Bill.ColWidth(7) = GetOrigColWidth(7) - 50
            Bill.ColWidth(8) = GetOrigColWidth(8) - 50
            Bill.ColWidth(9) = GetOrigColWidth(9) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(1) = GetOrigColWidth(1)
            Bill.ColWidth(2) = GetOrigColWidth(2)
            Bill.ColWidth(10) = GetOrigColWidth(10)
            
            Bill.ColWidth(7) = GetOrigColWidth(7)
            Bill.ColWidth(8) = GetOrigColWidth(8)
            Bill.ColWidth(9) = GetOrigColWidth(9)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'���ܣ�������ʾ���е��к�
'������intRow=�Ӹ��п�ʼ
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, 0) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True) As Integer
'���ܣ����ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
'������tmpDetail=�������Ŀ,����Ϊ������,blnCommon=�Ƿ��������ж�,����Ϊҽ���򹫷Ѳ��˵��ж�
'���أ���ƥ�����,0Ϊ��ȷ
'˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    Dim i As Long, intְ��A As Integer, intְ��B As Integer
    Dim strTmp As String
    
    strTmp = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    
    If cbo������.ListIndex = -1 Then Exit Function
    If cbo������.ListIndex <= UBound(marrDr) Then
        If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 5 Then
            intְ��A = Val(Split(marrDr(cbo������.ListIndex), "|")(5))
        End If
    End If
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                If Not blnCommon Then
                    intְ��B = Val(Right(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    intְ��B = Val(Left(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        If Not blnCommon Then
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
'���ܣ��ж�ָ��ҩƷ�ڵ������Ƿ��Ѿ�����
'������objDetail=��Ŀ,intRow=Ҫ�жϵ���
'˵����ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ)
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    
    Check�������� = True
    
    '�޷����
    If txt���ʽ.Tag = "" Then Exit Function
    
    'ȷ����������
    bytType = Val(txt���ʽ.Tag)
    
    'ֻ���ҽ�����˺͹��Ѳ���
    If bytType <> 1 And bytType <> 2 Then Exit Function
    
    '��ȡ�������
    If bytType = 1 Then
        strSQL = "Select * From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
    Else
        strSQL = "Select * From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
    End If
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            rsTmp.Filter = "����='" & mobjBill.Details(intRow).Detail.���� & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Details(intRow).Detail.���� & """,����" & _
                    IIF(bytType = 1, "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.���� = "" Then
                If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check�������� = False: Exit For
                End If
            Else
                rsTmp.Filter = "����='" & mobjBill.Details(i).Detail.���� & "'"
                If rsTmp.EOF Then
                    If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                        mobjBill.Details(i).Detail.���� & """,����" & _
                        IIF(bytType = 1, "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!����) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����)
                If strInfo <> "" Then
                    mobjBill.Details(i).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(i).���մ���ID = Val(Split(strInfo, ";")(1))
                    mobjBill.Details(i).InComes(j).ͳ���� = Val(Split(strInfo, ";")(2))
                    mobjBill.Details(i).���ձ��� = CStr(Split(strInfo, ";")(3))
                    
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).ժҪ = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.���� = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            Next
        Next
    End If
End Sub

Private Function HaveStopClass() As Integer
'���ܣ��жϵ�ǰ�������Ƿ��л�ʿ��ֹ���������
    Dim i As Long, str���� As String
    
    If cbo������.ListIndex <> -1 Then
        If cbo������.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                str���� = Split(marrDr(cbo������.ListIndex), "|")(6)
            End If
        End If
    End If
    
    For i = 1 To mobjBill.Details.Count
        If str���� = "��ʿ" And InStr(",E,M,4,", mobjBill.Details(i).�շ����) = 0 Then
            HaveStopClass = i: Exit Function
        End If
    Next
End Function

Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, 10) = "" Then
            Checkִ�п��� = i: Exit Function
        End If
    Next
End Function

Public Sub InitLocPar()
'���ܣ���ʼ�����ñ�������
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mblnPay = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ҩ����", 1)) <> 0
    mblnTime = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "�������", 0)) <> 0
    mbln����ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ�����", 0)) = 1
    mbln����ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ����", 0)) = 1
    mstr�շ���� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�շ����", "")
    
    'ҩƷ��λ
    mblnҩ����λ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҩƷ��λ", 0)) <> 0
    If mint������Դ = 1 Then
        mstrҩ����λ = "���ﵥλ": mstrҩ����װ = "�����װ"
    Else
        mstrҩ����λ = "סԺ��λ": mstrҩ����װ = "סԺ��װ"
    End If
    
    'ȱʡҩ��
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
End Sub
