VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimpleBilling 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "סԺ�򵥼���"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimpleBilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6285
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSimpleBilling.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleBilling.frx":115E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleBilling.frx":1798
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
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
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   10365
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4380
      Width           =   10365
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
         Height          =   585
         Left            =   0
         TabIndex        =   34
         ToolTipText     =   "���:F6"
         Top             =   -105
         Width           =   10290
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   1680
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   1155
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   4815
            TabIndex        =   14
            Top             =   165
            Width           =   1785
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   7770
            TabIndex        =   15
            Top             =   165
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "Ӥ����"
            Height          =   240
            Left            =   1320
            TabIndex        =   12
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   7215
            TabIndex        =   36
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   240
            Left            =   4020
            TabIndex        =   35
            Top             =   225
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   7230
         TabIndex        =   19
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   7230
         TabIndex        =   18
         ToolTipText     =   "�ȼ���F2"
         Top             =   675
         Width           =   1500
      End
      Begin VB.Frame fraMoney 
         Height          =   1545
         Left            =   0
         TabIndex        =   37
         Top             =   360
         Width           =   3195
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
            Height          =   1335
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   165
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   4
            FixedCols       =   0
            RowHeightMin    =   300
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            FormatString    =   "^��Ŀ       |^���        "
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
      Begin VB.Frame fraStat 
         Height          =   1545
         Left            =   3195
         TabIndex        =   38
         Top             =   360
         Width           =   3405
         Begin VB.TextBox txtʵ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   870
            Width           =   1845
         End
         Begin VB.TextBox txtӦ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   345
            Width           =   1845
         End
         Begin VB.Label lblʵ�� 
            AutoSize        =   -1  'True
            Caption         =   "ʵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   300
            Left            =   450
            TabIndex        =   40
            Top             =   945
            Width           =   630
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   450
            TabIndex        =   39
            Top             =   420
            Width           =   630
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   645
      Left            =   30
      TabIndex        =   23
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   10275
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   210
         Width           =   1380
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   9690
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   9720
         TabIndex        =   31
         Top             =   195
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "סԺ���ʵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   75
         TabIndex        =   26
         ToolTipText     =   "���:F6"
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7365
         TabIndex        =   24
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1050
      Left            =   30
      TabIndex        =   22
      Top             =   405
      Width           =   10275
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   615
         TabIndex        =   45
         Top             =   195
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmSimpleBilling.frx":1DD2
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
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         MustSelectItems =   "����"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   840
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   8415
         TabIndex        =   9
         Text            =   "cbo��������"
         Top             =   615
         Width           =   1755
      End
      Begin VB.TextBox txtҽ�Ƹ��� 
         Height          =   360
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   1680
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
         Width           =   525
      End
      Begin VB.TextBox txt�ѱ� 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   1740
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1290
         MaxLength       =   64
         TabIndex        =   1
         Top             =   195
         Width           =   1680
      End
      Begin VB.TextBox txtOld 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   5295
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   5340
         TabIndex        =   44
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   3285
         TabIndex        =   43
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   7395
         TabIndex        =   42
         Top             =   675
         Width           =   960
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         Caption         =   "���ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   41
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   6195
         TabIndex        =   32
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   30
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Index           =   8
         Left            =   3495
         TabIndex        =   29
         Top             =   255
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Index           =   9
         Left            =   4770
         TabIndex        =   28
         Top             =   255
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���˷ѱ�"
         Height          =   240
         Index           =   12
         Left            =   7395
         TabIndex        =   27
         Top             =   255
         Width           =   960
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2895
      Left            =   30
      TabIndex        =   10
      Top             =   1455
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5106
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   8445
      Top             =   5205
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
            Picture         =   "frmSimpleBilling.frx":1DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSimpleBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
'2.����ʼ״̬������
Public mbytInState As Byte '0-ִ��,1-����,2-����,3-����
Public mstrInNO As String '��mbytInState=1ʱ��Ч,���ڵ��ݺ�
Public mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���,����ʱ����
 
Public mstrTime As String '�����������ݵĵǼ�ʱ��
Public mblnDelete As Boolean '�Ƿ�����˷ѵ���

Public mlngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
Public mlngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
Public mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
Public mlng����ID As Long '���ҷ�ɢ������
Public mstrPrivs As String
Public mlngModule As Long
Private mobjICCard As Object
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mcurModiMoney As Currency '�޸ĵ���ʱԭ���ݵĽ��
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID
'����������������������������������������������������������������������������������������������������������������������������������������
Private mblnNotCick As Boolean
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
    ��Ŀ = 0
    Ӧ�ս�� = 1
    ʵ�ս�� = 2
    ִ�п��� = 3
    ���� = 4
End Enum

'���ݶ���
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsMedPayMode As ADODB.Recordset '���п��õ�ҽ�Ƹ��ʽ

'�������
Private mobjBill As ExpenseBill '������õ��ݶ������
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes

'�������
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mrsWarn As ADODB.Recordset '����������
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '��ѡҽ���ͻ�ʿ

Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnValid As Boolean
Private mblnPrint As Boolean '��ȡ��˵�ʱ�Ƿ����Ҫ��ӡ���շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����

Private Const STR_HEAD = "��Ŀ,3000,1;Ӧ�ս��,1500,7;ʵ�ս��,1500,7;ִ�п���,1900,1;����,870,1"
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Private Sub Bill_cboClick(ListIndex As Long)
    Dim lngִ�п��� As Long, strִ�п��� As String
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    lngִ�п��� = .ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.cboObj.Text = strִ�п���
                        .ִ�в���ID = lngִ�п���: Exit Sub
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "ִ�п���"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "��ҩҩ��"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If mrsUnit Is Nothing Then Exit Sub
        If mrsUnit.State <> 1 Then Exit Sub
        If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
    End With
    Exit Sub
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    
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
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        '���¼��㲢ˢ��
        'Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ���������
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean
    Dim str��׼��Ŀ As String, int������Դ As Integer, int���� As Integer
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            int���� = mrsInfo!����
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
        End If
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            int������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            int������Դ = 1
        End If
    Else
        int������Դ = 2
    End If
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, False, "'Z'", , , str��׼��Ŀ, _
        , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim objDetail As Detail, lng��Ŀid As Long
    Dim str��׼��Ŀ As String, int������Դ As Integer, int���� As Integer
    Dim cur�ϼ� As Currency, strScope As String
    Dim dblPreMoney As Double, i As Long, lngDoUnit As Long
    Dim cur��� As Currency, curItemMoney As Currency
    
    On Error GoTo errH
    
    If Bill.ColData(Bill.Col) = 0 Then Exit Sub
    
    If KeyCode = 13 Then
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set objDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!����) Then
                                int���� = mrsInfo!����
                                '���˺�:24862
                                If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
                                
                            End If
                            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                                int������Դ = 2
                            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                                int������Դ = 1
                            End If
                        Else
                            int������Դ = 2
                        End If
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, False, "'Z'", Bill.Text, _
                            Bill.TxtHwnd, , , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        If lng��Ŀid <> 0 Then
                            Set objDetail = GetInputDetail(lng��Ŀid)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(���Ӳ���)
                    
                    'ҽ�����˷�������
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) Then
                            If objDetail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "��ĿID=" & objDetail.ID
                                If mrsMedAudit.RecordCount = 0 Then
                                    MsgBox "��ǰ����δ����׼ʹ��[" & objDetail.���� & "]��", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                            
                            'ҽ������
                            If Not CheckMediCareItem(objDetail.ID, mrsInfo!����, objDetail.����, _
                                objDetail.��� = False, , mstr��ͨ�۸�ȼ�) Then
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(objDetail, Bill.Row)
                    Call CalcMoneys(Bill.Row)
                    
                    '����ժҪ(������������и���ժҪ)
                    Dim strժҪ As String '90304
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    Else
                         strժҪ = gclsInsure.GetItemInfo(0, mobjBill.����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 1)
                         mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    mrsWarn.Filter = ""
                    If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        cur�ϼ� = GetBillTotal(mobjBill)
                        If cur�ϼ� > 0 Then
                            cur��� = Val(txtʵ��.Tag)
                            '���˺�:24491
                            curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                            
                            If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                            gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, cur�ϼ�, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, , , curItemMoney)
                            If gbytWarn = 2 Or gbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                        mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    '��һ�е�����ȷ��
                    If mobjBill.Details(Bill.Row).Detail.��� Then Bill.ColData(1) = 4 'Ӧ�ս��
                    
                    'ִ�п���!!!
                    Call FillBillComboBox(Bill.Row, 3)
                    If Bill.ListCount = 1 Then
                        Bill.ColData(3) = 5
                        mobjBill.Details(Bill.Row).Key = 1
                    Else
                        Bill.ColData(3) = 3
                        mobjBill.Details(Bill.Row).Key = Bill.ListCount
                    End If
                    
                    '������Ŀ����(��������Դ���༶����-�����Ĵ���...)
                    If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" Then
                        If ShouldDO(Bill.Row) Then
                            Set mcolDetails = New Details
                            Set mcolDetails = GetSubDetails(mobjBill.Details(Bill.Row).�շ�ϸĿID)
                            For i = 1 To mcolDetails.Count
                                If mobjBill.Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 0) = "" '�б�Ҫ����
                                
                                If mcolDetails(i).��� = mobjBill.Details(Bill.Row).�շ���� Then
                                    '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                                    lngDoUnit = mobjBill.Details(Bill.Row).ִ�в���ID
                                Else
                                    If mcolDetails(i).ִ�п��� = 0 Then
                                        '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                        lngDoUnit = mobjBill.Details(Bill.Row).ִ�в���ID
                                    End If
                                        '�������,ȡ�������õ�ִ�п���
                                End If
                                
                                Call SetDetail(mcolDetails(i), Bill.Rows - 1, Bill.Row, lngDoUnit)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            Next
                        End If
                    End If
                End If
            Case "Ӧ�ս��" 'ʵ�����ǵ���(��Ϊ���ݴ�ȱʡΪ1,�Ҳ��ܸ���)
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '��ֵ�Ϸ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '����Ȩ��
                    If CDbl(Bill.Text) < 0 Then
                        If InStr(mstrPrivsOpt, ";���Ƹ�������;") = 0 Then
                            MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                            Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                        Else
                            If mrsInfo.State = 1 Then
                                If Not IsNull(mrsInfo!����) Then
                                    If Not gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����) Then
                                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    Bill.Text = Format(Bill.Text, gstrDec)
                    
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
                        
                        '�����Ŀֻ��Ӧһ��������Ŀ
                        dblPreMoney = mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).InComes(1).��׼����
                        mobjBill.Details(Bill.Row).���� = Sgn(Val(Bill.Text))
                        mobjBill.Details(Bill.Row).InComes(1).��׼���� = Abs(Val(Bill.Text))
                        Call CalcMoneys(Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 Then
                            cur�ϼ� = GetBillTotal(mobjBill)
                            If cur�ϼ� > 0 Then
                                cur��� = Val(txtʵ��.Tag)
                                '���˺�:24491
                                curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                                If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, cur�ϼ�, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, , , curItemMoney)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).���� = Sgn(dblPreMoney)
                                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = Abs(dblPreMoney)
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
                If mobjBill.Details.Count >= Bill.Row Then
                    If Bill.ListIndex <> -1 Then
                        'If mobjBill.Details(Bill.Row).ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                            mobjBill.Details(Bill.Row).ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                            If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row)
                        'End If
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub SetSubDept(ByVal lngRow As Long)
    Dim i As Long, j As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
            If InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) = 0 Then
                With mobjBill
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID) '������ȡ
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
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mcolDetails(j).���, _
                                    mcolDetails(j).ID, mcolDetails(j).ִ�п���, .Details(i).ִ�в���ID, Get��������ID, Get������Դ, , mobjBill.����ID)
                            End If
                        End If
                    End If
                    
                    If .Details(i).ִ�в���ID > 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, 3) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, 3) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, 3) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    End If
                End With
            End If
        End If
    Next
    
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
    Dim i As Long
    
    If Not Bill.Active Then Exit Sub
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '����б༭����������ɫ
        Bill.SetColColor 0, &HE7CFBA '��ȻҪ�ɰ�ɫ
        Exit Sub
    End If
    
    If mbytInState = 0 Then
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
    End If

    '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
    If mobjBill.Details.Count >= Row Then
    If ItemHaveSub(Row) Or mobjBill.Details(Row).�������� > 0 Then
        Bill.ColData(0) = BillColType.Text_UnModify
    End If
    End If
    
    If mobjBill.Details.Count >= Bill.Row And mbytInState <> 2 Then
        If mobjBill.Details(Bill.Row).Key = "1" Then
            Bill.ColData(3) = 5
        Else
            Bill.ColData(3) = 3
        End If
    End If
    If Bill.ColData(Bill.Col) = 3 Then Call FillBillComboBox(Bill.Row, Bill.Col)
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "ִ�п���"
            SetWidth Bill.cboHwnd, 130
        Case "Ӧ�ս��"
            Bill.TextLen = 10
            If InStr(mstrPrivsOpt, ";���Ƹ�������;") = 0 Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
            
            If InStr(Bill.TextMask, "-") > 0 And mrsInfo.State = 1 Then
                If Not IsNull(mrsInfo!����) Then
                    If Not gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����) Then
                        Bill.TextMask = Replace(Bill.TextMask, "-", "")
                    End If
                End If
            End If
    End Select

    '������ʱ,�������ø��еı༭����
    If mobjBill.Details.Count >= Bill.Row Then
        If mobjBill.Details(Bill.Row).Detail.��� Then
            Bill.ColData(1) = 4
        Else
            Bill.ColData(1) = 5
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub



Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub SetDefaultDoctor()
'����:����ȱʡ������
    If cbo������.ListCount = 0 Then Exit Sub
    
    If cbo������.ListCount = 1 Then
        cbo������.ListIndex = 0
    Else
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!סԺҽʦ) Then
                Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, mrsInfo!סԺҽʦ, True))
            End If
        End If
    End If
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.��������ID = lng��������ID Then Exit Sub
    mobjBill.��������ID = lng��������ID
        
    '��������ȷ��ҽ��
    If Not gblnFromDr Then
        If cbo��������.ListIndex <> -1 Then
            If gbln������ Then
                Call FillDoctor(cbo������, mrs������)
            Else
                Call FillDoctor(cbo������, mrs������, lng��������ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
        
    '�������������Ŀ��ִ�п���
    If cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '�������շ���Ŀ
                If .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End With
        Next
    End If
    
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then cbo��������.Text = ""
End Sub

Private Sub cbo������_Click()
    Dim lng������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    If gblnFromDr Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            
            Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lng������ID)
            Call SetDefaultDept(cbo��������, mrs��������, mrs������, lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
    '����������ȷ��������ʱ,���ܴ�ʱ��ѡ������,��ȥ�����������Һ�����ѡ
    If gblnFromDr And gbln������ And cbo������.ListIndex = -1 And txtPatient.Text <> "" Then Cancel = True
End Sub

Private Sub chkCancel_Click()
    Dim i As Long
    
    mstrInNO = ""
    Call NewBill
    Call ClearRows
    Call Bill.ClearBill
    Call ClearMoney
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        
        fraInfo.Enabled = False
        fraAppend.Enabled = False
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = 0
        Next
        Call ShowDeleteCol(True)
        Bill.SetColColor 0, &HE7CFBA '��ȻҪ�ɰ�ɫ
        Bill.Active = True
        
        Call SetDisible
        cboNO.Locked = False
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        Call cbo��������_Click
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        Bill.SetColColor 0, &HE7CFBA '��ȻҪ�ɰ�ɫ
        
        If gbytBilling = 2 Then
            Call SetDisible
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            Call SetDisible(True)
            fraInfo.Enabled = True
            fraAppend.Enabled = True
            Bill.Active = True
            cboNO.Locked = True
            If mbytUseType = 1 And mlng����ID > 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
                Bill.SetFocus
            Else
                txtPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chk�Ӱ�_Click()
    If mbytInState = 1 Or chkCancel.Value = 1 Or gbytBilling = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime(zlDatabase.Currentdate)
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
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If (mobjBill.Details.Count > 0 Or txtPatient.Text <> "") And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        
        If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        If chkCancel.Value = Checked Then '�˾ݵ�״̬
            Call ClearRows: Call Bill.ClearBill
            
            chkCancel.Value = Unchecked
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Bill.Active Then '�������뵥��״̬'(����������²��˵���)
            Call ClearRows: Call Bill.ClearBill
            
            Call NewBill   '����ԭ���ݺ�
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String
    Dim curTotal As Currency, i As Long
    Dim lng����ID As Long, cur���ն� As Currency, cur��� As Currency
    Dim intInsure As Integer, Curdate As Date, blnTrans As Boolean
    Dim str��������IDs As String, str������s As String, cllPro As Collection
    Dim rsItems As ADODB.Recordset
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Visible And chkCancel.Value = 1) Then
        If mbytInState = 0 And mstrInNO = "" Then
            MsgBox "û�ж�ȡ��������,�������ʣ�", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, Bill.Cols - 1) = "��" And Bill.RowData(i) > 0 Then
                strSQL = strSQL & "," & Bill.RowData(i)
            End If
        Next
        If strSQL = "" Then
            MsgBox "������ѡ��һ��Ҫ���ʵķ��ã�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '������ѡ����
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        
        If strSQL <> "" And InStr(1, mstrPrivsOpt, ";��������;") = 0 Then
            MsgBox "��û�в������ʵ�Ȩ�ޣ�ֻ�ܶԸõ���ȫ�����ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlCheckIsExistsApplied(mstrInNO, strSQL, str��������IDs, str������s) Then
            '����:47416
            If MsgBox("ע��:" & vbCrLf & "    ����" & mstrInNO & "�д����������ʵ���Ŀ,���ʺ�,�����Զ�ȡ��" & vbCrLf & "�����˵�������Ŀ,�Ƿ��������?" & vbCrLf & "����������: " & str������s, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        'ҽ�����������ϴ�(ע���ж�˳��)
        If gbytBilling = 0 Then
            intInsure = BillExistInsure(mstrInNO, mstrTime) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, , intInsure) Then
                    'ȥ����ҽ������ƥ����
                    If Not gclsInsure.GetCapability(support�����ݳ�������, , intInsure) And strSQL <> "" Then '���ܲ�������
                        MsgBox "��Ϊҽ��������Ҫ,�õ����е���Ŀ����ȫ�����ʣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
         '����:47416
        Set cllPro = New Collection
        If str��������IDs <> "" Then
            strSQL = "zl_���˷�������_Delete('" & str��������IDs & "')"
            zlAddArray cllPro, strSQL
        End If
        strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        zlAddArray cllPro, strSQL
        
        On Error GoTo errH
         blnTrans = True
         zlExecuteProcedureArrAy cllPro, Me.Caption, True
            'ҽ�����������ϴ�
            If gbytBilling = 0 And intInsure <> 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ�����������ϴ�
        If gbytBilling = 0 And intInsure <> 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        End If
        
        If mbytInState = 0 Then
            mstrInNO = "": cboNO.Text = ""
            txtPatient.Text = "": txtOld.Text = ""
            txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney: Call NewBill
            Call SetMoneyList
            chkCancel.Value = 0
            If gbytBilling = 2 Then
                cboNO.SetFocus
            Else
                txtPatient.SetFocus
            End If
        Else
            Unload Me
        End If
    ElseIf mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        strInfo = Check����ʱ��(CDate(txtDate.Text), cboNO.Text)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not SaveModi() Then Exit Sub
        Unload Me
    ElseIf Bill.Active And chkCancel.Value = 0 Then '�������뵥��״̬
        If mrsInfo.State = adStateClosed Then
            MsgBox "û�з��ֲ�����Ϣ,��ȷ��������Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If txt�ѱ�.Text = "" Or mobjBill.�ѱ� = "" Then
            MsgBox "��ѡ���˷ѱ�", vbInformation, gstrSysName
            txt�ѱ�.SetFocus: Exit Sub
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
        
        If mobjBill.��������ID = 0 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check����ʱ��(CDate(txtDate.Text), mrsInfo!����ID)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If mobjBill.������ = "" And gbln������ Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        If Not PatiCanBilling(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), mstrPrivsOpt) Then Exit Sub
        
        If zlPatiIS�����ѱ�Ŀ(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0)) = True Then     '����:28725
            Exit Sub
        End If
        
        '49501
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
            Exit Sub
        End If
        
        '����ʱ����
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!����) And Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "���õķ���ʱ�䲻��С��ҽ�����˵���Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '�Ƿ���
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        
        'ҽ���������ʼ��    ��Ϊ����Ա�������䵥��,��ȷ������,����Ҫ�ټ��һ��
        If InStr(mstrPrivsOpt, ";���Ƹ�������;") > 0 Then
            If Not IsNull(mrsInfo!����) Then
                If Not gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����) Then
                    For i = 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).���� * mobjBill.Details(i).���� < 0 Then
                                MsgBox "�����е� " & i & " ���Ǹ���,����ҽ����֧�ָ������ʣ�", vbInformation, gstrSysName
                                Bill.SetFocus: Exit Sub
                        End If
                    Next
                End If
            End If
        End If
        
        'Ҫ������,���ҽ������
        If Not IsNull(mrsInfo!����) And Not mrsMedAudit Is Nothing Then
           If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!����) Then Exit Sub
        End If
        
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
        
        '���ʷ��౨��(ֻ��һ�����)
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            curTotal = GetBillTotal(mobjBill)
            If curTotal > 0 Then
                'ˢ�²���Ԥ������Ϣ
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!Ԥ�����
                    cmdCancel.Tag = rsTmp!�������
                    txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                End If
                '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
                sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                
                
                '���¶�ȡ���ն�
                cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                
                cur��� = Val(txtʵ��.Tag)
                If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                
                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), Val(txtҽ�Ƹ���.Tag) = 1 Or Not IsNull(mrsInfo!����), _
                    mrsWarn, cur���, cur���ն� - mcurModiMoney, curTotal, Nvl(mrsInfo!������, 0), "Z", "����", mstrWarn)
                If gbytWarn = 2 Or gbytWarn = 3 Then Exit Sub
            End If
        End If
        
        '��Ŀ���������(��Ҫ��Ϊ�����������۲���)
        If Check������� > 0 Then Exit Sub
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling)) = False Then
            Exit Sub
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
        mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then Exit Sub
        '����
        If Not SaveBill Then
            Exit Sub
        Else
            Call zlChargeSaveAfter_Plugin(mlngModule, mobjBill.����ID, mobjBill.��ҳID, False, 2, mobjBill.NO)
            If gbytBilling = 0 And gbln���ʴ�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=0", "PrintEmpty=0", "�ش�=0", 2)
            ElseIf gbytBilling = 1 And gbln���۴�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=0", "PrintEmpty=0", "�ش�=0", 2)
            End If
            
            If mstrInNO = "" Then
                sta.Panels(2) = "��һ�ŵ���:" & mobjBill.NO
                Call ClearRows: Call Bill.ClearBill
                Call ClearMoney
                Call SetMoneyList
                mstrInNO = ""
                
                If mrsInfo.State = 1 Then
                    Call NewBill(False)
                    txtPatient.Tag = "-" & mrsInfo!����ID
                    
                    With mobjBill
                        .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                        .��ҳID = IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID)
                        
                        .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                        .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                        
                        .���� = "" & mrsInfo!����
                        .��ʶ�� = IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��)
                        .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
                        .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        .�ѱ� = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
                        
                        .Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
                        .������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
                    End With
                    
                    If mbytUseType = 1 Then
                        Call txtPatient_KeyPress(13) 'ˢ��һЩ������Ϣ
                        Bill.SetFocus
                    Else
                        txtPatient.SetFocus
                    End If
                Else
                    Call NewBill
                    txtPatient.SetFocus
                End If
            Else '�޸�
                Unload Me
            End If
        End If
    ElseIf Not Bill.Active Then '���סԺ����״̬
        If mstrInNO = "" Then
            MsgBox "û��סԺ���۵���,�������룡", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        'ȡ������˵������
        strSQL = ""
        For i = 1 To Bill.Rows - 1
            If Bill.RowData(i) > 0 Then
                strSQL = strSQL & "," & Bill.RowData(i)
            End If
        Next
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        'ҽ�����
        intInsure = BillExistInsure(mstrInNO, , True)
        If intInsure > 0 Then
            'ȥ����ҽ������ƥ����
        End If
        
        '���ñ���
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            If Not AuditingWarn(mstrPrivsOpt, mrsWarn, mstrInNO, strSQL) Then Exit Sub
        End If
        
        Curdate = zlDatabase.Currentdate
        strSQL = "zl_סԺ���ʼ�¼_Verify('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & strSQL & "',NULL,To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            'ҽ���ϴ�
            If intInsure <> 0 Then
                'ҽ�����������ϸ
                If gclsInsure.GetCapability(support�����ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                    strInfo = ""
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                        gcnOracle.RollbackTrans
                        If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ���ϴ�
        If intInsure <> 0 Then
            'ҽ�����������ϸ
            If gclsInsure.GetCapability(support�����ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                strInfo = ""
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                    If strInfo <> "" Then
                        MsgBox strInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mstrInNO & """��������ҽ������ʧ��,�õ�������ˣ�", vbInformation, gstrSysName
                    End If
                    Exit Sub
                End If
            End If
        End If
        
        On Error GoTo 0
        
        If gbytBilling = 2 And gbln��˴�ӡ And mblnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mstrInNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=0", "PrintEmpty=0", "�ش�=0", 2)
        End If
        
        mstrInNO = "": cboNO.Text = ""
        txtPatient.Text = "": txtOld.Text = ""
        txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney: Call NewBill
        Call SetMoneyList
        cboNO.Locked = False: cboNO.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mbytUseType = 1 And mlng����ID <> 0 And mbytInState = 0 Then
        If gblnFromDr Then
            cbo������.SetFocus
        Else
            Bill.SetFocus
        End If
    ElseIf gbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    ElseIf mstrInNO <> "" Then
        Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Long, strPre As String, lngPre As Long, strTmp As String
    
    RestoreWinState Me, App.ProductName, mbytInState
    
    gblnOK = False: mblnValid = False
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    Call initCardSquareData
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        If Not InitData Then Unload Me: Exit Sub
    Else
        If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
            Exit Sub
        End If
    End If
    mstrUnitIDs = GetUserUnits
    
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then '��ʾ�����������ʵ���(1,2,3)
        If Not ReadBill(mstrInNO, (mbytInState = 3)) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
    Else '����
        mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
        mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
        mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportBill(mstrInNO, False, Me, True, True, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            If mobjBill.NO = "" Then
                MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                mcurModiMoney = GetBillMoney(2, mobjBill.NO) 'Ҫ�ڶ�ȡ������Ϣǰ�ȶ�
                
                lngPre = mobjBill.��������ID
                strPre = mobjBill.������
                
                txtPatient.Text = "-" & mobjBill.����ID
                Call txtPatient_KeyPress(13)
                                
                Call ReCalcInsure '���¼���ͳ����
                
                '��ʾ����ԭ���ݺ�,��������µ��ݺ�
                cboNO.Text = mobjBill.NO
                Bill.ClearBill
                Bill.Rows = mobjBill.Details.Count + 1
                '����б༭����������ɫ
                Bill.SetColColor 0, &HE7CFBA
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 3, &HE7CFBA
                
                txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
                chk�Ӱ�.Value = mobjBill.�Ӱ��־
                
                mobjBill.��������ID = lngPre
                mobjBill.������ = strPre
                Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
                Call zlControl.CboLocate(cboBaby, mobjBill.Ӥ����, True)
                                
                '�޸�ʱӦ���浱ǰ����Ա������
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                Call ShowDetails
                Call ShowMoney
                
                For i = 1 To mobjBill.Details.Count
                    '���⴦��
                    Bill.RowData(i) = Asc(mobjBill.Details(i).�շ����)
                Next
            End If
        Else
            If mbytUseType = 1 And mlng����ID <> 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
            End If
        End If
    End If
    '����:47798
    If mbytInState = 0 Then
        Call GetRegisterItem(g˽��ģ��, Me.Name, "idkind", strTmp)
        Err = 0: On Error Resume Next
        mblnNotCick = True
        IDKIND.IDKIND = Val(strTmp)
        mblnNotCick = False
        Err = 0: On Error GoTo 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mbytInState
    mbytInState = Empty
    mstrInNO = Empty
    mblnNOMoved = False  '�����˳������,����Ӱ���������
    mstrTime = ""
    mblnDelete = False
    gbytBilling = 0
    mlngDeptID = 0
    mbytUseType = 0
    mlng����ID = 0
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrsWarn = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    '����:47798
    If mbytInState = 0 Then
        Call SaveRegisterItem(g˽��ģ��, Me.Name, "idkind", IDKIND.IDKIND)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - fraTitle.Height - fraInfo.Height + 230
    Me.Refresh
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
             Call FindPati(objCard, True, txtPatient.Text)
        End If
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotCick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    If txtPatient.Locked Then Exit Sub
    If objPatiInfor.���� = "" Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    Call FindPati(objCard, True, txtPatient.Text)
     
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln�����л� = False Then Exit Sub
    
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
        gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
    End If
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    If gbytBilling = 2 Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String
    Dim intInsure As Integer, vDate As Date, i As Long
    Dim strInfo As String, intTmp As Integer
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)
        
        If chkCancel.Value = 1 Then
            '����
            
            If gbytBilling = 0 Then
                '�Ƿ���ת������ݱ���
                If zlDatabase.NOMoved("סԺ���ü�¼", cboNO.Text, , 2, Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                    mblnNOMoved = False
                End If
            End If
            
            '�����˻���ȫ��˵Ĳ���������
            If Not BillIdentical(cboNO.Text) Then
                MsgBox "�����а������ݲ�ȫ����˻�ֶ����˵����ݣ����������������ʡ�" & _
                    vbCrLf & "���˻ع��������˳���Ӧ�ĵ������ݣ�Ȼ�������ʡ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '����Ȩ��
            If Not ReadBillInfo(2, cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, ";���в���Ա;") <= 0 Then
                If UserInfo.���� <> strOper Then
                    MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ�������!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "����", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '��Ŀ����Ȩ��
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If Not CheckDelPriv(cboNO.Text, mstrPrivsOpt) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            
            '���۲���Ȩ��
            strInfo = Check���۲���(cboNO.Text, mstrPrivsOpt)
            If strInfo <> "" Then
                MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '�Ƿ���ִ��
            i = BillCanDelete(cboNO.Text, 2)
            If i <> 0 Then
                Select Case i
                    Case 1 '�õ��ݲ�����
                        MsgBox "ָ�������е����ݲ����ڣ�", vbInformation, gstrSysName
                    Case 2 '�Ѿ�ȫ����ȫִ��
                        MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
                    Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                        MsgBox "ָ�������е�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "����") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '�Ƿ��Ѿ�����
            intTmp = HaveBilling(2, cboNO.Text, False)
            If intTmp <> 0 Then
                intInsure = BillExistInsure(cboNO.Text)
                If intInsure <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure) Then
                        'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                        If intTmp = 1 Then
                            MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        Else
                            MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            Else
                                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                            End If
                    End Select
                End If
            End If
            
            '�Ƿ������������¼
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
                    "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Me.Caption
            End If
        ElseIf mobjBill.Details.Count = 0 Then
            '���ʻ��۵�(�������)
            
            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "���") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            If Not BillExistMoney(cboNO.Text, 2) Then
                MsgBox "�õ��ݷ����Ѿ�ȫ�����ʻ򵥾ݲ����ڣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
        
        '���ʻ����ʱ,���ݱ���Ϊ���շѵĵ���
        If Not BillisSimple(cboNO.Text) Then
            MsgBox "�õ��ݲ����ڻ��Ǽ򵥼��ʵ��ݣ�", vbInformation, gstrSysName
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        
        If chkCancel.Value = 1 Then '��ȡ�˷ѵ�
            blnRead = ReadBill(cboNO.Text, True)
        ElseIf mobjBill.Details.Count = 0 Then '��ȡסԺ���۵�
            blnRead = ReadBill(cboNO.Text, False)
        End If
        
        If blnRead Then
            mstrInNO = cboNO.Text 'ȷ��ʱ��mstrInNOΪ׼
            If chkCancel.Value = 0 Then '���۵�
                Bill.Active = False
            Else '����
                'Call SetDisible 'cboNO�ڻ�ȡ�����unLock
                Bill.Active = True
            End If
            cmdOK.SetFocus
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub txtOld_Gotfocus()
    zlControl.TxtSelAll txtOld
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mobjBill.���� = txtOld.Text
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
    
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0
    txtPatient.SelLength = Len(txtPatient.Text)
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
    
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        'Bill.RemoveMSFItem Row'������AllowAddRow����
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        .ColData(1) = 5 'Ӧ��ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
        '����б༭����������ɫ
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
    End With
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lngҽ��ID)
    
    If zlSelectDept(Me, mlngModule, cbo��������, mrs��������, cbo��������.Text) = False Then
        Call Beep: mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Exit Sub





''    Dim lngIdx As Long
''
''    If KeyAscii >= 32 And Not cbo��������.Locked Then
''        lngIdx = zlControl.CboMatchIndex(cbo��������.hwnd, KeyAscii)
''        If lngIdx = -1 And cbo��������.ListCount > 0 Then lngIdx = 0
''        cbo��������.ListIndex = lngIdx
''    ElseIf KeyAscii = 13 Then
''        If cbo��������.ListIndex = -1 Then
''            Beep
''        Else
''            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
''            Call zlCommFun.PressKey(vbKeyTab)
''        End If
''    End If
End Sub
Private Function isCheck������Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
    
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
            
            
            '���˺�:22383
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs������.Filter = strFilter: iCount = 0
            With mrs������
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs������.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(mrs������!���) Like strText & "*" Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    End Select
                    mrs������.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            '���˺�:ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
                    If gbyt��������ʾ = 1 Then '����
                        rsTemp.Sort = "����"
                    Else
                        rsTemp.Sort = "���"
                    End If
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheck������Exists(Nvl(rsReturn!����), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing

            
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = ""
            If gblnFromDr Then Exit Sub
        Else
            mobjBill.������ = zlStr.NeedName(cbo������.Text)
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
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF6 '�����ǰ��������,�����µ�״̬
            txtPatient.SetFocus
            Call zlControl.TxtSelAll(txtPatient)
        Case vbKeyF7 '�л����뷨
            If gbln�����л� = False Then Exit Sub   '34242
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '��(�Զ������¼�)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String
    Dim Curdate As Date     '��������ǰʱ��
    Err = 0: On Error GoTo errH:
    Curdate = zlDatabase.Currentdate
    
    '�Զ�ʶ��Ӱ�
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk�Ӱ�.Value = Checked
    End If
            
    If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
        Exit Function
    End If
    
    'ִ�в���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID and B.������� IN(2,3) " & _
        " Order by B.�������,A.����"
    Set mrsUnit = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsUnit, strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '���õ��ݱ��ʽ
    With Bill
        .LocateCol = 0
        .PrimaryCol = 0
        .Font.Size = 11
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        
        If mbytInState = 0 And gbytBilling <> 2 Then
            .ColData(0) = 1 '��Ŀ����,��Ť��ѡ
            .ColData(1) = 5 'Ӧ�ս��ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(2) = 5 'ʵ�ս������
            .ColData(3) = 3 'Ĭ��ȡ�������һ���һ����
            .ColData(4) = 5
        End If
        
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    Call SetMoneyList
    
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("PY").Visible = mbytInState = 0 And gbln�����л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln�����л�
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If

    '����
    Select Case gbytBilling
        Case 0
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�"
        Case 1
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�(����)"
        Case 2
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�(���)"
    End Select
    
    txtӦ��.Text = gstrDec: txtʵ��.Text = gstrDec
    
    Select Case mbytInState
        Case 0 'ִ��
            If mstrInNO <> "" Or _
                (InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
                    And InStr(mstrPrivsOpt, ";��������;") = 0 _
                    And InStr(mstrPrivsOpt, ";��������;") = 0) Then
                chkCancel.Visible = False
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            Select Case gbytBilling
                Case 0, 1 'ִ�м��ʡ�����
                    txtPatient.Enabled = (mstrInNO = "")
                Case 2 'ִ�����
                    Call SetDisible
                    cboNO.Locked = False
                    fraInfo.Enabled = False
                    fraAppend.Enabled = False
                    Bill.Active = False
            End Select
        Case 1 '����
            Call SetDisible
            
            chkCancel.Visible = False
            If mblnDelete Then
                lblFlag.Visible = True
            Else
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Case 2 '����
            Call SetDisible
            txtDate.Enabled = True
            chkCancel.Visible = False
            lblNO.Left = lblNO.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            Bill.Active = False
        Case 3 '����
            Call SetDisible
            
            chkCancel.Visible = False
            lblNO.Left = lblNO.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraAppend.Enabled = False
            
            Call ShowDeleteCol(True)
            Bill.Active = True
    End Select
    
    '�������������뿪����λ��
    If gblnFromDr Then
        Call ExChangeLocate(cbo��������, cbo������)
        Call ExChangeLocate(lbl��������, lbl������)
        cbo��������.TabStop = False
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    txtPatient.Locked = Not bln
    cbo��������.Locked = Not bln
    chk�Ӱ�.Enabled = bln
    
    cbo������.Locked = Not bln
    txtDate.Enabled = bln
    Bill.Active = bln
    
    If Not bln Then
        txtPatient.BackColor = &HE0E0E0
        txtҽ�Ƹ���.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
        txtҽ�Ƹ���.BackColor = &HFFFFFF
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            If (mbytUseType = 0 Or mbytUseType = 1) Then
                .mlngUnitID = mlngUnitID
            Else
                .mlngUnitID = mlngDeptID
            End If
            .mbytUseType = mbytUseType
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
        End With
    Else
        If IDKIND.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.���� = "�����" Or IDKIND.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            
            'ˢ�²�����Ϣ:"-����ID"
            Call GetPatient(IDKIND.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '��������ʱ�����ܴ�ʱ����������˷��ã�������Աû��"��Ժδ��ǿ�Ƽ���"Ȩ�ޣ�����������
                txtPatient.Text = "": txtOld.Text = ""
                txt����.Text = ""
                Exit Sub
            End If
            
            'ˢ�²���Ԥ������Ϣ
            curTotal = GetBillTotal(mobjBill)
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, CDbl(mcurModiMoney), True, 2)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!Ԥ�����
                cmdCancel.Tag = rsTmp!�������
                txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
            End If
            '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
            sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
            strInfo = GetPatientDue(Val(mrsInfo!����ID))
            If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/Ӧ�տ�:" & Format(strInfo, "0.00")
            Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
            If Not mblnValid Then Bill.SetFocus
            Exit Sub
        End If
        KeyAscii = 0
        '69282,������,2014-01-03,ͨ������+סԺ�ŷ�ʽ�Ҳ��˳��������
        Call FindPati(IDKIND.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMsg As Boolean
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    
    '20030617:�����δ������
    If mobjBill.Details.Count = 0 Then
        Call ClearMoney
        txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
    End If
    
    '��ȡ������Ϣ
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "��һ��*") Then
        sta.Panels(2) = ""
    End If
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = "": txt����.Text = "": Exit Sub
        Else
            If Not blnMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
            zlControl.TxtSelAll txtPatient
            If mstrInNO = "" Then txtOld.Text = "": txt����.Text = ""
            Exit Sub
        End If
        Exit Sub
    End If
    
    '���￨������
    If (objCard.���� Like "*IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ And mstrPassWord = "" Then blnCard = False
    If Mid(gstrCardPass, 6, 1) = "1" And blnCard Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    
    If mbytUseType = 1 And mrsInfo!����ID <> mlng����ID Then mlng����ID = 0
     
      '�Զ����ÿ�������(ͬʱ���ü��ʱ�����Ϣ),ҽ�����ʲ��˿��Ҳ�һ���ǿ�������
     If mbytUseType = 2 Then lngUnit = cbo��������.ListIndex
    
    If gblnFromDr Then
        If Not IsNull(mrsInfo!סԺҽʦ) Then
            cbo������.ListIndex = -1
            cbo������.ListIndex = cbo.FindIndex(cbo������, mrsInfo!סԺҽʦ, True)
        End If
    Else
        cbo��������.ListIndex = -1
        cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID))
        If cbo��������.ListIndex <> -1 Then
            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        ElseIf mbytUseType = 2 And lngUnit <> -1 Then
            cbo��������.ListIndex = lngUnit
        End If
    End If
    
    '����Ԥ������Ϣ
    curTotal = GetBillTotal(mobjBill)
    Set rsTmp = GetMoneyInfo(mrsInfo!����ID, CDbl(mcurModiMoney), True, 2)
    If Not rsTmp Is Nothing Then
        cmdOK.Tag = rsTmp!Ԥ�����
        cmdCancel.Tag = rsTmp!�������
        txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
    Else
        cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    '���˺�:26952
    Dim cur��� As Currency, curItemMoney As Currency
    
    cur��� = Val(txtʵ��.Tag)
    
    '���˺�:24491
    curItemMoney = 0 ' GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
     
    If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
    
    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), "", "", _
                 mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1), curItemMoney, True)
    '����:0;û�б���,����
    '     1:������ʾ���û�ѡ�����
    '     2:������ʾ���û�ѡ���ж�
    '     3:������ʾ�����ж�
    '     4:ǿ�Ƽ��ʱ���,����
    '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
    If gbytWarn = 2 Or gbytWarn = 3 Then
        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "":
         mlng����ID = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
    sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
    sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
    sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
    strInfo = GetPatientDue(Val(mrsInfo!����ID))
    If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/Ӧ�տ�:" & Format(strInfo, "0.00")
    
    Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
                
    '������Ϣ
    txtPatient.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    txtSex.Text = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
    txtOld.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    txt�ѱ�.Text = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
    txtҽ�Ƹ���.Text = IIf(IsNull(mrsInfo!ҽ�Ƹ��ʽ), "", mrsInfo!ҽ�Ƹ��ʽ)
    txtҽ�Ƹ���.Tag = GetMedPayMode(txtҽ�Ƹ���.Text, mrsMedPayMode)
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), txtҽ�Ƹ���.Text, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
        If mobjBill.Details.Count > 0 Then
            '���¼��㲢ˢ��
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
    txt����.Text = "" & mrsInfo!����
    txt������.Text = IIf(IsNull(mrsInfo!������), "", mrsInfo!������)
    txt������.Text = Format(IIf(IsNull(mrsInfo!������), "", mrsInfo!������), "0.00")
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
     
     With mobjBill
         .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
         .��ҳID = IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID)
         
         .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
         .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
         
         .���� = "" & mrsInfo!����
         .��ʶ�� = IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��)
         .���� = txtPatient.Text
         .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
         .���� = txtOld.Text
         .�ѱ� = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
     End With
     If Not IsNull(mrsInfo!��Ժ����) Then
         MsgBox "��������" & vbCrLf & vbCrLf & "�ò������� " & Format(mrsInfo!��Ժ����, "yyyy-MM-dd") & " ��Ժ�����ڶԸò���ǿ�ƽ��м��ʣ�", vbInformation, gstrSysName
         txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
     Else
         txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
     End If
     If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "��һ��*") Then
         If Not IsNull(mrsInfo!��Ժ����) Then
             sta.Panels(2).Text = "��Ժ:" & Format(mrsInfo!��Ժ����, "yyyy-MM-dd")
             strInfo = GetInsureInfo(mrsInfo!����ID)
             If strInfo <> "" Then sta.Panels(2).Text = sta.Panels(2).Text & "/�ʺ�:" & Split(strInfo, ";")(1)
         End If
     End If
     If Visible Then
        If gblnFromDr Then
            cbo������.SetFocus
        Else
            cbo��������.SetFocus
        End If
     End If
 End Sub


Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:
    '����:���ҵ�����,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 17:54:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strWhere As String
    Dim rsOutSel As ADODB.Recordset, bln���в��� As Boolean
    On Error GoTo errH
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    'b.�Ƿ���Լ����в�������
    bln���в��� = True
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        bln���в��� = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.��ǰ����ID+0=[3]"
        Else
            strIF = strIF & " And B.��ǰ����ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.��������,0)=0"
    End If
    
    strSQL = _
    " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
    "       A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
    "       Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
    "       A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
    "       Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,B.��˱�־,B.��������" & _
    " From ������Ϣ A,������ҳ B,������� X " & _
    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
    "        And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2  And A.ͣ��ʱ�� is NULL " & strIF
    If blnCard = True And objCard.���� Like "����*" Then  'ˢ��
    
        If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
            lng�����ID = IDKIND.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strWhere = strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '��λ��
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '������ȷ��������ͨ������ȷ������
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "       A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "       Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "       A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "       Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,B.��˱�־,B.��������" & _
            " From ������Ϣ A,������ҳ B,��λ״����¼ C,������� X" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            "       And Nvl(B.��ҳID,0)<>0 And A.����ID=C.����ID And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL" & _
            "       And C.����ID=[3] And C.����=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
        strWhere = strWhere & " And A.�����=[1]"
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If Not mrsInfo.EOF Then
                        If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
                If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln���в���, mstrUnitIDs, gintOutDay, lng����ID, strErrMsg, txtPatient.hWnd, txtPatient.Height) = False Then
                    If strErrMsg = "" Then blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then GoTo GoYJReadPati:
                    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                End If
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
         End Select
    End If
    
    strSQL = strSQL & vbCrLf & strWhere
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngUnitID, mstrUnitIDs)
    
    If Not mrsInfo.EOF Then
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then    '����:28725
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        
        If mrsInfo!����ID <> mobjBill.����ID Or mbytInState = 0 And mstrInNO <> "" Then    'ͬһ���˲����ظ���ȡ
            If GetMedPayMode("" & mrsInfo!ҽ�Ƹ��ʽ, mrsMedPayMode) = 1 Then
                Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
            Else
                Set mrsMedAudit = Nothing
            End If
        End If
         mstrPassWord = strPassWord
        If Not blnHavePassWord Then
            mstrPassWord = Nvl(mrsInfo!����֤��)
        End If
        GetPatient = True
        Exit Function
    Else
        Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������
    End If
    
        
    'ҽ�����Ҽ��ʣ�û�з���סԺ(��Ժ���Ժ)����,�����ﲡ�˶�
    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
GoYJReadPati:
        '76451,Ƚ����,2014-8-19
        strSQL = _
        " Select A.����ID,Nvl(A.��ҳID,0) as ��ҳID,A.��ǰ����ID as ����ID,A.��ǰ����ID as ����ID," & _
        "       A.��Ժʱ�� as ��Ժ����,A.���￨��,A.����֤��,A.סԺ��,A.��ǰ���� as ����,A.����,A.�Ա�,A.����," & _
        "       A.��Ժʱ�� as ��Ժ����,A.�ѱ�,A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,null)) ������,Zl_Patiwarnscheme(A.����id) As ���ò���,NULL as סԺҽʦ,A.ҽ�Ƹ��ʽ," & _
        "       zl_PatiDayCharge(A.����ID) as ���ն�,A.����,-1 as ��������" & _
        " From ������Ϣ A Where A.ͣ��ʱ�� is NULL "
        If blnCard = True And objCard.���� Like "����*" Then   'ˢ��
            If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
                lng�����ID = IDKIND.GetfaultCard.�ӿ����
            Else
                lng�����ID = "-1"
            End If
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
            If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            If lng����ID <= 0 Then GoTo NotFoundPati:
            strInput = "-" & lng����ID
            blnHavePassWord = True
            strSQL = strSQL & " And A.����ID=[1] "
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
            strSQL = strSQL & " And A.����ID=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
            strSQL = strSQL & " And A.�����=[1]"
        Else '��������
            Select Case objCard.����
                  Case "����", "��������￨"
                      If mrsInfo.State = 1 Then
                          If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                      End If
                      strSQL = strSQL & " And A.����=[2]"
                  Case "ҽ����"
                      strInput = UCase(strInput)
                      strSQL = strSQL & " And A.ҽ����=[2]"
                  Case "�����"
                      If Not IsNumeric(strInput) Then strInput = "0"
                      strSQL = strSQL & " And A.�����=[2]"
                  Case "סԺ��"
                      If Not IsNumeric(strInput) Then strInput = "0"
                      strSQL = strSQL & " And A.סԺ��=[2]"
                  Case Else
                      '��������,��ȡ��صĲ���ID
                      If objCard.�ӿ���� > 0 Then
                          lng�����ID = objCard.�ӿ����
                          If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                          If lng����ID = 0 Then GoTo NotFoundPati:
                      Else
                          If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                              strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                      End If
                      If lng����ID <= 0 Then GoTo NotFoundPati:
                      strSQL = strSQL & " And A.����ID=[1]"
                      strInput = "-" & lng����ID
                      blnHavePassWord = True
               End Select
        End If
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
        If Not mrsInfo.EOF Then
            If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then    '����:28725
                Set mrsInfo = New ADODB.Recordset
                blnOutMsg = True
                Exit Function
            End If
            mstrPassWord = strPassWord
            If Not blnHavePassWord Then
               mstrPassWord = Nvl(mrsInfo!����֤��)
            End If
            GetPatient = True
            Exit Function
        End If
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������'
    Set mrsInfo = New ADODB.Recordset
    If strWhere = "" Then Exit Function '������������ֱ���˳�
    
    'δ�ҵ����ˣ���Ҫ�Ըò��˵ľ��������Ϣ������ʾ
    strSQL = _
    " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,a.��Ժ,B.��Ժ����,B.��Ժ����,X.�������,B.״̬, " & _
    "       nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,Nvl(B.��������,0) as ��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,������� X" & _
    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
    "   And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsOutSel.EOF Then Exit Function
    
    '1.�������
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!����ID) & ",") = 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "�������㸺��Ĳ���,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '2.���۲��˼��(�Ƿ����۲��˼���Ȩ��)
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        If Val(Nvl(rsOutSel!��������)) = 2 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��ΪסԺ���۲���,�㲻�߱���סԺ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        If Val(Nvl(rsOutSel!��������)) = 1 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ�������۲���,�㲻�߱����������ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!��������)) <> 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ" & IIf(Val(Nvl(rsOutSel!��������)) = 1, "����", "סԺ") & "���۲���,�㲻�߱��������סԺ ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    '124007
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) <> 0) Then
              
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                    strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                    strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) = 0) Then
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����)) Then
            If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
            Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
            End If
        End If
    End If
    
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������'
        blnOutMsg = True
        Exit Function
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    If mobjBill.Details.Count = 0 Then Exit Sub
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(lngRow As Long)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '�û�����ı�۽��
    Dim dbl�Ӱ�Ӽ��� As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    If mstr��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    
    gstrSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,b.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID, mstr��ͨ�۸�ȼ�)
    
    If rsTmp.EOF Then
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        Exit Sub
    End If
    
    '�Ȼ�ȡ����Ա��ǰ����ı�۽��
    With mobjBill.Details(lngRow)
        If .Detail.��� Then
            If .InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
                dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
            Else                        '��ȡ����Ա��ǰ����ı�۽��
                dblMoney = .InComes(1).��׼����
                '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                If CheckScope(Val(Nvl(rsTmp!ԭ��)), Val(Nvl(rsTmp!�ּ�)), dblMoney) <> "" Then
                    dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
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
            .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
            .ԭ�� = Nvl(rsTmp!ԭ��, 0)
            .�ּ� = Nvl(rsTmp!�ּ�, 0)
            If mobjBill.Details(lngRow).Detail.��� Then
                .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
            Else
                .��׼���� = Format(Nvl(rsTmp!�ּ�), gstrFeePrecisionFmt)
            End If
            
            'Ӧ�ս��=���� * ���� * ����
            .Ӧ�ս�� = .��׼���� * IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����) * mobjBill.Details(lngRow).����
            '�������������ü���(����������Ŀ)
            If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
            End If
            '�Ӱ�����ʼ���
            dbl�Ӱ�Ӽ��� = 0
            If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)
                .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
            End If
            
            .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
            
            If mobjBill.Details(lngRow).Detail.���ηѱ� Then
                .ʵ�ս�� = .Ӧ�ս��
            Else
                If .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                    mobjBill.Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                    .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, 0, 0, 0, dbl�Ӱ�Ӽ���), gstrDec))
                End If
            End If
            
            '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
            If mrsInfo.State = 1 Then
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details(lngRow).ժҪ & "||" & mobjBill.Details(lngRow).����)
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
            End If
            
            mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, , .ͳ����
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    Dim curTotal As Currency
    
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
        '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
        sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim curMoney As Currency
    Dim i As Long, j As Long
    '���������
    For i = 0 To Bill.Cols - 1
        '����ʱ�շ�������
        If Not (i = 0 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    'ˢ�µ�����
    For i = 0 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "Ӧ�ս��" 'ʵ�����ǵ���
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                curMoney = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                curMoney = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "ִ�п���"
                If mbytInState = 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                    Else
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    '�������ֻ(��)��ʾ����
                    Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Private Function GetInputDetail(ByVal lng��Ŀid As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select A.ID,A.���,B.���� as �������,A.����,A.����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,F.Ҫ������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,����֧����Ŀ F" & _
            " Where A.���=B.���� And A.ID=[1] And A.ID=F.�շ�ϸĿID(+) And F.����(+)=[2]"
    Else
        strSQL = _
            " Select A.ID,A.���,B.���� as �������,A.����,A.����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,0 as Ҫ������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B" & _
            " Where A.���=B.���� And A.ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, Optional bytParent As Byte = 0, Optional ByVal lngDoUnit As Long)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������
    Dim tmpIncomes As New BillInComes
    Dim dblTime As Double, i As Long
        
     'ִ�п���
    If bytParent <> 0 Then
        '������Ŀ��ִ�п���,��������������ͬ,����Ϊ����ȷִ�п���,��ȡ����ִ�п���,����ȡ�����
        If lngDoUnit <> 0 Then
            lngDoUnit = mobjBill.Details(bytParent).ִ�в���ID
        Else
            If cbo��������.ListIndex <> -1 Then lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
            
            lngDoUnit = Get�շ�ִ�п���ID("Z", Detail.ID, Detail.ִ�п���, lngDoUnit, Get��������ID, Get������Դ, , mobjBill.����ID)
        End If
    Else
        lngDoUnit = mobjBill.����ID
        If lngDoUnit = 0 And cbo��������.ListIndex <> -1 Then
            lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        lngDoUnit = Get�շ�ִ�п���ID("Z", Detail.ID, Detail.ִ�п���, lngDoUnit, Get��������ID, Get������Դ, , mobjBill.����ID)
    End If
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = mobjBill.Details(bytParent).����
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                dblTime = 1
            End If
            mobjBill.Details.Add Detail, .ID, CByte(lngRow), CInt(bytParent), 0, 0, 0, 0, "", "", "", _
            0, 0, mobjBill.�ѱ�, 0, .���, .���㵥λ, "", 1, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else
        '��������Ѿ�����,���޸�
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .�ѱ� = mobjBill.�ѱ�
            .���� = 1
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = 1
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
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
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    strSQL = _
        "Select A.ID,A.���,B.���� as �������,A.��������,A.����,A.����,A.���,A.Ҫ������," & _
        " A.���㵥λ,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,C.���д���,C.��������" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.���='Z' And C.����ID=[1]" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .���� = rsTmp!����
            .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
            .��� = Nvl(rsTmp!���)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
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
            .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
            GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                1, .���㵥λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, , , , , , , .Ҫ������
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
        For i = 0 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill(Optional blnPati As Boolean = True)
'���ܣ���ʼ��һ���µĵ���(�������)
'������blnPati=�Ƿ��ʼ��������Ϣ
    Dim blnKeepDate As Boolean
    Dim Curdate As Date     '��������ǰʱ��
    mcurModiMoney = 0
            
    If mrsInfo.State = 0 Then txtPatient.ForeColor = Me.ForeColor
    
    If blnPati Then
        cmdOK.Tag = "": cmdCancel.Tag = "": txtʵ��.Tag = ""
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txt�ѱ�.Text = "": txt����.Text = "": txtҽ�Ƹ���.Text = ""
        txt������.Text = "": txt������.Text = ""
                
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
        sta.Panels(3).Text = ""
    End If
            
    mstrWarn = ""
    cboNO.Text = ""
    Set mobjBill = New ExpenseBill
        
    Curdate = zlDatabase.Currentdate
    chk�Ӱ�.Value = IIf(OverTime(Curdate), 1, 0)
    
    If Not blnPati And mrsInfo.State = 1 Then
        If mrsInfo!��Ժ���� < Curdate Then blnKeepDate = True
    End If
    If Not blnKeepDate Then txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    
    Call cbo��������_Click
    
    With mobjBill
        .�����־ = 2
        .������ = UserInfo.����
        .������ = zlStr.NeedName(cbo������.Text)
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .Ӥ���� = 0
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
    End With
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim i As Long, j As Long, arrSQL As Variant
    Dim int��� As Integer, int�к� As Integer, strNO As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim str��Ϣ As String, intInsure As Integer, blnTrans As Boolean
    
    mobjBill.NO = zlDatabase.GetNextNo(14)
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Details
        intParent = 0: intParentNO = int���
        For Each mobjBillIncome In mobjBillDetail.InComes
            int��� = int��� + 1 '��ǰ��¼���
            
            '��������
            With mobjBill
                gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & _
                    IIf(Val(.��ʶ��) = 0, "NULL", .��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                    IIf(.����ID = 0, .��������ID, .����ID) & "," & IIf(.����ID = 0, .��������ID, .����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
            End With
            
            '�շ�ϸĿ����
            With mobjBillDetail
                '�����������
                If .��� <> int�к� Then
                    int�к� = .���
                    '���´����������
                    If mobjBill.Details(.���).�������� = 0 Then
                        For i = .��� + 1 To mobjBill.Details.Count
                            If mobjBill.Details(i).�������� = .��� Then
                                mobjBill.Details(i).�������� = int��� '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                            End If
                        Next
                    End If
                End If
                gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                
                gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & IIf(.���մ���ID = 0, "NULL", .���մ���ID) & ",'" & .���ձ��� & "',"
                
                gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & .���� & "," & .���ӱ�־ & "," & .ִ�в���ID & ","
            End With
            
            '������Ŀ����
            With mobjBillIncome
                intParent = intParent + 1
                gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .������ĿID & "," & _
                    "'" & .�վݷ�Ŀ & "'," & .��׼���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & "," & _
                    IIf(.ͳ���� = 0, "NULL", .ͳ����) & ","
            End With
                                            
            '��������:�����Ϊ�Ǽ򵥼���(��ҩ����)
            gstrSQL = gstrSQL & _
                "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & mstrInNO & "'," & IIf(gbytBilling = 1, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'" & mobjBillDetail.Detail.���� & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
        Next
    Next
    
    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        '���ж��Ƿ�ҽ�����˼ǵ���,�����Ϸ��Լ��(�����޸�ʱ������һ������ж�)
        If gbytBilling = 0 Then
            intInsure = BillExistInsure(mstrInNO)
            If intInsure > 0 Then
                'ȥ����ҽ������ƥ����
            End If
        End If
        
        gstrSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
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
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�
            If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
                    
            '2.����ʵʱ�ϴ�
            If gbytBilling = 0 And Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If gclsInsure.GetCapability(support�����ϴ�, , mrsInfo!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsInfo!����) Then
                    str��Ϣ = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '1.ҽ�����������ϴ�
        If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ��,�õ��������ʣ�", vbInformation, gstrSysName
                End If
            End If
        End If
                
        '2.����ʵʱ�ϴ�
        If gbytBilling = 0 And Not IsNull(mrsInfo!����) Then
            'ҽ�����������ϸ
            If gclsInsure.GetCapability(support�����ϴ�, , mrsInfo!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsInfo!����) Then
                str��Ϣ = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                    If str��Ϣ <> "" Then
                        MsgBox str��Ϣ, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For 'ֻ��ʾ10��
        Next
        
        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Boolean
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strNO=���ݺ�
'      blnDelete=True:���ʵ���ʱ����,False:���ĵ���ʱ����
    Dim rsTmp As ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim curTotal As Currency, curӦ��Total As Currency
    Dim intInsure As Integer, blnDo As Boolean
    Dim strSQL1 As String, intSign As Integer
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    mblnPrint = False
        
    Call ClearRows: Call Bill.ClearBill: Call ClearMoney
    
    '��ȡ��������
    strNO = GetFullNO(strNO, 14)
   
    strSQL = _
    " Select A.����ID,Nvl(A.��ҳID,0) as ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.����," & _
    "       A.���˲���ID,A.��������ID,Nvl(A.�Ӱ��־,0) as �Ӱ��־,Nvl(A.Ӥ����,0) as Ӥ����," & _
    "       A.������,A.������,A.����Ա����,A.����ʱ��,A.����ID,B.������,B.������" & _
    " From " & IIf(mblnNOMoved And gbytBilling = 0, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ,������Ϣ B,��Ա�� C " & _
    " Where NO=[1] And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0 And Nvl(A.����Ա����,A.������)=C.����" & _
    "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
    "       And A.����ID=B.����ID And Rownum=1 And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            IIf(mbytInState = 0 And gbytBilling = 0, " And A.����Ա���� is Not Null", "") & _
            IIf(mbytInState = 0 And gbytBilling = 1, " And A.����Ա���� is Null And A.������ is Not NULL", "") & _
            IIf(mbytInState = 0 And gbytBilling = 2, " And A.����Ա���� is Null And A.������ is Not NULL", "")
   
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ��ݣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mbytUseType = 0 Or mbytUseType = 1 Then
            If InStr(mstrPrivs, ";���в���;") = 0 And mlngUnitID > 0 Then
                If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!���˲���ID), 0, rsTmp!���˲���ID) & ",") = 0 Then
                    MsgBox "��û��Ȩ�޶�ȡ���������ĵ��ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        ElseIf mbytUseType = 2 Then
            If InStr(mstrPrivs, ";���п���;") = 0 And mlngDeptID > 0 Then
                If IIf(IsNull(rsTmp!��������ID), 0, rsTmp!��������ID) <> mlngDeptID Then
                    MsgBox "��û��Ȩ�޶�ȡ�������ҿ����ĵ��ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    '���ݺ�
    cboNO.Text = strNO

    '����
    txtPatient.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    
    '�Ա�
    txtSex.Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
    '����
    txtOld.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    
    txt������.Text = IIf(IsNull(rsTmp!������), "", rsTmp!������)
    txt������.Text = Format(IIf(IsNull(rsTmp!������), "", rsTmp!������), "0.00")
    
    '�ѱ�
    txt�ѱ�.Text = IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
    txtҽ�Ƹ���.Text = Get����ҽ�Ƹ��ʽ(rsTmp!����ID, rsTmp!��ҳID)
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsTmp!����ID)), Val(Nvl(rsTmp!��ҳID)), txtҽ�Ƹ���.Text, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    End If
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    chk�Ӱ�.Value = IIf(IsNull(rsTmp!�Ӱ��־), 0, rsTmp!�Ӱ��־)
    Call LoadPatientBaby(cboBaby, rsTmp!����ID, rsTmp!��ҳID)
    Call zlControl.CboLocate(cboBaby, rsTmp!Ӥ����, True)
        
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, Nvl(rsTmp!������), Nvl(rsTmp!��������ID, 0))
    
    '���˷�����Ϣ
    If Not IsNull(rsTmp!����ID) Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID, , True, 2)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "Ԥ��:" & Format(rsPatiMoney!Ԥ�����, "0.00") & _
            "/����:" & Format(rsPatiMoney!�������, gstrDec) & _
            "/ʣ��:" & Format(rsPatiMoney!Ԥ����� - rsPatiMoney!�������, "0.00")
        End If
    End If
    
    '-----------------------------------------------------------------
    If blnDelete Then
         '���ʵ����迼�Ǻ󱸱�,ǰ��Ĳ����ѽ�ֹ
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "Select Nvl(�۸񸸺�,���) From סԺ���ü�¼ " & _
            " Where ��¼����=2 And �����־=2 And Nvl(�ಡ�˵�,0)=0" & _
            " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
            IIf(mstrTime <> "", " And �Ǽ�ʱ��=[2]", "")
            
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        intInsure = BillExistInsure(strNO)
        If intInsure <> 0 Then
            blnDo = Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure)
        Else
            blnDo = gbytBillOpt = 2
        End If
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From סԺ���ü�¼ " & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
                    
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ," & _
            " Avg(Nvl(A.����,1)*A.����) as ����,Sum(A.��׼����) as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
            " And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0" & _
            " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " Group by A.��¼״̬,Nvl(A.�۸񸸺�,A.���),C.����,C.����,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־"
            
        '��������(ʣ��������Ϊ׼������,���ؼ���)
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        strSQL = _
            " Select A.���,A.����,A.���,A.����,A.���," & _
            " A.��������,A.���㵥λ,A.ִ�в���,A.���ӱ�־," & _
            " Sum(A.����) as ����,A.����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��" & _
            " From (" & strSQL & ") A" & _
            " Group by A.���,A.����,A.���,A.����,A.���,A.��������," & _
            " A.���㵥λ,A.����,A.ִ�в���,A.���ӱ�־" & _
            " Having Sum(A.����)<>0" & _
            " Order by A.���"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '��ȡ���ʻ��۵�(�������),ֻ��ȡʣ������,���
        '���۵����漰�󱸱�
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,C.����,C.���� as ���," & _
            " B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ,Avg(Nvl(A.����,1)*A.����) as ����," & _
            " Sum(A.��׼����) as ����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
            " Where A.��¼״̬=0 And A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
            " And A.��¼����=2 And Nvl(A.�ಡ�˵�,0)=0 And �����־=2 And A.NO=[1]" & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��¼״̬,C.����,C.����,B.����,B.���," & _
            " Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���," & _
            " C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ," & _
            " Avg(" & intSign & "*Nvl(A.����,1)*A.����) as ����," & _
            " Sum(A.��׼����) as ����,Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��, " & _
            " Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
            " And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0 And A.NO=[1]" & _
            " And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            " Group by Nvl(A.�۸񸸺�,A.���),C.����,C.����,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־" & _
            " Order by ���"
    End If
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        If gbytBilling = 2 And Not mblnPrint Then mblnPrint = True
    
        Bill.RowData(i) = rsTmp!��� '���ڼ������ʼ��������
        Bill.TextMatrix(i, 0) = rsTmp!����
        Bill.TextMatrix(i, 1) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, 2) = Format(rsTmp!ʵ�ս��, gstrDec)
        Bill.TextMatrix(i, 3) = rsTmp!ִ�в���
        Bill.TextMatrix(i, 4) = IIf(IsNull(rsTmp!��������), "", rsTmp!��������)
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "��"
        End If
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Bill.SetColColor 0, &HE7CFBA
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 3, &HE7CFBA
    Bill.Redraw = True
    
    '-------------------------------------------------------------------------------
    '��ȡ����������Ŀ
    If blnDelete Then
         '�˷ѵ����迼�Ǻ󱸱�,ǰ��Ĳ����ѽ�ֹ
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "Select Nvl(�۸񸸺�,���) From סԺ���ü�¼ " & _
            " Where ��¼����=2 And �����־=2 And Nvl(�ಡ�˵�,0)=0" & _
            " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
            IIf(mstrTime <> "", " And �Ǽ�ʱ��=[2]", "")
            
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From סԺ���ü�¼ " & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
        strSQL = _
            " Select A.���,A.����," & _
                " Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
                " Sum(A.ʵ�ս��) as ʣ��ʵ��" & _
            " From (" & _
                " Select A.��¼״̬,A.���,B.����," & _
                " Nvl(A.����,1)*A.���� as ����,A.Ӧ�ս��,A.ʵ�ս��" & _
                " From סԺ���ü�¼ A,������Ŀ B" & _
                " Where A.��¼����=2 And A.�����־=2 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.�ಡ�˵�,0)=0" & _
                    " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
                    " And A.������ĿID=B.ID" & _
                " ) A" & _
            " Group by A.���,A.���� Having Sum(����)<>0"
                    
        '��������(׼��������ʣ������,������������)
        strSQL = _
            " Select A.����,Sum(A.ʣ��Ӧ��) as Ӧ�ս��," & _
            " Sum(A.ʣ��ʵ��) as ʵ�ս��" & _
            " From (" & strSQL & ") A" & _
            " Group by A.����"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '��ȡ���ʻ��۵�(�������),ֻ��ȡδ��˲���
        '���۵����漰�󱸱�
        strSQL = _
            "Select B.����,Sum(A.Ӧ�ս��) as Ӧ�ս��," & _
            " Sum(A.ʵ�ս��) as ʵ�ս�� " & _
            " From סԺ���ü�¼ A,������Ŀ B" & _
            " Where A.��¼״̬=0 And A.��¼����=2 And A.�����־=2" & _
            " And Nvl(A.�ಡ�˵�,0)=0 And A.NO=[1] And A.������ĿID=B.ID" & _
            " Group By B.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            "Select B.����," & _
            " Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��," & _
            " Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս�� " & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ,������Ŀ B" & _
            " Where A.������ĿID=B.ID And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            " And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0 And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            " Group By B.����"
    End If
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 4 Then mshMoney.Rows = 4
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

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub


Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
'���ܣ����ݵ��������������б������
    Dim rsTmp As New ADODB.Recordset
    Dim str��Ա���� As String, strTmp As String
    Dim strSQL As String, i As Long
    Dim lng����ID As Long, lng����ID As Long
    
    Bill.Clear
    
    On Error GoTo errHandle
    

    Select Case Bill.TextMatrix(0, lngCol)
        Case "ִ�п���"
            Bill.cboStyle = DropDownAndEdit
            
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    Bill.TextMatrix(lngRow, lngCol) = ""
                    
                    lng����ID = mobjBill.����ID
                    If lng����ID = 0 Then lng����ID = Get��������ID
                    
                    lng����ID = mobjBill.����ID
                    If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
                    If lng����ID = 0 Then lng����ID = lng����ID
                    
                    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                    Select Case .Detail.ִ�п���
                        Case 0 '����ȷ
                            mrsUnit.Filter = 0
                        Case 1 '���˿���
                            mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                        Case 2 '���˲���
                            mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                        Case 3 '����Ա����
                            mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        Case 4 'ָ������
                            strSQL = "" & _
                            "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                            "   From �շ�ִ�п��� A,���ű� C" & _
                            "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
                            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                            "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                            "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                            " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, Get������Դ, lng����ID)
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
                        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
                        Case 6 '�����˿���
                           mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & .ִ�в���ID
                    End Select
                    If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                    If Not mrsUnit.EOF Then
                        For i = 1 To mrsUnit.RecordCount
                            strTmp = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                            '���˺�:28947
                            If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                            
                            'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                Bill.AddItem strTmp
                                Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                
                                '����ȱʡִ�п���
                                If lngRow = 1 Then
                                    If mrsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                ElseIf lngRow > 1 Then
                                    If mrsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� Then
                                        Bill.ListIndex = Bill.NewIndex
                                    ElseIf mrsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    End If
                                End If
                            End If
                            mrsUnit.MoveNext
                        Next
                        
                        If .Detail.ִ�п��� = 4 Then    'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                    End If
                    
                    If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetMoneyList()
'����:���ݵ�ǰ������Ŀ�����������п�
    Dim lngW As Long
    lngW = mshMoney.Width - 75
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    
    mshMoney.ColWidth(0) = lngW * 0.45
    mshMoney.ColWidth(1) = lngW * 0.55
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
    
    mshMoney.TextMatrix(0, 0) = "��Ŀ"
    mshMoney.TextMatrix(0, 1) = "���"
    mshMoney.Row = 0
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean, curTotal As Currency, curӦ��Total As Currency
    mshMoney.Redraw = False
    
    '�����ʾ
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    
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
    mshMoney.Rows = IIf(mcolMoneys.Count = 0, 2, mcolMoneys.Count + 1)
    If mshMoney.Rows < 4 Then mshMoney.Rows = 4
    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).������Ŀ
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).ʵ�ս��, gstrDec)
        curTotal = curTotal + mcolMoneys(i).ʵ�ս��
        curӦ��Total = curӦ��Total + mcolMoneys(i).Ӧ�ս��
    Next
    txtʵ��.Text = Format(curTotal, gstrDec)
    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    
    mshMoney.Redraw = True
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
    mshMoney.Rows = 4
    mshMoney.Redraw = True
    
    '20030617:��������
    'txtʵ��.Text = gstrdec
    'txtӦ��.Text = gstrdec
End Sub

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "����"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(0) = GetOrigColWidth(0) - 300
            Bill.ColWidth(3) = GetOrigColWidth(3) - 250
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(0) = GetOrigColWidth(0)
            Bill.ColWidth(3) = GetOrigColWidth(3)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Function SaveModi() As Boolean
    '���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String
    '  No_In       ������ü�¼.NO%Type,
    '  ��¼����_In ������ü�¼.��¼����%Type,
    '  ������_In   ������ü�¼.������%Type,
    '  ����ʱ��_In ������ü�¼.����ʱ��%Type,
    '  ����_In     ������ü�¼.����%Type := Null,
    '  ��Դ_In Integer:=1
  
    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo������.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL,2)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    Check�������� = True
    
    On Error GoTo errHandle
    

    '�޷����
    If txtҽ�Ƹ���.Text = "" Then Exit Function
    
    'ҽ���򹫷Ѳ���
    '����:45605
    If zlIsCheckMedicinePayMode(txtҽ�Ƹ���.Text, blnҽ��, bln����) = False Then Exit Function
    'ȷ����������
    bytType = IIf(blnҽ��, 1, 2)
    
    '��ȡ�������
    If bytType = 1 Then
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
    Else
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
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
                    IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
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
                        IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long
    Dim strInfo As String
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            For i = 1 To mobjBill.Details.Count
                For j = 1 To mobjBill.Details(i).InComes.Count
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����, _
                     mobjBill.Details(i).ժҪ & "||" & mobjBill.Details(i).����)
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
    End If
End Sub

Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, 3) = "" Then
            Checkִ�п��� = i: Exit Function
        End If
    Next
End Function

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Function Check�������() As Integer
'���ܣ���鵱ǰ���˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
'˵������Ϊ�������������۲���,�����д˼��
'���أ���һ�µķ�����,Ϊ0ʱ����
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            'סԺ���˻�סԺ���۲���,������ֻ�������������Ŀ
            If mobjBill.Details(i).Detail.������� = 1 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            '������Ժ����(ҽ������)���������۲���,������ֻ������סԺ����Ŀ
            If mobjBill.Details(i).Detail.������� = 2 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        End If
    Next
End Function

Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function
Private Function Get������Դ() As Integer
'���ܣ���ȡ��ǰ���˵���Դ(��Ϊ���Զ��������۲��˼���)
    If mrsInfo.State = 1 Then
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            Get������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            Get������Դ = 1 '���ﲡ��(ҽ������)���������۲���
        End If
    Else
        Get������Դ = 2 'ȱʡΪ2
    End If
End Function
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKIND.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKIND.Cards.��ȱʡ������
End Sub

