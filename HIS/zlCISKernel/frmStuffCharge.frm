VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmStuffCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "�������ϼ���"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStuffCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   7875
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmStuffCharge.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "�������"
            Object.ToolTipText     =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffCharge.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffCharge.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
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
      Height          =   2865
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   11805
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5010
      Width           =   11805
      Begin MSComctlLib.ImageList imgList 
         Left            =   11070
         Top             =   1980
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
               Picture         =   "frmStuffCharge.frx":1A92
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   9780
         TabIndex        =   21
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1785
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   7965
         TabIndex        =   20
         ToolTipText     =   "�ȼ���F2"
         Top             =   1785
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
         TabIndex        =   36
         ToolTipText     =   "���:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�������"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6555
            TabIndex        =   16
            Top             =   180
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   17
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
            TabIndex        =   13
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   240
            Left            =   5790
            TabIndex        =   38
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   8820
            TabIndex        =   37
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraDrawDept 
         Height          =   720
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   13575
         Begin VB.ComboBox cboִ�в��� 
            Height          =   360
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   255
            Width           =   2265
         End
         Begin VB.TextBox txt���˱�ע 
            BackColor       =   &H00E0E0E0&
            Height          =   360
            Left            =   5145
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2700
         End
         Begin VB.Label lblִ�в��� 
            Caption         =   "ִ�в���"
            Height          =   315
            Left            =   105
            TabIndex        =   53
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label lbl���˱�ע 
            Caption         =   "���˱�ע"
            Height          =   225
            Left            =   4155
            TabIndex        =   49
            Top             =   308
            Width           =   1005
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
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
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   39
         Top             =   1065
         Width           =   3675
         Begin VB.TextBox txtPreNO 
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
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1845
         End
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
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   750
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   250
            Width           =   1845
         End
         Begin VB.Label lblPreNO 
            AutoSize        =   -1  'True
            Caption         =   "����"
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
            TabIndex        =   51
            Top             =   1298
            Width           =   690
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
            TabIndex        =   41
            Top             =   818
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
            TabIndex        =   40
            Top             =   318
            Width           =   690
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   45
      TabIndex        =   24
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   11865
      Begin VB.CommandButton cmdSel 
         Caption         =   "����"
         Height          =   375
         Left            =   75
         TabIndex        =   54
         ToolTipText     =   "���:F11"
         Top             =   645
         Width           =   855
      End
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
         Caption         =   "�������ʵ�"
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
         TabIndex        =   28
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
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   9375
      TabIndex        =   23
      Top             =   855
      Width           =   2505
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   135
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "cbo��������"
         Top             =   615
         Width           =   2265
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   22
      Top             =   855
      Width           =   9345
      Begin VB.TextBox txtסԺ�� 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1290
      End
      Begin VB.TextBox txt�ѱ� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txt���ʽ 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   2085
      End
      Begin VB.TextBox txt�Ա� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   765
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   240
         Left            =   7140
         TabIndex        =   46
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   7140
         TabIndex        =   45
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   5145
         TabIndex        =   44
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
         Left            =   2400
         TabIndex        =   43
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   5385
         TabIndex        =   34
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   32
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   2370
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   3705
         TabIndex        =   30
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3105
      Left            =   15
      TabIndex        =   11
      Top             =   1920
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5477
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
      TabIndex        =   33
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmStuffCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��ڲ���
'����������������������������������������������������������������������������������������������������������������������������������������
Private mlngҽ��ID As Long  '��������ʱ��
Private mlng���ͺ� As Long  '��������ʱ��
Private mlng����ID As Long  'ȷ��Ҫ�ƷѵĲ���ID
Private mlng��ҳID As Long  'ȷ��Ҫ�Ʒѵ���ҳID
Private mstrPrivs As String
Private mstrPrivsOpt As String 'סԺ���ʲ��������Ȩ��
Private mint������Դ As Integer  '1-���ﲡ��,2-סԺ����
Private mint��¼���� As Integer  '1-�շ�(����),2-����(��/ס)
Private mstrFeeTab As String
Private mlng����ⷿID As Long
Private mbln���õǼ� As Boolean  '���Ǽ�,����ʵ�ս��
Private mlng��������ID As Long  'Ϊ��ǰ������ҽ������
Private mlng���˿���id As Long  '��Ҫ������ȷ�����ﲡ�˵Ŀ���ID
Private mblnCboNotClick As Boolean
Private mlng��������ID As Long
Private mstr����ҽ�� As String
Private mblnUnload As Boolean
Private mrsAll�������� As ADODB.Recordset

Private mbytInState As Byte  '0-ִ��,1-����,2-����(��֧��),3-ɾ��
Private mstrInNO As String  '�������ĵ��ݺ�(ִ��ʱΪ�޸�)
Private mstrOriginalNO As String  '����������ʱ,ҽ�������еĵ��ݺ�

Private mstrTime As String  '�����������ݵĵǼ�ʱ��
Private mblnDelete As Boolean  '�Ƿ����˷ѵ���(����)
Private mblnWarnCloseed As Boolean  '���˺�:����ñ��������Ĺر�
Private mblnSendMateria  As Boolean
Private mbytSendMateria As Byte '0-���ʺ󲻷�ҩ,1-�Զ���ҩ,2-��ʾ��ҩ
Private mlngִ�пⷿID As Long

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
    ��Ŀ = 1
    ��Ʒ�� = 2
    ��� = 3
    ��λ = 4
    ���� = 5
    ���� = 6
    ���� = 7
    Ӧ�ս�� = 8
    ʵ�ս�� = 9
    ��Ʒ���� = 10
    �ڲ����� = 11
    ���� = 12
End Enum

Private Enum Pan
    C2��ʾ��Ϣ = 2
End Enum

'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ʵʱ��� As Boolean
    ҽ��ȷ���������� As Boolean 'Ŀǰֻ�б���ҽ��ר��
End Type
Private MCPAR As TYPE_MedicarePAR
Private mrsDept As ADODB.Recordset
'ҽ������վ���ط��ò���
'����������������������������������������������������������������������������������������������������������������������������������������
Private mstrLike As String '����ƥ�䷽ʽ
Private mblnPay As Boolean '��ҩ�Ƿ����븶��
Private mblnTime As Boolean '����Ƿ����븶��
Private mlngPreRow As Long '��¼��ǰ��,�����ı���ʱ
'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
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
Private mintSuccess As Integer
Private Const STR_HEAD = "��,450,4;��Ŀ,2175,1;��Ʒ��,930,1;���,900,1;��λ,520,4;����,520,1;����,570,1;����,795,7;Ӧ�ս��,945,7;ʵ�ս��,945,7;��Ʒ����,1450,4;�ڲ�����,1450,4;����,520,1"
Public Function zlBillEdit(ByVal frmMain As Form, _
    ByVal bytInState As Byte, ByVal lngModule As Long, ByVal strPrivs As String, _
    Optional int��¼���� As Integer = 2, Optional ByVal strInNO As String, _
    Optional int������Դ As Integer = 2, Optional lng����ID As Long, Optional lng��ҳID As Long, _
    Optional ByVal lng��������ID As Long, Optional ByVal lng���˿���ID As Long, _
    Optional ByVal lng��������ID As Long, Optional ByVal str����ҽ�� As String, _
    Optional ByVal bln���õǼ� As Boolean, Optional ByVal str���ݵǼ�ʱ�� As String, _
    Optional ByVal lngҽ��ID As Long, Optional ByVal lng���ͺ� As Long, _
    Optional strOriginalNO As String, _
    Optional strFeeTab As Long, Optional blnDelete As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ��Ļ�༭���
    '���:bytInState:0-ִ��,1-����,2-����(��֧��),3-ɾ��
    '       strInNO:�������ĵ��ݺ�(ִ��ʱΪ�޸�)
    '       int��¼���� :1-�շ�(����),2-����(��/ס)
    '       int������Դ:1-���ﲡ��,2-סԺ����
    '       lngҽ��ID -��������ʱ��
    '       lng���ͺ�-��������ʱ��
    '       lng����ID-����ID
    '       strFeeTab:
    '       bln���õǼ�:���Ǽ�,����ʵ�ս��
    '       strOriginalNO -����������ʱ,ҽ�������еĵ��ݺ�
    '       str���ݵǼ�ʱ��:�����������ݵĵǼ�ʱ��
    '       blnDelete-�Ƿ����˷ѵ���(����)
    '����:
    '����:
    '����:���˺�
    '����:2010-12-13 17:09:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytInState = bytInState: mstrInNO = strInNO: mint������Դ = int������Դ: mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�: mlng����ID = lng����ID: mlng��ҳID = lng��ҳID: mintSuccess = 0
    mstrFeeTab = strFeeTab: mbln���õǼ� = bln���õǼ�: mlng��������ID = lng��������ID
    mlng���˿���id = lng���˿���ID
    mlng��������ID = lng��������ID:    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.pסԺ���ʲ���)
    mstrPrivs = strPrivs: mint��¼���� = int��¼����
    mstr����ҽ�� = str����ҽ��: mstrOriginalNO = strOriginalNO: mblnDelete = blnDelete
    mblnUnload = False
    Me.Show 1, frmMain
    zlBillEdit = mintSuccess > 0
 End Function
Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    Exit Sub
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then
        mobjBill.��������ID = 0
        cbo��������.Text = ""
    End If
End Sub
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
            If CheckItemHaveSub(lngMainRow) Then
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

Private Sub ShowStock(str���� As String, dbl��� As Double)
'���ܣ���ʾҩƷ�����ĵĿ��
    If InStr(1, mstrPrivs, "��ʾ���") > 0 Then
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & str���� & "]���ÿ��:" & dbl���
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & str���� & "]" & IIF(dbl��� > 0, "��", "��") & "���."
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
    Dim blnCancel As Boolean
    If SelectItem(False) = False Then
         mblnSelect = False: Exit Sub
    End If
    mblnSelect = True
    Bill.Text = mobjDetail.ID
    Call bill_KeyDown(13, 0, blnCancel)
    Bill.SetFocus
    mblnSelect = False
    If Not blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False
        Call zlCommFun.PressKey(13)
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '���ܣ�����������
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency
    Dim blnStock As Boolean, lngDoUnit As Long, strժҪ As String
    Dim lng��Ŀid As Long, str��׼��Ŀ As String, str��� As String
    Dim blnInput As Boolean, cur��� As Currency, lng���˿���ID As Long, int���� As Integer, lngOld���� As Long
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
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels(4).Text = ""
                    blnInput = True
                    If Not mblnSelect Then
                        If SelectItem(True) = False Then
                              Bill.Text = "": Bill.TxtVisible = False
                              Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    mblnSelect = False '��������ñ�־
                    Bill.TxtVisible = False '(���Ӳ���)
                    
                    'ҽ��������Ŀ�Ƿ��������
                    If mint������Դ = 2 And mint��¼���� = 2 And Not IsNull(mrsInfo!����) Then
                        If mobjDetail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                            mrsMedAudit.Filter = "��ĿID=" & mobjDetail.ID
                            If mrsMedAudit.RecordCount = 0 Then
                                MsgBox "��ǰ����δ����׼ʹ�ø���Ŀ��", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf Not IsNull(mrsMedAudit!��������) Then
                                If mrsMedAudit!�������� <= 0 Then
                                    MsgBox "��ǰ����ʹ��[" & mobjDetail.���� & "]�Ѵﵽ��׼��ʹ������" & FormatEx(mrsMedAudit!ʹ������, 5) & "��", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    '�������ò��˲�������
                    If mint������Դ = 2 And mint��¼���� = 2 Then
                        If Not CheckFeeItemLimitDept(mobjDetail.ID) Then
                            MsgBox "���������϶Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
           
                    
                    '���˿���ID
                    lng���˿���ID = mobjBill.����ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    Call ShowStock(mobjDetail.����, mobjDetail.���)
                    
                    '����֧����Ŀ��Ӧ���
                    If Not IsNull(mrsInfo!����) Then
                        
                        If zlCheck������۸����(mobjDetail.ID, Not mobjDetail.���) Then
                            '����:27286
                        Else
                            If Not ItemExistInsure(mrsInfo!����ID, mobjDetail.ID, mrsInfo!����) Then
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
                    End If
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Details(Bill.Row).ժҪ
                        End If
                    End If
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    ElseIf mint������Դ = 2 And Not IsNull(mrsInfo!����) Then
                        strժҪ = gclsInsure.GetItemInfo(mrsInfo!����, mrsInfo!����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney��ҽ�����ܷ���ժҪ
                    If mobjBill.Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Details(Bill.Row).ժҪ
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mint��¼���� = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        '���˺�:30504: and mbln���õǼ�=False
                        If curTotal > 0 And mbln���õǼ� = False Then
                            cur��� = Val(txtʵ��.Tag)
                            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID, mint������Դ)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mint��¼���� = 2 Then
                        If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        Bill.ColData(BillCol.����) = 4 '����
                        Bill.ColData(BillCol.����) = 5 '����
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        Call CheckValidity(.�շ�ϸĿID, mlngִ�пⷿID, .����, False, .Detail.����)     '��ȷ������,��������
                        '���õ���������,�������ó�������Ŀ
                    End With
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
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
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
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '�����Ϸ��Լ��
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� < 0 Then
                        'Ȩ��
                        If InStr(mstrPrivs, "���Ƹ�������") = 0 Then
                            MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.���� Then
                                MsgBox "�����������ϲ��������븺����", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint��¼���� = 2 Then
                                If Not IsNull(mrsInfo!����) Then
                                    If Not MCPAR.�������� Then
                                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    'ҩƷ�����
                    With mobjBill.Details(Bill.Row)
                        If .Detail.���� Or .Detail.��� Then
                            '������ʱ��ҩƷ�����ֹ����
                            If .���� * CSng(Bill.Text) > .Detail.��� Then
                                MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = .����: Cancel = True: Exit Sub
                            End If
                        Else
                            Set colStock = mcolStock2
                            If .���� * CSng(Bill.Text) > .Detail.��� Then
                                Call MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,���ܼ���!", vbInformation, gstrSysName)
                                    Bill.Text = .����: Cancel = True: Exit Sub
                             End If
                        End If
                    
                        dblPreTime = .����
                        .���� = Bill.Text
                        '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                        If .�������� <> 0 And .Detail.���д��� <> 0 Then
                            sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                    End With
                
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
                        
                        '���˺�:2010-07-01 10:23:11:30504:and mbln���õǼ�=False
                        If curTotal > 0 And mbln���õǼ� = False Then
                            cur��� = Val(txtʵ��.Tag)
                            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID, mint������Դ)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).���� = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                                     
                    If mint��¼���� = 2 Then
                        If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), Bill.Row)) = False Then
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
                        If mobjBill.Details(i).�������� = Bill.Row Then
                            '28136
                            '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                            With mobjBill.Details(i)
                                If .Detail.���д��� = 0 Then  '�ǹ��д���
                                    If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                    .���� = IIF(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                                ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                    .���� = IIF(Val(Bill.Text) < 0, -1, 1) * IIF(.Detail.�������� = 0, 1, .Detail.��������)
                                ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                    .���� = Val(Bill.Text) * .Detail.��������
                                Else
                                     GoTo NotCalc:
                                End If
                            End With
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
NotCalc:
                        End If
                    Next
                    
                    Call ShowMoney

                 ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
               End If
                If CheckItemHaveSub(Bill.Row) Then
                    KeyCode = 0
                    Call LocateMainItemNextRow(Bill.Row)
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
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).���� > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
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
                            '30504:and mbln���õǼ�=False
                            If curTotal > 0 And mbln���õǼ� = False Then
                                cur��� = Val(txtʵ��.Tag)
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID, mint������Դ)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
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
                            '���Ͽ����:��̬ҩ��,������ʱ�۲���ҲҪ�����
                            If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                If .���� * .���� > .Detail.��� Then
                                    MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Cancel = True
                                End If
                            Else
                                Set colStock = mcolStock2
                                If .���� * .���� > .Detail.��� Then
                                    MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Cancel = True
                                End If
                            End If
                        
                            '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                            Call CheckValidity(.�շ�ϸĿID, mlngִ�пⷿID, .����, False, .Detail.����)     '��ȷ������,��������
                            If CheckItemHaveSub(Bill.Row) Then
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
        Bill.Col = BillCol.��Ŀ
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.��Ŀ
    End If
    '����:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
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
 

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                CheckItemHaveSub = True: Exit Function
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
        Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            Call ShowStock(.Detail.����, .Detail.���)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
             '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If CheckItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
            End If
            
            '����Ƿǵ���״̬
            If mbytInState <> 2 Then
                Bill.ColData(BillCol.����) = 5
                '���������������
                Bill.ColData(BillCol.����) = 4
                Bill.ColData(BillCol.����) = 5
            End If
        End With
    End If
   
    '������δ�������,��ָ��е�����
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı��������ݴ������ʾ����
    Bill.RowData(Row) = Asc("4")
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "ִ�п���"
            Call zlControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                Bill.TextMask = "." & Bill.TextMask
                
                '�ɷ����븺��
                If Not mobjBill.Details(Bill.Row).Detail.���� Then
                    If InStr(mstrPrivs, "���Ƹ�������") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint��¼���� = 2 Then
                            If Not IsNull(mrsInfo!����) Then
                                If Not MCPAR.�������� Then
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
 

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, strDoctor As String
    Dim rsReturn As ADODB.Recordset
    Dim intIndex As Integer
    If mbytInState <> 0 Then Exit Sub
    mrsAll��������.Filter = ""
    If cbo��������.ItemData(cbo��������.ListIndex) = 0 And cbo��������.Text Like "��������*" Then
        If zlDatabase.zlShowListSelect(Me, glngSys, 1150, cbo��������, mrsAll��������, True, "", "ȱʡ,���ȼ�", rsReturn) = False Then
            mobjBill.��������ID = 0
            Exit Sub
        End If
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.State <> 1 Then Exit Sub
        If rsReturn.RecordCount = 0 Then Exit Sub
        rsReturn.MoveFirst
        If zlControl.CboLocate(cbo��������, Val(rsReturn!ID), True) = False Then
            cbo��������.RemoveItem cbo��������.ListCount - 1
            cbo��������.AddItem IIF(zlIsShowDeptCode, rsReturn!���� & "-", "") & rsReturn!����
            cbo��������.ItemData(cbo��������.ListCount - 1) = Val(Nvl(rsReturn!ID))
            intIndex = cbo��������.NewIndex
            cbo��������.AddItem "�������ҡ�"
            cbo��������.ItemData(cbo��������.ListCount - 1) = 0
            cbo��������.ListIndex = intIndex
        End If
        Exit Sub
    End If
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
                Bill.RowData(Bill.Rows - 1) = 0
            End If
        End If
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        Call GetCboIndex(cbo������, NeedName(cbo������.Text))
        If cbo������.ListIndex = -1 Then cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
End Sub

Private Sub cboִ�в���_Click()
    If mblnCboNotClick = True Then Exit Sub
    mlngִ�пⷿID = cboִ�в���.ItemData(cboִ�в���.ListIndex)
    mlng����ⷿID = Set����ⷿID(mlngִ�пⷿID)
    If mlng����ⷿID = 0 Then
        MsgBox "ע��:" & vbCrLf & "    ִ�пⷿ������ⷿδ���ö�Ӧ��ϵ,�������Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
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
    If chk�Ӱ�.value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.value = Checked
        End If
    End If
    If chk�Ӱ�.value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIF(chk�Ӱ�.value = Checked, 1, 0)
    
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
    Dim strSQL As String, i As Long, j As Long
    Dim strItems As String, str���� As String
    Dim str��λ As String, dbl���� As Double
    Dim strValues(0 To 10) As String, intR As Long
    Dim strSubTable As String, dbl���κϼ� As Double, dbl�ѽ����� As Double
    
    '����:26951
    If InStr(1, mstrPrivs, ";�������ʲ���鷢����Ŀ;") > 0 Then
        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
        CheckNegative = True: Exit Function
    End If
    
    CheckNegative = True
    If mobjBill.����ID = 0 Then Exit Function
    
    strItems = ""
    strSubTable = ""
    intR = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And mlngִ�пⷿID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        '"           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As ����, "
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select  Column_Value As �շ�ϸĿID" & _
                        " From Table(Cast(f_num2list([" & intR + 4 & "]) As ZLTOOLS.t_numlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select  Column_Value As �շ�ϸĿID " & _
                        " From Table(Cast(f_num2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_numlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                'strItems = strItems & "," & .�շ�ϸĿID & "_" & .Detail.���� & ""
                strItems = strItems & "," & .�շ�ϸĿID
            End If
        End With
    Next
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            " Select  Column_Value As �շ�ϸĿID" & _
            " From Table(Cast(f_num2list([" & intR + 4 & "]) As ZLTOOLS.t_numlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            " Select  Column_Value As �շ�ϸĿID" & _
            " From Table(Cast(f_num2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_numlist))"
        End If
    End If
    
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
  
    strSQL = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.�շ�ϸĿID,A.ִ�в���ID, " & _
    "             Nvl(Sum(Decode(A.��¼����, 2, 1, 3, 1, 0) * Nvl(A.����, 1) * A.����), 0) As ����, " & _
     "            Sum(Decode(nvL(Mod(M.��¼״̬ , 3),1),  0, 1, 1, 1, -1) * Decode(A.����id, Null, 0, 1) * Nvl(����, 1) * ����) As �������� " & _
     "     From " & mstrFeeTab & " A ,   ���˽��ʼ�¼ M,C1   " & _
     "     Where  A.����id = M.ID(+)     And A.���ʷ���=1 And A.�۸񸸺� Is Null  And A.��¼״̬<>0 " & _
     "             And A.����ID=[1] " & IIF(mint������Դ = 2, "  And Nvl(A.��ҳID,0)=[2]", "") & _
     "             and A.ִ�в���ID=[3]  " & _
    "               And A.�շ�ϸĿID=c1.�շ�ϸĿID  " & _
     "     Group By  A.�շ�ϸĿID,A.ִ�в���ID" & _
     "     Union ALL Select �շ�ϸĿID,[3]+0 as ִ�в���ID,0 as ����,0 as �������� From C1 "
    
    'strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.�շ�ϸĿID,A.ִ�в���ID,nvl(A.����,0) as ����,Sum(Nvl(A.����,1)*A.����) as ����," & _
    "           Sum(decode(A.����ID,NULL,0,1)* Nvl(A.����,1)*A.����) as �������� " & _
    " From  " & mstrFeeTab & " A " & _
    " Where A.��¼״̬<>0 And A.���ʷ���=1 and A.ִ�в���ID=[3] And A.�۸񸸺� is NULL" & _
    "           And A.����ID=[1] " & IIF(mint������Դ = 2, "  And Nvl(A.��ҳID,0)=[2]", "") & _
    "           And (A.�շ�ϸĿID+0,A.����,0,0) in (select * From C1) " & _
    " Group by A.�շ�ϸĿID,A.ִ�в���ID,A.����" & _
    " Union ALL Select �շ�ϸĿID,[3] as ִ�в���ID,����,����,�������� From C1"
    
    strSQL = "" & _
    "   Select �շ�ϸĿID,ִ�в���ID,0 as ����,Sum(����) as ����,sum(��������) as �������� " & _
    "   From (" & strSQL & ") " & _
    "   Group by �շ�ϸĿID,ִ�в���ID"
    
    On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.����ID, mobjBill.��ҳID, mlngִ�пⷿID, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And mlngִ�пⷿID <> 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID & " And ִ�в���ID=" & mlngִ�пⷿID & " And ����=" & .Detail.����
                If Not rsTmp.EOF Then
                    str��λ = .Detail.���㵥λ
                    dbl���� = Nvl(rsTmp!����, 0)
                    dbl���κϼ� = Abs(.����) * .����
                    dbl�ѽ����� = Val(Nvl(rsTmp!��������))
                    '���ܴ���������ͬ�ļ�¼
                    '����:29412
                    For j = i + 1 To mobjBill.Details.Count
                         If .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .Detail.���� And mobjBill.Details(j).Detail.���� _
                            And mobjBill.Details(j).���� < 0 Then
                            dbl���κϼ� = dbl���κϼ� + Abs(.����) * .����
                         End If
                    Next
                    '����:32106
                    If dbl���κϼ� > dbl���� - dbl�ѽ����� Then
                        Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
                        Case 0  '����
                            If dbl���κϼ� > dbl���� Then
                                str���� = Trim(cboִ�в���.Text)
                                MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                    " �����ѼƷ����� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                            End If
                        Case 1   '����
                            str���� = Trim(cboִ�в���.Text)
                            If dbl���κϼ� > dbl���� Then
                                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                        " �����ѼƷ����� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                    CheckNegative = False: Exit Function
                            End If
                            
                            If MsgBox("�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " �а������ѽᲿ��(δ��:" & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "; �ѽ�:" & FormatEx(dbl�ѽ�����, 5) & str��λ & ") ��" & vbCrLf & _
                                " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckNegative = False: Exit Function
                            End If
                        Case 2   '��ֹ
                            str���� = Trim(cboִ�в���.Text)
                            MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " �����ѼƷ����� " & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "��", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]����������Ϊ�㣬�����������", vbInformation, gstrSysName
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, strTmp As String
    Dim i As Long, j As Long, lng����ID As Long
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur��� As Currency, dbl���� As Double
    Dim cur���ն� As Currency, colStock As Collection
    Dim blnTrans As Boolean, strNos As String
    
    If mbytInState = 3 Then
        If mint��¼���� <> 1 And (False Or mlngҽ��ID <> 0) Then '������ȫ��ɾ��
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then
                    strSQL = strSQL & "," & Bill.RowData(i)
                End If
            Next
            If strSQL = "" Then
                MsgBox "������ѡ��һ��Ҫɾ���ķ��ã�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '������ѡ����
            strSQL = Mid(strSQL, 2)
            i = GetBillRows(mstrInNO, mint��¼����, mint������Դ)
            If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        Else
            '��ΪҪ����Ϊȫ�ˣ�������ʺ��������ʣ����ݽ��ʺ��Ҫ���
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint��¼����, mint������Դ)
            If j < i Then
                MsgBox "�����еĲ�����Ŀ��ǰ�Ѳ���������(������ִ�л��ѽ��ʵ���Ŀ)��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        'ҽ�����������ϴ�(ע���ж�˳��)
        If mint������Դ = 2 Then
            intInsure = BillExistInsure(mstrInNO) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) Then
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
        gcnOracle.BeginTrans: blnTrans = True
        
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        
            'ҽ�����������ϴ�
            If mint������Դ = 2 And intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ�����������ϴ�
        If mint������Դ = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ɾ��������ҽ������ʧ�ܣ��õ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0
        mintSuccess = 1
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
        If cboִ�в���.ListIndex < 0 Then
            MsgBox "����û��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.��������ID = 0 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus: Exit Sub
        End If
        
        If mobjBill.������ = "" Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��ʿ���:�жϷǷ�����
        '����ʱ����
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        If mint������Դ = 2 Then
            If Not PatiCanBilling(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), mstrPrivs, pҽ�����ѹ���) Then Exit Sub
            If zlPatiIS�����ѱ�Ŀ(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0)) Then Exit Sub
            '49501:סԺ
            If zlIsAllowFeeChange(mrsInfo!����ID, Val(Nvl(mrsInfo!��ҳID))) = False Then Exit Sub
        End If
        
        '���˺� ����:?? ����:2010-01-07 10:37:09
        If zlCheck����ҽ��(Val(Nvl(mrsInfo!����))) = False Then Exit Sub
        
        
        
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
         dbl���� = 0
        For i = 1 To mobjBill.Details.Count
           '27467,52828
            If mobjBill.Details(i).���� <> 0 And dbl���� = 0 Then
                dbl���� = mobjBill.Details(i).����
            End If
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
             '8407
            End If
        Next
        '27467,52828
        If mbytInState = 0 And Round(dbl����, 7) = 0 Then
            MsgBox "����������Ҫ��һ����Ϊ�������,���飡", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
                
        '���ʷ��౨��
        If mint��¼���� = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '���ݷ���
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                'ˢ�²��˷���״��
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIF(mint������Դ = 1, 0, mlng��ҳID), mcurModiMoney)
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
                cur��� = Val(txtʵ��.Tag)
                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID, mint������Դ)
                
                If mbln���õǼ� = False Then    '30504
                    For i = 1 To mobjBill.Details.Count
                        mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, cur���ն� - mcurModiMoney, curTotal, IIF(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, mintWarn)
                        If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        If mint��¼���� = 2 And Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� Then
            If gclsInsure.CheckItem(mrsInfo!����, 1, 2, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text))) = False Then
                Exit Sub
            End If
        End If
        
              
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.���� Or .Detail.���) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .Detail.���� = mobjBill.Details(j).Detail.���� Then
                            MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        'ҩƷ�����(�������ֹʱ�����ʱ��ҩƷ)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = mcolStock2
                If .Detail.���� Or .Detail.��� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, mlngִ�пⷿID, .Detail.����)
                    If dblTotal > .Detail.��� Then
                        MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIF(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                ElseIf colStock("_" & mlngִ�пⷿID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, mlngִ�пⷿID, .Detail.����)
                    If dblTotal > .Detail.��� Then
                        MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIF(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End With
        Next
        
        '����������ϵ����Ч��
 
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, mlngִ�пⷿID, .Detail.����)
                If Not CheckValidity(.�շ�ϸĿID, mlngִ�пⷿID, dblTotal, True, .Detail.����) Then Exit Sub
            End With
        Next
        If InStr(mstrPrivs, ";ҩƷ��ҩ;") = 0 Then mblnSendMateria = False
        
        If mblnSendMateria And mbytSendMateria = 2 Then
            If MsgBox("������ɺ��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        '�����˷Ѽ��
        If mint��¼���� = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        'ˢ�������鿨
        If mint������Դ = 1 And mint��¼���� = 2 And gbln������֤ Then
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, curTotal) Then Exit Sub
            End If
        End If
        '74231,Ƚ����,2014-7-21,��Ŀ�����������շѻ�������
        If gobjSquareCard Is Nothing Then
            If mint������Դ = 1 And gbln�������������� Then
                If MsgBox("ע�⣺" & vbCrLf & "      ҽ�ƿ�������zl9CardSquare��δ���������������󽫲��ܽ����շѻ������ˣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not SaveBill(strNos) Then Exit Sub
        
        '74231,Ƚ����,2014-7-21,��Ŀ�����������շѻ�������
        If mint������Դ = 1 And gbln�������������� And strNos <> "" Then
            If Not gobjSquareCard Is Nothing Then
                Call gobjSquareCard.zlSquareAffirm(Me, pҽ�����ѹ���, mstrPrivs, mlng����ID, , , mint��¼����, strNos)
            End If
        End If
        
        mintSuccess = mintSuccess + 1
        '���˺�:��ӡ��ҩ��:25490
        If mblnSendMateria Then
            If InStr(1, mstrPrivs, ";��ҩ�嵥��ӡ;") > 0 Then
                If MsgBox("����""" & mobjBill.NO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        End If
        
        If mstrInNO <> "" Or mstrOriginalNO <> "" Then
            gblnOK = True: Unload Me: Exit Sub
        Else
            txtPreNO.Text = mobjBill.NO
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
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str������ As String, ByVal str�������� As String, _
    Optional ByVal lngRow As Long) As ADODB.Recordset
    '���ܣ����ݵ��ݶ������ݴ���һ����ϸ��¼����Ϣ(���ۼ۵�λ)
    '�ֶΣ�����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������
    '������intPage=ָ���ĵ���,lngRow=ָ�����У���ָ��ʱ�������е��ݵ�������
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset
    
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    '69788:���ϴ�,2014-6-5,�����������ֶδ�С����20��Ϊ100
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "��������", adVarChar, 50, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    If lngRow = 0 Then
        intB = 1
        intE = objBill.Details.Count
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl���� = 0: curʵ�� = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                rsTmp!����ID = objBill.����ID
                rsTmp!��ҳID = objBill.��ҳID
                
                rsTmp!�շ���� = .�շ����
                rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                
                
                For j = 1 To .InComes.Count
                    dbl���� = dbl���� + .InComes(j).��׼����
                    curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                Next
                rsTmp!���� = IIF(.���� = 0, 1, .����) * .����
                rsTmp!���� = Format(dbl����, gstrDecPrice)
                rsTmp!ʵ�ս�� = Format(curʵ��, gstrDec)
                
                rsTmp!������ = str������
                rsTmp!�������� = str��������
            Else
                For j = 1 To .InComes.Count
                    dbl���� = dbl���� + .InComes(j).��׼����
                    curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                Next
                rsTmp!���� = rsTmp!���� + IIF(.���� = 0, 1, .����) * .����
                rsTmp!���� = Format((rsTmp!���� + Format(dbl����, gstrDecPrice)) / 2, gstrDecPrice)
                rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + Format(curʵ��, gstrDec)
            End If
            
            rsTmp.Update
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub cmdSel_Click()
    Dim rsSel As ADODB.Recordset
    If frmStuffInSel.zlSelect(Me, 1250, mstrPrivs, mlng����ⷿID, rsSel) = False Then
        Bill.SetFocus
        Exit Sub
    End If
    Call LoadSelBillData(rsSel)
    Bill.SetFocus
End Sub

Private Sub LoadSelBillData(ByVal rsSel As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2010-12-16 17:04:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnALLIgnore As Boolean '���������ظ���
    Dim bln�ظ� As Boolean, i As Long
    Dim objDetail As Detail, IntMsg As VbMsgBoxResult
    '���˿��һ򿪵�����ID
    With rsSel
        blnALLIgnore = False
        If .RecordCount > 0 Then .MoveFirst
        
        Do While Not .EOF
            bln�ظ� = False
            For i = 1 To mobjBill.Details.Count
                If mobjBill.Details(i).Detail.ID = Val(Nvl(!�շ���ĿID)) _
                    And mobjBill.Details(i).Detail.���� = Val(Nvl(!����)) Then
                    If blnALLIgnore = False Then
                        IntMsg = MsgBox("ע��:" & vbCrLf & "  �ڵ�" & i & " & �����Ѿ��������������ϡ�" & Nvl(!��������) & "��,�Ƿ���Դ���?" & _
                                            "���ǡ���ʾ�����Ե�ǰ��������!" & vbCrLf & _
                                            "���񡻱�ʾ���������Ѿ��ڵ����д��ڵ����Σ�" & vbCrLf & _
                                            "��ȡ������ʾ�˳�����ѡ�������!", vbYesNoCancel + vbQuestion + vbDefaultButton3, gstrSysName)
                        If IntMsg = vbCancel Then Exit Sub
                        If IntMsg = vbNo Then
                            blnALLIgnore = True
                        End If
                    End If
                    bln�ظ� = True
                End If
            Next
            If bln�ظ� = False Then
                Set objDetail = GetInputDetail(Val(!�շ���ĿID))
                objDetail.���� = Val(Nvl(!����))
                objDetail.��Ʒ���� = Trim(Nvl(!��Ʒ����))
                objDetail.�ڲ����� = Trim(Nvl(!�ڲ�����))
                objDetail.��� = Val(Nvl(!���ÿ��))
                
                '��������
                Call SetDetail(objDetail, mobjBill.Details.Count + 1, mlng����ⷿID)
                mobjBill.Details(mobjBill.Details.Count).���� = Val(Nvl(!����))
                Call CalcMoneys(mobjBill.Details.Count)
            End If
            .MoveNext
        Loop
    End With
    Bill.ClearBill: Call SetColNum
    Bill.Rows = mobjBill.Details.Count + 1
    '����б༭����������ɫ
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    
    Call ShowDetails
    Call ShowMoney
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            Bill.RowData(i) = Asc("4") '���⴦��
            
        End With
    Next
    Call SetColNum
    If Bill.Enabled Then Bill.SetFocus
    
End Sub

Private Sub Form_Activate()
    If mblnUnload Then
        Unload Me: Exit Sub
    End If
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
    
    mblnWarnCloseed = False
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    
    gblnOK = False
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    mstrFeeTab = IIF(mint������Դ = 2, "סԺ���ü�¼", "������ü�¼")
    Call InitLocPar
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Then
        If Not InitData Then
            mblnUnload = True: Exit Sub
        End If
    End If
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then
        If Not ReadBill(mstrInNO, mbytInState = 3) Then
            mblnUnload = True: Exit Sub
        End If
    Else
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportStuffBill(mint������Դ, mstrInNO, mint��¼����, mlng����ⷿID)
            If mobjBill.NO = "" Then
                MsgBox "������ȷ��ȡ�Ʒѵ��ݵ����ݣ�", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '����б༭����������ɫ
                Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
                Bill.SetColColor BillCol.����, &HE7CFBA
                Bill.SetColColor BillCol.����, &HE0E0E0
                Bill.SetColColor BillCol.����, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                               
                
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                If mint��¼���� = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '�ڶ�ȡ����ǰȡ
                End If
                
                '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
                Call GetPatient(mlng����ID, mlng��ҳID)
                If mrsInfo.State = 0 Then
                    If Not mblnWarnCloseed Then
                        MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                    Unload Me: Exit Sub
                End If
                
                Call FindCboIndex(cbo��������, mobjBill.��������ID, False)
                Call GetCboIndex(cbo������, mobjBill.������)
                Call zlControl.CboLocate(cboBaby, mobjBill.Ӥ����, True)
                
                If gbln��������ۿ� Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '�������:�޸�ʱ���Ͻ�Ҫ�˻صĿ��
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc("4") '���⴦��
                        .Detail.��� = .Detail.��� + .���� * .����
                    End With
                Next
                Call SetColNum
            End If
        Else
            '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
            Call GetPatient(mlng����ID, mlng��ҳID)
            If mrsInfo.State = 0 Then
                If Not mblnWarnCloseed Then
                    MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                End If
                mblnUnload = True: Exit Sub
            End If
            If Not IsNull(mrsInfo!����) Then
                MCPAR.�������� = gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����)
                MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mrsInfo!����ID, mrsInfo!����)
                MCPAR.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, mrsInfo!����ID, mrsInfo!����)
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
    mstrOriginalNO = ""
    
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

 
Private Sub picAppend_Resize()
    err = 0: On Error Resume Next
    With picAppend
        txt���˱�ע.Width = .ScaleWidth - txt���˱�ע.Left - 100
    End With
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Not gbln����ƥ�䷽ʽ�л� Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("���뷽ʽ", IIF(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIF(sta.Panels("WB").Bevel = sbrInset, 1, 0)))
        
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
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
            .ColData(BillCol.����) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
        End If
        '����б༭����������ɫ
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        
        .TextMatrix(Row, BillCol.��) = Row
        
        '����ط��ֶ����ò�ִ��
        If Row > 0 And .ColData(BillCol.��Ŀ) <> 5 And Me.Visible And Not mblnNewRow Then
            'Call zlCommFun.PressKey(13)
            
        End If
    End With
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
     '���˺� ����:27378 ����:2010-01-27 16:20:02
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If cbo��������.Locked Then Exit Sub
    If mrsAll�������� Is Nothing Then Exit Sub
    
    If zlSelectDept(Me, 1150, cbo��������, mrsAll��������, cbo��������.Text, True, , , True) = False Then
        mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = cbo������.Text
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then
                Call zlControl.CboSetIndex(cbo������.Hwnd, -1)
            Else
                zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
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
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
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
            If Not gbln����ƥ�䷽ʽ�л� Then Exit Sub
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF11
            If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
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
    Dim strOperDoc As String
    
    On Error GoTo errH
    
    Set mcolStock2 = GetStockCheck(1)
    
    '��������
    strSQL = "Select ��������ID,����ҽ�� From ����ҽ����¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
    If Not rsTmp.EOF Then
        mlng��������ID = Nvl(rsTmp!��������id, 0)
        mstr����ҽ�� = Nvl(rsTmp!����ҽ��)
    End If
    If mlng��������ID = 0 Or mstr����ҽ�� = "" Then
        MsgBox "û�з���Դҽ����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = _
    "   Select A.ID, A.����, A.����, A.����, 0 As ȱʡ, B.��������, D.���ȼ�" & vbNewLine & _
    "   From ���ű� A, ��������˵�� B," & vbNewLine & _
    "       (Select ����id, Max(Decode(�������, 2, 1, 2)) As ���ȼ� From ��������˵�� Where ������� <> 0 Group By ����id) D" & vbNewLine & _
    "   Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And A.ID = B.����id" & vbNewLine & _
    "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
    "       And B.����id = D.����id And (B.������� IN(1,2,3) AND B.�������� IN('�ٴ�','����') Or b.��������='����')" & vbNewLine & _
    "Order By ���ȼ�,����"
    Set mrsAll�������� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '70434:������,2014-02-12,�������������б���������ҽ������
    strOperDoc = Getҽ����������(mlngҽ��ID, "����ҽ������")
    
    If mbln���õǼ� Then
        '��Ϊ��ǰѡ���ҽ������
        strSQL = "(Select ID,����,����,���� From ���ű� Where ID=[1]"
    Else
        '��Ϊ��ǰѡ���ҽ�����һ�������
        strSQL = "(Select ID,����,����,���� From ���ű� Where ID IN([1],[2])"
    End If
    
    If strOperDoc <> "" Then
        strSQL = strSQL & " Union " & _
                "Select ID,����,����,���� From ���ű� Where ����=[3]"
    End If
    strSQL = strSQL & ") Order By ����"
    
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��������ID, mlng��������ID, strOperDoc)
    cboִ�в���.Clear
    mblnCboNotClick = True
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cbo��������.AddItem IIF(zlIsShowDeptCode, mrsDept!���� & "-", "") & mrsDept!����
            cbo��������.ItemData(cbo��������.ListCount - 1) = mrsDept!ID
            If mrsDept!ID = mlng��������ID Then
                cbo��������.ListIndex = cbo��������.NewIndex
                cboִ�в���.AddItem IIF(zlIsShowDeptCode, mrsDept!���� & "-", "") & mrsDept!����
                cboִ�в���.ItemData(cboִ�в���.NewIndex) = Val(Nvl(mrsDept!ID))
                cboִ�в���.ListIndex = cboִ�в���.NewIndex
                mlngִ�пⷿID = cboִ�в���.ItemData(cboִ�в���.NewIndex)
            End If
            mrsDept.MoveNext
        Next
        cbo��������.AddItem "�������ҡ�"
        cbo��������.ItemData(cbo��������.ListCount - 1) = 0
        If cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    Else
        MsgBox "����ȷ���������ң����ȵ����Ź��������á�", vbInformation, gstrSysName
        mblnCboNotClick = False
        Exit Function
    End If
    mblnCboNotClick = False
    mlng����ⷿID = Set����ⷿID(mlngִ�пⷿID)
    If mlng����ⷿID = 0 Then
        MsgBox "ע��:" & vbCrLf & "    ִ�пⷿ������ⷿδ���ö�Ӧ��ϵ,�������Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If cboִ�в���.ListCount = 0 Then
        MsgBox "����ȷ��ִ�в��ţ����ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    End If
    mlng�������ID = ExistIOClass(IIF(mint��¼���� = 1, 40, 41))
    If mlng�������ID = 0 Then
        MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
        Exit Function
    End If
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
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.��Ŀ
        .PrimaryCol = BillCol.��Ŀ
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.��) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(BillCol.��) = 5
            .ColData(BillCol.��Ŀ) = 1 '��Ŀ����,��Ť��ѡ
            .ColData(BillCol.����) = 4 '��/������
            '���˺�:27990 2010-02-23 12:04:37
            .ColData(BillCol.��Ʒ��) = 5 '�������
            .ColData(BillCol.���) = 5 '�������
            .ColData(BillCol.��λ) = 5 '��λ����
            .ColData(BillCol.����) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.Ӧ�ս��) = 5 'Ӧ�ս������
            .ColData(BillCol.ʵ�ս��) = 5 'ʵ�ս������
            .ColData(BillCol.����) = 5 '����ȱʡ����
            .ColData(BillCol.��Ʒ����) = 5
            .ColData(BillCol.�ڲ�����) = 5
        End If
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gbytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    
    Call SetMoneyList

    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln����ƥ�䷽ʽ�л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln����ƥ�䷽ʽ�л�
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���
        i = Val(zlDatabase.GetPara("���뷽ʽ"))
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
    If mint��¼���� = 1 Then
        lblTitle.Caption = gstrUnitName & "�����շѵ�"
    ElseIf mint��¼���� = 2 Then
        lblTitle.Caption = gstrUnitName & "���˼��ʵ�"
    End If
    txtӦ��.Text = gstrDec: txtʵ��.Text = gstrDec
    
    Select Case mbytInState
        Case 0 'ִ��
            Call SetShowCol
        Case 1 '����
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraDrawDept.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdSel.Visible = False
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Case 3 '����
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraDrawDept.Enabled = False
            fraAppend.Enabled = False
            cmdSel.Visible = False
            '��ʱ��֧�ֲ���ɾ��
            If mint��¼���� <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
    
    If mbytInState <> 0 Then
        lblPreNO.Visible = False: txtPreNO.Visible = False
        lblӦ��.Top = lblӦ��.Top + txtPreNO.Height / 2
        txtӦ��.Top = txtӦ��.Top + txtPreNO.Height / 2
        lblʵ��.Top = lblʵ��.Top + txtPreNO.Height * 0.75
        txtʵ��.Top = txtʵ��.Top + txtPreNO.Height * 0.75
    End If
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
    
    mblnWarnCloseed = False
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
        " A.����ID,Nvl(B.��ҳID,0) ��ҳID,To_Number(Nvl(B.��ǰ����ID,[3])) as ����ID," & _
        " Nvl(B.��Ժ����ID,[3]) as ����ID,B.��Ժ����,B.��Ժ����," & _
        " A.�����,B.סԺ��,B.��Ժ���� as ����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
        " A.������," & IIF(mint������Դ = 2 And mint��¼���� = 2, "Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,", "A.������,") & _
        " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Y.���� as ������,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���," & _
        " zl_PatiDayCharge(A.����ID) as ���ն�,Nvl(B.����,A.����) as ����,Nvl(B.��������,0) as ��������,b.��˱�־,B.��ע as ���˱�ע" & _
        " From ������Ϣ A,������ҳ B,������� X,ҽ�Ƹ��ʽ Y" & _
        " Where A.����ID=B.����ID(+) And A.����ID=X.����ID(+) And X.����(+) = " & IIF(mint������Դ = 1, 1, 2) & strSQL & _
        " And A.����ID=[1] And B.��ҳID(+)=[2]   And A.ҽ�Ƹ��ʽ=Y.����(+)"
        
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, mlng���˿���id)
    If Not mrsInfo.EOF Then
        If Not IsNull(mrsInfo!����) Then
            txt����.ForeColor = vbRed
        End If
        
        '�������ﻮ������Ҫ���������
        If mint��¼���� = 2 Then
            If mint������Դ = 2 Then
                '49501:סԺ
                If zlIsAllowFeeChange(mrsInfo!����ID, Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt����.Text = "":
                    mlng����ID = 0
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
            End If
            'ˢ�²��˷���״��
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIF(mint������Դ = 1, 0, mlng��ҳID), mcurModiMoney)
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
            strSQL = "Select Nvl(��������,1) as ��������," & _
                " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
                " Where ���ò���=[2] And " & IIF(mint������Դ = 1, "Nvl(����ID,0)=0", "����ID=[1]")
            Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!����ID, 0)), CStr(Nvl(mrsInfo!���ò���)))
            
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------
            '���˺�:26952
            Dim cur��� As Currency, curItemMoney As Currency, curTotal As Double
            curItemMoney = 0
            curTotal = GetBillTotal(mobjBill)
            cur��� = Val(txtʵ��.Tag)
            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(mrsInfo!����ID, mint������Դ)
            
            If mbln���õǼ� = False Then    '30504
            
                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!����, cur���, Val(Nvl(mrsInfo!���ն�)) - mcurModiMoney, curTotal, _
                     Nvl(mrsInfo!������, 0), "", "", mstrWarn, mintWarn, , True)
                '����:0;û�б���,����
                '     1:������ʾ���û�ѡ�����
                '     2:������ʾ���û�ѡ���ж�
                '     3:������ʾ�����ж�
                '     4:ǿ�Ƽ��ʱ���,����
                '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
                If mbytWarn = 2 Or mbytWarn = 3 Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt����.Text = "":
                    mlng����ID = 0
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                If mrsWarn.EOF Then mrsWarn.Close '���ں���״̬�ж�
            End If
        End If
                            
        'סԺ���ʲŴ��������
        If mint������Դ = 2 Then
            '�������
            If Not IsNull(mrsInfo!����) Then
                chk����.value = 0: chk����.Visible = True
            Else
                chk����.value = 0: chk����.Visible = False
            End If
            
            '����ʱ��
            If Not IsNull(mrsInfo!��Ժ����) Then
                txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
        
        '��ʾ������Ϣ
        txt����.Text = Nvl(mrsInfo!����)
        txt�Ա�.Text = Nvl(mrsInfo!�Ա�)
        txt����.Text = Nvl(mrsInfo!����)
        txt�ѱ�.Text = Nvl(mrsInfo!�ѱ�)
        txt���ʽ.Text = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
        txt���ʽ.Tag = Nvl(mrsInfo!������, 0) '��Ҫ��дΪ��
        txt����.Text = Nvl(mrsInfo!����)
        
        '���˺� ����:26953 ����:2009-12-25 15:21:47
        txt���˱�ע = Nvl(mrsInfo!���˱�ע)
        If mint������Դ = 1 Then
            lblסԺ��.Caption = "�����"
            txtסԺ��.Text = Nvl(mrsInfo!�����)
        Else
            lblסԺ��.Caption = "סԺ��"
            txtסԺ��.Text = Nvl(mrsInfo!סԺ��)
        End If
        
        txt������.Text = Nvl(mrsInfo!������)
        txt������.Text = Format(Nvl(mrsInfo!������), "0.00")
        
        With mobjBill
            .����ID = Nvl(mrsInfo!����ID, 0)
            .��ҳID = Nvl(mrsInfo!��ҳID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .���� = Nvl(mrsInfo!����)
            .��ʶ�� = IIF(mint������Դ = 1, Nvl(mrsInfo!�����), Nvl(mrsInfo!סԺ��))
            .���� = Nvl(mrsInfo!����)
            .�Ա� = Nvl(mrsInfo!�Ա�)
            .���� = Nvl(mrsInfo!����)
            .�ѱ� = Nvl(mrsInfo!�ѱ�)
        End With
        
        '�ڵ�һ�ν���ʱ��ȡ��������������Ŀ��Ϣ
        If Not Visible And mint������Դ = 2 And mint��¼���� = 2 And mbytInState = 0 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
                If CheckItemHaveSub(i) Then                          '����������
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
        
    Dim dblAllTime As Double, dblPrice As Double, dblPriceSingle As Double
    
    On Error GoTo errH
    
    gstrSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        "       And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL)) " & _
        "       And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            '�Ȼ�ȡ����Ա��ǰ����ı�۽��
            If .Detail.��� Then
                '����ҩƷʱ��(�����򲻷���)
                '��Ȼ�м�¼(�������Ŀʱ���ж�)
                dblAllTime = .���� * .����
                If dblAllTime <> 0 Then
                    dblPrice = Getʱ�۲���Ӧ�ս��(mlng����ⷿID, .�շ�ϸĿID, .Detail.����, dblAllTime, gstrDec, dblPriceSingle)
                    If dblAllTime <> 0 Then
                        '����δ�ֽ����
                        MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        dblMoney = 0
                    Else
                        'ע�⣺���������ֻ�ܱ���4λС��,�Ҳ���������,������Ҫ�ֹ�����;�����������ڼ��㾫������������
                        dblAllTime = .���� * .����
                        dblMoney = IIF(dblPriceSingle = 0, Format(dblPrice / dblAllTime, gstrDecPrice), dblPriceSingle) '�������ǰ��ۼ۵�λ
                    End If
                Else
                    dblMoney = 0
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
                    .��׼���� = Format(dblMoney, gstrDecPrice)
                Else
                    .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gstrDecPrice)
                End If
                'Ӧ�ս��=���� * ���� * ����
                If mobjBill.Details(lngRow).Detail.��� Then
                    .Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                Else
                    .Ӧ�ս�� = .��׼���� * mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
                End If
                
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100
                    .Ӧ�ս�� = .Ӧ�ս�� * (1 + dbl�Ӱ�Ӽ���)
                End If
                
                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
                
                dblAllTime = mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
                If mbln���õǼ� Or .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                Else
                    If mobjBill.Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Then
                        .ʵ�ս�� = .Ӧ�ս��
                    Else
                        .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, _
                            mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, _
                            dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))
                    End If
                End If
                
                '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details(lngRow).ժҪ & "||" & dblAllTime)
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
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "��Ʒ��"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ��
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��λ"
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIF(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gstrDecPrice)
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
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "��Ʒ����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ����
            Case "�ڲ�����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�ڲ�����
                
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
 
Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = 1
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
                Bill.RowData(lngRow) = Asc("4")
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIF(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                dblTime = 1
            End If
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .���, .���㵥λ, intPay, dblTime, 0, mlngִ�пⷿID, ""
        End With
    Else '��������Ѿ�����,���޸�
        dblTime = 1
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
            .ִ�в���ID = mlngִ�пⷿID
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    strSQL = "Select count(����ID) as NUM from �շѴ�����Ŀ where ����ID=[1]"
    On Error GoTo errH
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
    chk�Ӱ�.value = IIF(OverTime, 1, 0)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
    Call LoadPatientBaby(cboBaby, 0, 0)
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
        .�Ӱ��־ = chk�Ӱ�.value
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

Private Function SaveBill(Optional ByRef strNos As String) As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim int�к� As Integer, int��� As Integer, int�۸񸸺� As Integer
    Dim dbl���� As Double, dbl���� As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int���� As Integer, bln�ϴ� As Boolean
    Dim strSQL As String, strStuffDept As String '��¼���Ϸ��ϲ���
    Dim strDeptIDs As String, str���ܺ� As String
    Dim cllProExeute As New Collection, varTemp As Variant
    Dim rsTmp As ADODB.Recordset
    Dim blnTrans As Boolean
    
    strNos = ""
    If mstrOriginalNO = "" Then
        If mint��¼���� = 1 Then
            mobjBill.NO = zlDatabase.GetNextNo(13)
        Else
            mobjBill.NO = zlDatabase.GetNextNo(14)
        End If
    Else
        mobjBill.NO = mstrOriginalNO
    End If
    mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
    
    int��� = 0
    arrSQL = Array()
    Set cllProExeute = New Collection
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                
                '��������
                With mobjBill
                    If mint������Դ = 2 Then
                        gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                            IIF(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                            ZVal(.����ID) & "," & ZVal(.����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                    Else
                        If mint��¼���� = 2 Then
                            gstrSQL = "zl_������ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & _
                                IIF(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "'," & _
                                "'" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & "," & _
                                ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
                        Else
                            gstrSQL = "zl_���ﻮ�ۼ�¼_Insert('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                                IIF(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & ",'" & IIF(Val(txt���ʽ.Tag) = 0, "", txt���ʽ.Tag) & "','" & .���� & "'," & _
                                "'" & .�Ա� & "','" & .���� & "','" & .�ѱ� & "'," & .�Ӱ��־ & "," & _
                                  ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
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
                        If mobjBill.Details(.���).�������� = 0 Then
                            For i = .��� + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = .��� Then
                                    mobjBill.Details(i).�������� = int���
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                    
                    If mint������Դ = 2 Then
                        gstrSQL = gstrSQL & IIF(.������Ŀ��, 1, 0) & "," & ZVal(.���մ���ID) & ",'" & .���ձ��� & "',"
                    ElseIf mint��¼���� = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl���� = .����
                    gstrSQL = gstrSQL & IIF(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & mlngִ�пⷿID & ","
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    dbl���� = .��׼����
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
                If int���� = 0 Then bln�ϴ� = True 'ֻҪ���ڲ��ǻ��۵���Ҫ�ϴ�
                
                '�ռ����Ϸ��ϲ���,�Ա��Զ�����,���ﲡ�˽�����ʱ(����Ϊ����ʱ����),סԺ����ֻ�м���
                'mint������Դ :1-���ﲡ��,2-סԺ����
                'mint��¼���� :1-�շ�(����),2-����(��/ס)
                
                With mobjBillDetail
                    If (mint������Դ = 1 And mint��¼���� = 2 And gbln�����Զ����� Or mint������Դ = 2 And gblnסԺ�Զ�����) And int���� = 0 Then
                        strStuffDept = "," & mlngִ�пⷿID
                    End If
                End With
                
                If mint������Դ = 2 Then
                    gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                        "0," & mlng�������ID & "," & _
                        "NULL,'" & mobjBillDetail.ժҪ & "'," & chk����.value & "," & ZVal(mlngҽ��ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.���� & "'," & _
                        IIF(mobjBill.��������ID = mlng��������ID, "1", "0") & "," & mlng��������ID & ",NULL,-1,1," & mobjBillDetail.Detail.���� & ")"
                        '    ҽ�����ٴ�����_In Number := 0,
                        '    ��ҩ����id_In     ҩƷ�շ���¼.�Է�����id%Type := Null,
                        '    ��ҩ��̬_In       סԺ���ü�¼.����%Type := Null,
                        '    ҽ��С��id_In     סԺ���ü�¼.ҽ��С��id%Type := -1,
                        '    ��������_In       Number := 0,
                        '    ����_In           ҩƷ�շ���¼.����%Type := Null
                Else
                    If mint��¼���� = 2 Then
                        gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                              mlng�������ID & "," & _
                            "NULL,'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ",NULL,NULL,NULL,NULL,NULL,1,NULL,1," & mobjBillDetail.Detail.���� & ")"
                            '    Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
                            '    ����_In       ҩƷ�շ���¼.����%Type := Null,
                            '    �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
                            '    ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
                            '    �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
                            '    �����־_In   ������ü�¼.�����־%Type := 1,
                            '    ��ҩ��̬_In   ������ü�¼.����%Type := Null,
                            '    ��������_In   Number := 0,
                            '    ����_In       ҩƷ�շ���¼.����%Type := Null
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.���� & "'," & _
                             mlng�������ID & "," & _
                            "'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ",NULL,NULL,NULL,NULL,NULL,1,NULL,NULL,NULL,NULL,NULL,1," & mobjBillDetail.Detail.���� & "  )"
                            'Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
                            '����_In       ҩƷ�շ���¼.����%Type := Null,
                            '�÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
                            '��Ч_In       ҩƷ�շ���¼.����%Type := Null,
                            '�Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
                            '������Դ_In   Number := 1,
                            '���ձ���_In   ������ü�¼.���ձ���%Type := Null,
                            '��������_In   ������ü�¼.��������%Type := Null,
                            '������Ŀ��_In ������ü�¼.������Ŀ��%Type := Null,
                            '���մ���id_In ������ü�¼.���մ���id%Type := Null,
                            '��ҩ��̬_In   ������ü�¼.����%Type := Null,
                            '��������_In   Number := 0,
                            '����_In       ҩƷ�շ���¼.����%Type := Null
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    If mstrOriginalNO = "" Then
        '����ҽ��Ժ�ӷ���
        gstrSQL = "ZL_����ҽ������_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mint��¼���� & ",'" & mobjBill.NO & "')"
    Else
        '��������
        gstrSQL = "ZL_����ҽ������_�Ʒ�(" & mlngҽ��ID & "," & mlng���ͺ� & ")"
    End If
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
        gcnOracle.BeginTrans: blnTrans = True
        
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '-----------------------------------------------------------------------
            'ִ���Զ�����
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                varTemp = Split(strStuffDept, ",")
                For i = 0 To UBound(varTemp)
                    '69902:������,2014-02-09,ֻ��ͬ��������һ�µ�ִ�п�����Ŀ�����Զ�����
                    If Val(varTemp(i)) = Val(cbo��������.ItemData(cbo��������.ListIndex)) Then
                        strSQL = "zl_�����շ���¼_��������(" & Val(varTemp(i)) & ",25,'" & mobjBill.NO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                        zlAddArray cllProExeute, strSQL
                    End If
                Next
            End If
            'ִ�з�ҩ�ͷ���
            zlExecuteProcedureArrAy cllProExeute, Me.Caption, False, False
            '-----------------------------------------------------------------------
            
            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�
            If mint������Դ = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If gclsInsure.GetCapability(support�����ϴ�, mlng����ID, mrsInfo!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mrsInfo!����) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ���ӿ�
        '1.ҽ�����������ϴ�
        If mint������Դ = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ҽ������ʧ��,�õ��ݵķ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.����ʵʱ�ϴ�
        If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
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
    '74231,Ƚ����,2014-7-21,��Ŀ�����������շѻ�������
    strNos = mobjBill.NO
    
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
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
        If mint��¼���� = 1 Or (mint��¼���� = 2 And mint������Դ = 1) Then
            blnNOMoved = zlDatabase.NOMoved("������ü�¼", strNO, "��¼����=", mint��¼����)
        Else
            blnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, "��¼����=", mint��¼����)
        End If
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    If mstrFeeTab = "סԺ���ü�¼" Then
        strSQL = _
        " Select A.����ID,Nvl(A.��ҳID,0) ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.����,A.��ʶ��," & _
        "           A.���˲���ID,A.��������ID,A.�Ӱ��־,A.Ӥ����,A.������,A.������,A.����Ա����," & _
        "           A.��������ID,A.ִ�в���ID," & IIF(zlIsShowDeptCode, "C.����||'-'||", "") & "C.���� as ��������," & IIF(zlIsShowDeptCode, "C.����||'-'||", "") & "C.����  as ִ�в���,A.����ʱ��," & _
        "            B.ҽ�Ƹ��ʽ,B.������,B.������,A.�Ƿ���,B1.��ע as ���˱�ע" & _
        " From סԺ���ü�¼ A,������Ϣ B,���ű� C,���ű� C1,������ҳ B1 " & _
        " Where Rownum=1  And A.����id=B1.����id(+) and A.��ҳid=B1.��ҳID(+) And NO=[1] And A.��¼����=[2]" & _
        "       And A.����ID=B.����ID And Instr([3],A.��¼״̬)>0" & _
                IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
        "     And A.��������ID=C.ID and A.ִ�в���ID=C1.ID(+)"
    Else
        strSQL = _
        " Select A.����ID,0 as ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.���ʽ as ����,A.��ʶ��," & _
        "           0 as ���˲���ID,A.��������ID,A.�Ӱ��־,A.Ӥ����,A.������,A.������,A.����Ա����," & _
        "           A.��������ID,A.ִ�в���ID," & IIF(zlIsShowDeptCode, "C.����||'-'||", "") & "C.���� as �������� ," & IIF(zlIsShowDeptCode, "C.����||'-'||", "") & "C.����  as ִ�в���,A.����ʱ��," & _
        "           B.ҽ�Ƹ��ʽ,B.������,B.������,A.�Ƿ���,Null as ���˱�ע" & _
        " From ������ü�¼ A,������Ϣ B,���ű� C ,���ű� C1" & _
        " Where Rownum=1  And NO=[1] And A.��¼����=[2]" & _
        "           And A.����ID=B.����ID And Instr([3],A.��¼״̬)>0" & _
                    IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
        "           And A.��������ID=C.ID and A.ִ�в���ID=C1.ID(+)"
    End If
    If blnNOMoved Then
        strSQL = Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab)
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ��ݡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mlng����ID = 0 Then mlng����ID = Nvl(rsTmp!����ID, 0)
    
    cboNO.Text = strNO
    txt����.Text = Nvl(rsTmp!����)
    txt�Ա�.Text = Nvl(rsTmp!�Ա�)
    txt����.Text = Nvl(rsTmp!����)
    If Nvl(rsTmp!��ҳID, 0) <> 0 Then
        txt����.Text = Nvl(rsTmp!����)
    End If
    
    '���˺� ����:26953 ����:2009-12-25 15:23:48
    txt���˱�ע.Text = Nvl(rsTmp!���˱�ע)
    If mint������Դ = 1 Then
        lblסԺ��.Caption = "�����"
    Else
        lblסԺ��.Caption = "סԺ��"
    End If
    txtסԺ��.Text = Nvl(rsTmp!��ʶ��)
    
    txt�ѱ�.Text = Nvl(rsTmp!�ѱ�)
    txt������.Text = Nvl(rsTmp!������)
    txt������.Text = Format(Nvl(rsTmp!������), "0.00")
    txt���ʽ.Text = Nvl(rsTmp!ҽ�Ƹ��ʽ)
    
    mblnCboNotClick = True
    cbo��������.AddItem Nvl(rsTmp!��������)
    cbo��������.ItemData(cbo��������.NewIndex) = Nvl(rsTmp!��������ID, 0)
    cbo��������.ListIndex = cbo��������.NewIndex
    
    
    mlngִ�пⷿID = Nvl(rsTmp!ִ�в���ID, 0)
    cboִ�в���.AddItem Nvl(rsTmp!ִ�в���)
    cboִ�в���.ItemData(cboִ�в���.NewIndex) = mlngִ�пⷿID
    cboִ�в���.ListIndex = cboִ�в���.NewIndex

    mlng����ⷿID = Set����ⷿID(mlngִ�пⷿID)
    mblnCboNotClick = False
    
    If Nvl(rsTmp!�Ƿ���, 0) = 1 Then
        chk����.value = 1: chk����.Visible = True
    End If
    
    chk�Ӱ�.value = Nvl(rsTmp!�Ӱ��־, 0)
    Call LoadPatientBaby(cboBaby, rsTmp!����ID, rsTmp!��ҳID)
    Call zlControl.CboLocate(cboBaby, Nvl(rsTmp!Ӥ����, 0), True)
    
    '������
    Call GetCboIndex(cbo������, Nvl(rsTmp!������))
    If cbo������.ListIndex = -1 And Not IsNull(rsTmp!������) Then
        cbo������.AddItem rsTmp!������
        cbo������.ListIndex = cbo������.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    If mint��¼���� = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID, IIF(mint������Դ = 1, 0, rsTmp!��ҳID))
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
            "       Nvl(A.����,1)*A.���� as ԭʼ����" & _
            " From " & mstrFeeTab & " A " & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
            "            And A.��¼����=[2]"
        
        '��ȡҩƷ�շ���¼�е�׼����
        strSQL2 = _
            " Select A.����ID,Max(A.����) as ����,Max(A.��Ʒ����) as ��Ʒ���� ,Max(�ڲ�����) as �ڲ�����, " & _
            "       Sum(Nvl(A.����,1)*A.ʵ������)  as ׼������" & _
            " From ҩƷ�շ���¼ A " & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            "       And A.����� is NULL And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        strSQL = "Select Nvl(�۸񸸺�,���) From " & mstrFeeTab & _
            " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1]" & _
            " And Nvl(ִ��״̬,0)<>1" & IIF(mlngҽ��ID <> 0, " And ҽ�����+0=[8]", "")
        
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        If mint��¼���� = 2 Then
            If mint������Դ = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, rsTmp!����ID, intInsure)
            Else
                blnDo = gbytBillOpt = 2
            End If
            If blnDo Then
                strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                    " (" & _
                    " Select Nvl(�۸񸸺�,���) as ���" & _
                    " From " & mstrFeeTab & _
                    " Where NO=[1] And ��¼���� IN(2,12)" & _
                    " Group by Nvl(�۸񸸺�,���)" & _
                    " Having Sum(Nvl(���ʽ��,0))=0" & _
                    " )"
            End If
        End If
        
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            "       C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������, A.���㵥λ," & _
            "       Avg(Nvl(A.����,1)) as ����, Avg(A.����) as ����," & _
            "       Sum(A.��׼����) as ����, Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            "       D.���� as ִ�в���,A.���ӱ�־" & _
            " From " & mstrFeeTab & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            "           And A.��¼����=[2]" & _
            "           And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),C.����,C.����,A.�շ�ϸĿID,B.����," & _
            "           B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־"
            
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = _
            " Select A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������,A.���㵥λ, " & _
            "           max(C.����) as ����,Max(C.��Ʒ����) as ��Ʒ����,Max(C.�ڲ�����) as �ڲ�����," & _
            "           Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            "           Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            "           Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            "           A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.���=B.��� And B.ID=C.����ID(+)" & _
            " Group by A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������," & _
            "           A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���,A.���ӱ�־" & _
            " Having Sum(A.����*A.����)<>0"
            
        strSQL = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���," & _
            "       A.��������,A.���㵥λ,A.����,A.��Ʒ����,A.�ڲ�����,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
            "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
            "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
            "       A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            "       And  A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIF(mblnDelete, -1, 1) '����,�����������
        
        strSQL2 = _
            " Select A.����ID,Max(A.����) as ����,Max(A.��Ʒ����) as ��Ʒ���� ,Max(�ڲ�����) as �ڲ����� " & _
            " From ҩƷ�շ���¼ A " & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1 And Instr([4],A.��¼״̬)>0 " & _
            "       And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
            
        strSQL = _
            "   Select A.�շ�ϸĿID,A.�շ����,A.ִ�в���ID,Nvl(A.�۸񸸺�,A.���) as ���,B.����,B.��Ʒ����,B.�ڲ�����," & _
            "           A.���㵥λ,A.����,A.����,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��,A.���ӱ�־,A.��������" & _
            "   From " & mstrFeeTab & " A,( " & strSQL2 & ") B " & _
            "   Where A.��¼����=[2] And Instr([4],A.��¼״̬)>0 And A.NO=[1]" & _
                        IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "") & _
            "           A.ID=B.����ID(+) "
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace("ҩƷ�շ���¼", Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab), "HҩƷ�շ���¼")
        End If
        
        strSQL = _
            " Select A.���,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������, A.���㵥λ, " & _
            "           max(A.����) as ����,Max(A.��Ʒ����) as ��Ʒ����,Max(A.�ڲ�����) as �ڲ�����," & _
            "           Avg(Nvl(A.����,1)) as ����, Avg([7]*A.����) as ����," & _
            "           Sum(A.��׼����) as ����," & _
            "           Sum([7]*A.Ӧ�ս��) as Ӧ�ս��,Sum([7]*A.ʵ�ս��) as ʵ�ս��, " & _
            "           D.���� as ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ����" & _
            "           And A.ִ�в���ID=D.ID(+) " & _
            " Group by A.���,C.����,C.����,A.�շ�ϸĿID,B.����,B.���," & _
            "           Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־ "
            
        strSQL = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
            "       A.���㵥λ,A.����,A.��Ʒ����,A.�ڲ�����,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            "       And  A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " Order by ���"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, IIF(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), IIF(gbytҩƷ������ʾ = 1, 3, 1), intSign, mlngҽ��ID)
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
        Bill.TextMatrix(i, BillCol.��Ŀ) = rsTmp!����
        Bill.TextMatrix(i, BillCol.��Ʒ��) = Nvl(rsTmp!��Ʒ��)
        Bill.TextMatrix(i, BillCol.���) = Nvl(rsTmp!���)
        Bill.TextMatrix(i, BillCol.��λ) = Nvl(rsTmp!���㵥λ)
        Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = FormatEx(rsTmp!����, 5)
        Bill.TextMatrix(i, BillCol.����) = Format(rsTmp!����, gstrDecPrice)
        Bill.TextMatrix(i, BillCol.Ӧ�ս��) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, BillCol.ʵ�ս��) = Format(rsTmp!ʵ�ս��, gstrDec)
       Bill.TextMatrix(i, BillCol.�ڲ�����) = Nvl(rsTmp!�ڲ�����)
       Bill.TextMatrix(i, BillCol.��Ʒ����) = Nvl(rsTmp!��Ʒ����)
       Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!��������)
        
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "��"
        End If
        
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))

        '��ȡҩƷ�շ���¼�е�׼����
        strSQL1 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������) as ׼������" & _
            " From ҩƷ�շ���¼ A " & _
            " Where    A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            "           And A.����� is NULL And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "" & _
            "   Select Nvl(�۸񸸺�,���) From " & mstrFeeTab & _
            "   Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1]" & _
            "       And Nvl(ִ��״̬,0)<>1" & IIF(mlngҽ��ID <> 0, " And ҽ�����+0=[7]", "")
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From " & mstrFeeTab & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
        
        strSQL = _
            "   Select Sum(A.ID) as ID,A.���,A.����,A.�շ����," & _
            "       Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
            "       Sum(A.ʵ�ս��) as ʣ��ʵ�� " & _
            "   From (  Select Decode(A.��¼״̬,2,0,A.ID) as ID,A.���,B.����,A.�շ����," & _
            "                       Nvl(A.����,1)*A.����  as ����, A.Ӧ�ս��,A.ʵ�ս��" & _
            "               From " & mstrFeeTab & " A,������Ŀ B " & _
            "               Where A.��¼����=[2] And A.NO=[1]" & _
            "                           And A.������ĿID=B.ID And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            "              ) A" & _
            " Group by A.���,A.����,A.�շ����" & _
            " Having Sum(A.����)<>0"
                    
        '��������
        strSQL = _
            "   Select A.����,Sum(A.ʣ��Ӧ��*(A.׼������/A.ʣ������)) as Ӧ�ս��," & _
            "       Sum(ʣ��ʵ��*(A.׼������/A.ʣ������)) as ʵ�ս��  " & _
            "   From ( Select A.����,A.ʣ������,A.ʣ��Ӧ��,A.ʣ��ʵ��," & _
            "                   Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������)) as ׼������" & _
            "               From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
            "               Where A.ID=B.����ID(+)" & _
            "              ) A  " & _
            "   Group by A.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIF(mblnDelete, -1, 1) '����,�����������
        
        strSQL = "Select A.������ĿID,A.Ӧ�ս��,A.ʵ�ս�� From " & mstrFeeTab & " A" & _
            " Where Instr([4],A.��¼״̬)>0 And A.��¼����=[2] And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        strSQL = _
            " Select B.����,Sum([6]*A.Ӧ�ս��) as Ӧ�ս��,Sum([6]*A.ʵ�ս��) as ʵ�ս�� " & _
            " From (" & strSQL & ") A,������Ŀ B Where A.������ĿID=B.ID Group By B.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint��¼����, IIF(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), intSign, mlngҽ��ID)
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetShowCol()
    '���ܣ������еĿ���(���ʱչ��)
    Bill.ColWidth(BillCol.����) = 0
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub
 
Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSQL As String, bytDay As Byte
    Dim strҩ�� As String, lng��������ID As Long
    
    lng��������ID = mrsInfo!����ID    '������������
    If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)

    strSQL = _
    " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
    "       And B.������� IN([1],3) And B.����ID=C.ID" & _
    "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
    "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
    "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
    "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
    "       And A.�շ�ϸĿID=[3]" & _
    " Order by B.�������,C.����"
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint������Դ, lng��������ID, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load������(ByVal lng����id As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngOldID As Long
    
    cbo������.Clear
    
    '����ҽ����ʿ
    strSQL = _
    "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
    "           C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
    " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
    " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
    "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
    "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
    "       And C.��Ա���� IN('ҽ��','��ʿ') And B.����ID=[1]  " & _
    "  Order by ����,��Ա���� Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id)
    
    i = IIF(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo������.AddItem IIF(IsNull(rsTmp!����), "", rsTmp!���� & "-") & rsTmp!����
                cbo������.ItemData(cbo������.ListCount - 1) = rsTmp!����ID
                marrDr(cbo������.ListCount - 1) = rsTmp!ID & "|" & rsTmp!����ID & "|" & Nvl(rsTmp!���) & "|" & rsTmp!���� & "|" & Nvl(rsTmp!����) & "|" & rsTmp!ְ�� & "|" & Nvl(rsTmp!��Ա����)
                
                If rsTmp!���� = mstr����ҽ�� Then cbo������.ListIndex = cbo������.NewIndex
                If rsTmp!ID = UserInfo.ID And cbo������.ListIndex = -1 Then cbo������.ListIndex = cbo������.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo������.ListCount > 0 Then ReDim Preserve marrDr(cbo������.ListCount - 1)
        
        If cbo������.ListCount = 1 And cbo������.ListIndex = -1 Then cbo������.ListIndex = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
            
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ) - 100
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����) - 50
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��) - 50
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ)
           Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����)
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��)
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��)
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
        Bill.TextMatrix(i, BillCol.��) = i
    Next
    Bill.Redraw = True
End Sub
 

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
'���ܣ��ж�ָ�������ڵ������Ƿ��Ѿ�����
'������objDetail=��Ŀ,intRow=Ҫ�жϵ���
'˵����ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ)
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                        vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        PhysicExist = True
                    End If
                    Exit Function
                Else
                    If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        PhysicExist = True
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
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    Check�������� = True
    
    '�޷����
    If txt���ʽ.Tag = "" Then Exit Function
    
    '45605
    'ֻ���ҽ�����˺͹��Ѳ���
    If zlIsCheckMedicinePayMode(txt���ʽ.Text, blnҽ��, bln����) = False Then Exit Function
    'ȷ����������
    bytType = IIF(blnҽ��, 1, 2) ' Val(txt���ʽ.Tag)
    
    '��ȡ�������
    If bytType = 1 Then
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
    Else
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
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
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!����) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                dblAllTime = mobjBill.Details(i).���� * mobjBill.Details(i).����
                strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����, _
                    mobjBill.Details(i).ժҪ & "||" & dblAllTime)
                    
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


Public Sub InitLocPar()
'���ܣ���ʼ�����ñ�������
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mblnTime = Val(zlDatabase.GetPara("�����������", glngSys, pҽ�����ѹ���)) <> 0
    mbytSendMateria = Val(zlDatabase.GetPara("���ʺ�ҩ", glngSys, pҽ�����ѹ���))
    'mlng���ϲ��� = Val(zldatabase.GetPara(IIF(mint������Դ = 2, "סԺ", "����") & "ȱʡ���ϲ���", glngSys, pҽ�����ѹ���))
End Sub

Public Function zlCheck����ҽ��(ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ա���ҽ����һЩ���
    '���:intInsuer-����
    '����:
    '����:���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-07 10:25:04
    '����:27278
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    If intInsure = 0 Then zlCheck����ҽ�� = True: Exit Function
    
    err = 0: On Error GoTo Errhand:
    '���˺�:???
    'mint������Դ:1-���ﲡ��,2-סԺ����
    'mint��¼����:1-�շ�(����),2-����(��/ס)
    'mbytInState :0-ִ��,1-����,2-����(��֧��),3-ɾ��
    
    'ֻ�л��۲�֧�ּ��
    If (mint������Դ = 2 Or mint��¼���� = 2) And mbytInState <> 0 Or MCPAR.ҽ��ȷ���������� = False Then
        zlCheck����ҽ�� = True: Exit Function
    End If
    
    'showmsgbox
    '������strCaption=��Ϣ�������
    '      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
    '      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
    '              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
    '              ÿ����ť�������֧��4������
    '      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
    '���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    strTemp = zlCommFun.ShowMsgBox("��������", "��ȷ����ǰҽ�����˱���Ҫ���͵�ҩƷ���������͡�", "!ҽ����(&A),ҽ����(&B),?ȡ��(&C)", Me)
    If strTemp = "" Then Exit Function
    '����ǲ������շѻ��۵�������ҽ�����ˣ���ҽ��������supportҽ��ȷ���������͡���Чʱ������ʱ��ʾ�õ����ǡ�ҽ���ڣ�ҽ���⡱�������ҽ���ڷ��ü�¼ժҪ�д��1��ҽ������2��
    strTemp = IIF(strTemp = "ҽ����", 1, 2)
    For Each mobjBillDetail In mobjBill.Details
        mobjBillDetail.ժҪ = strTemp
    Next
    zlCheck����ҽ�� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlCheck������۸����(ByVal lng�շ�ϸĿID As Long, bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ������(����Ϊ���)
    '���:
    '����:
    '����:���������ĿΪ��,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-12 11:22:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
   '���˺� ����:27286 ���۵ļ۸�Ϊ��Ĳ����м����� ����:2010-01-07 15:13:45
   Dim strSQL As String, rs�۸� As ADODB.Recordset, dbl�۸� As Double
    err = 0: On Error GoTo Errhand:
   zlCheck������۸���� = False
    If bln���� Then
        strSQL = _
        " Select  B.�ּ� " & _
        " From �շѼ�Ŀ B " & _
        " Where   ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        "       And B.�շ�ϸĿID=[1]"
        Set rs�۸� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ�۸�", lng�շ�ϸĿID)
        If rs�۸�.EOF = False Then
            dbl�۸� = Val(Nvl(rs�۸�!�ּ�))
        Else
            dbl�۸� = 0
        End If
        If dbl�۸� = 0 Then zlCheck������۸���� = True: Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
 
Public Function Get����ҩ�嵥(strNO As String, strTime As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݷ��õ��ݺ�,�Ǽ�ʱ��,��ȡ����ҩƷ�嵥
    '��Σ�strNO-���ݺ�
    '          strTime-�Ǽ�ʱ��
    '���Σ�
    '���أ�����ҩ�嵥
    '���ƣ����˺�
    '���ڣ�2010-03-19 18:59:27
    '˵������ͨ��ҩʱΪ���˿��ң����ҽ����Ϊ�������ҡ�
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.�ⷿID,A.�Է�����ID" & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B" & _
        " Where A.NO=[1] And A.����=[2] And Mod(A.��¼״̬,3)=1 And A.����� is NULL" & _
        " And A.NO=B.NO And A.����ID=B.ID And B.��¼״̬<>0 And B.�Ǽ�ʱ��+0=[3]" & _
        " Order by A.ҩƷID"
    If strTime <> "" Then
        Set Get����ҩ�嵥 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 9, CDate(strTime))
    Else
        Set Get����ҩ�嵥 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 9)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub SetDetailtStock(ByVal lngִ�п���ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ�Ŀ������
    '���ƣ����˺�
    '���ڣ�2010-07-12 14:27:51
    '˵����
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim strҩ��IDs As String, dblStock As Double
    '��ȡ���
    '�������ҩƷ������
    '����
    dblStock = GetStock(objDetail.ID, lngִ�п���ID)
    objDetail.��� = dblStock
End Sub

Private Function GetInputDetail(ByVal lng��Ŀid As Long) As Detail
    '���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    
    If lngMediCareNO > 0 Then
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,M.Ҫ������," & _
        "       D.����ID as ҩ��ID, D.���÷���  as ����, 1 as ҩ����װ, A.���㵥λ  as ҩ����λ,D.��������,A.¼������" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,����֧����Ŀ M" & _
        " Where   A.ID=D.����ID(+) And B.����=A.���" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And A.ID=[1] And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[3]"
    Else
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,0 as Ҫ������, D.����ID as ҩ��ID," & _
        "       D.���÷��� as ����, 1 as ҩ����װ, A.���㵥λ as ҩ����λ,D.��������,A.¼������" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1" & _
        " Where  A.ID=D.����ID(+) And B.����=A.���" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, IIF(gbytҩƷ������ʾ = 1, 3, 1), lngMediCareNO)
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
        .��ҩ��̬ = 0
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
        
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SelectItem(ByVal blnInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ��������������Ŀ
    '���:strKey-����ѡ��
    '       blnInput-����
    '����:
    '����:
    '����:���˺�
    '����:2010-12-14 14:32:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTXTHwnd As Long, strInput As String
    Dim str��׼��Ŀ As String, int���� As Integer, lng��Ŀid As Long
    Dim rsItem As ADODB.Recordset
    
    On Error GoTo errHandle
    If blnInput Then
        lngTXTHwnd = Bill.TxtHwnd
        strInput = Bill.Text
    End If
    If Not IsNull(mrsInfo!����) Then
        int���� = mrsInfo!����
        '���˺�:24862
        'mint������Դ As Integer '1-���ﲡ��,2-סԺ����
        'mint��¼���� As Integer '1-�շ�(����),2-����(��/ס)
        If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), (mint��¼���� = 1 Or mint������Դ = 1)) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
    End If
    
    If frmStuffSelect.ShowSelect(Me, mstrPrivs, mint������Դ, int����, strInput, lngTXTHwnd, str��׼��Ŀ, mlng����ⷿID, False, rsItem) = False Then GoTo GoNotSel
    If rsItem Is Nothing Then GoTo GoNotSel:
    If rsItem.State <> 1 Then GoTo GoNotSel:
    If rsItem.RecordCount = 0 Then GoTo GoNotSel:
    lng��Ŀid = Val(Nvl(rsItem!�շ���ĿID))
    Set mobjDetail = GetInputDetail(lng��Ŀid)
    If int���� <> 0 Then sta.Panels(4).Text = Getҽ������(lng��Ŀid, int����)
    mobjDetail.���� = Val(Nvl(rsItem!����))
    mobjDetail.��Ʒ���� = Trim(Nvl(rsItem!��Ʒ����))
    mobjDetail.�ڲ����� = Trim(Nvl(rsItem!�ڲ�����))
    mobjDetail.��� = Val(Nvl(rsItem!���ÿ��))
    SelectItem = True
GoNotSel:
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Set����ⷿID(ByVal lngִ�п��� As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ִ�п���,ȷ������ⷿID
    '����:����ⷿID
    '����:���˺�
    '����:2010-12-15 10:06:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select ����ⷿid  From ����ⷿ���� Where ����id = [1] And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngִ�п���)
    If Not rsTemp.EOF Then
        Set����ⷿID = Val(Nvl(rsTemp!����ⷿid))
    Else
        Set����ⷿID = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
