VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\ZLIDKind\ZLIDKIND.vbp"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmSchSchedule 
   Caption         =   "�����ĿԤԼ"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13545
   Icon            =   "frmSchSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   13545
   StartUpPosition =   1  '����������
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox pictDay 
      BackColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picTimeTable 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   4920
      ScaleHeight     =   7095
      ScaleWidth      =   5655
      TabIndex        =   3
      Top             =   240
      Width           =   5655
      Begin VB.Frame frmTimeTable 
         Caption         =   "ԤԼʱ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8415
         Begin zl9PACSWork.ucScheduleTimetable schTimeTable 
            Height          =   6615
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   11668
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   960
      ScaleHeight     =   9135
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Frame frmInfo 
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3615
         Begin VB.ComboBox cboSchDevice 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3045
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtNotice 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1635
            Width           =   2055
         End
         Begin VB.TextBox txtPhone 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label lblSchDevice 
            Caption         =   "CT"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1400
            TabIndex        =   21
            Top             =   3090
            Width           =   1335
         End
         Begin VB.Label lblOrderInfo 
            Caption         =   "ҽ�����ݣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label lblAddress 
            Caption         =   "���ע�⣺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1665
            Width           =   1095
         End
         Begin VB.Label lblPhone 
            Caption         =   "�绰��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label lblInfo 
            Caption         =   "ԤԼʱ�䣺10:20 - 11:30"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   3960
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "ԤԼ���ڣ�2018-2-3"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   3525
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmSchSchedule.frx":0442
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   2415
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "ԤԼ�豸��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   3090
            Width           =   1335
         End
         Begin VB.Label lblInfo 
            Caption         =   "���䣺25   ��Դ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   795
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "����������    �Ա���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3135
         End
      End
      Begin XtremeCalendarControl.DatePicker dpCalendar 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   4320
         Width           =   3615
         _Version        =   1048579
         _ExtentX        =   6376
         _ExtentY        =   5106
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowNoneButton  =   0   'False
         ShowNonMonthDays=   0   'False
         Show3DBorder    =   0
         AskDayMetrics   =   -1  'True
         TextTodayButton =   "ѡ�����"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSchOther 
         Height          =   1815
         Left            =   0
         TabIndex        =   8
         Top             =   7320
         Width           =   3615
         _cx             =   6376
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   8940
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmSchSchedule.frx":0454
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17004
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":288C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":365E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":4430
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":5202
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":5FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":6DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":7B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":894A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":971C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":A4EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":B2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":C092
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":CE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":DC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":EA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":F7DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":105AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1137E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":12150
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":12F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":13CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":14AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":15898
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1666A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1743C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1820E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":18FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":19DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1AB84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   1440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSchSchedule.frx":1B956
      Left            =   360
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOrderID As Long                 'ҽ��ID
Private mschDate As Date                    '��ǰ��ԤԼ����
Private mlngSchDeviceID As Long             '��ǰѡ�е�ԤԼ�豸ID
Private mblnISScheduled As Boolean          '�Ƿ��Ѿ�ԤԼ
Private mblnNewSchedule As Boolean          '�Ƿ��½�ԤԼ
Private mstrDefaultPatientType As String    'ȱʡ��������
Private mfrmParent As Object                '������
Private mstrDeptIDs As String               '����ID��
Private mlngDeptID As Long                  '��ǰ����ID
Private mstrModifiedOrderID As String       '�����ԤԼ��Ϣ��ҽ��ID�����á�,������
Private mblnExecFee As Boolean              '�Ƿ�ԤԼʱִ�з���
Private mblnAutoPrint As Boolean            '�Ƿ��Զ���ӡԤԼ��
Private mstrSchRestDate As String           '������Ϣ��
Private mblnCheckIn As Boolean              '�Ƿ񱣴�ԤԼ�󱨵�
Private mblnLoadingDevice As Boolean        '�Ƿ����ڼ���ԤԼ�豸
Private mlngPatSource As Long               '������Դ
Private mstrPrivs As String                 '�����ߵ�Ȩ��
Private mblnIsForceModify As Boolean        '�Ƿ�ǿ���޸�סԺ������Ϣ
Private mblnLoadDone As Boolean             '�Ƿ���ɴ���ļ���

'���ԤԼ�豸
Private Enum constScheduleDeviceList
    col_SchDevice_ID = 0
    col_SchDevice_ѡ�� = 1
    col_SchDevice_Ӱ����� = 2
    col_SchDevice_�豸���� = 3
    col_SchDevice_�豸˵�� = 4
End Enum

'�����豸�ϵ�ԤԼ��Ϣ
Private Enum constSchOtherList
    col_SchOther_ID = 0
    col_SchOther_ԤԼ�豸���� = 1
    col_SchOther_ԤԼ���� = 2
    col_SchOther_ҽ������ = 3
    col_SchOther_ԤԼ��ʼʱ�� = 4
    col_SchOther_ԤԼ����ʱ�� = 5
End Enum

Private Sub InitCommandBar()
'------------------------------------------------
'���ܣ���ʼ��������
'������ ��
'���أ� ��
'------------------------------------------------
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '�ⲿ��ȫ�����ã��Ƿ��Ҫ��
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
        
    With cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '����ʾ�˵�
    cbrMain.ActiveMenuBar.Visible = False
    
    '��ʾ������
    Set cbrToolBar = cbrMain.Add("ԤԼ������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Save, "����ԤԼ")
        cbrControl.iconid = 6823
        cbrControl.ToolTipText = "����ԤԼ��Ϣ"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Print, "��ӡԤԼ��")
        cbrControl.iconid = 103
        cbrControl.ToolTipText = "��ӡ���ߵ�ԤԼ֪ͨ��"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_New, "�½�ԤԼ")
        cbrControl.iconid = 6886
        cbrControl.ToolTipText = "�½�һ�����ԤԼ"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Delete, "ɾ��ԤԼ")
        cbrControl.iconid = 6822
        cbrControl.ToolTipText = "ɾ��һ�����ԤԼ"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Refresh, "ˢ��")
        cbrControl.iconid = 791
        cbrControl.ToolTipText = "ˢ������"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_SaveAndCheckin, "���汨��")
        cbrControl.iconid = 744
        cbrControl.ToolTipText = "����ԤԼ���رմ��ڣ���鱨��"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_SaveAndQuit, "�����˳�")
        cbrControl.iconid = 3013
        cbrControl.ToolTipText = "����ԤԼ���رմ���"
         
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Quit, "�˳�")
        cbrControl.iconid = 191
        cbrControl.ToolTipText = "�رմ���"
        
    End With
    
    cbrToolBar.Position = xtpBarTop
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboSchDevice_Click()
    If cboSchDevice.ListIndex >= 0 And mblnLoadingDevice = False Then
        '�޸ĵ�ǰ��ѡ�е�ԤԼ�豸ID
        mlngSchDeviceID = cboSchDevice.ItemData(cboSchDevice.ListIndex)
        Call RefreshSchedule(False, True)
        Call RefreshCalendar
    End If
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_New        '�½�ԤԼ
            Call NewSchedule
            
        Case conMenu_PacsSchdule_Delete     'ɾ��ԤԼ
            Call DelSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Print      '��ӡԤԼ��
            Call PrintSchedule
            
        Case conMenu_PacsSchdule_Refresh    'ˢ��
            Call RefreshForm
            
        Case conMenu_PacsSchdule_ModifyInfo '�޸���Ϣ
            Call ModifyPatInfo
        
        Case conMenu_PacsSchdule_Save       '����ԤԼ
            Call SaveSchedule
            Call loadTimeTable
            
        Case conMenu_PacsSchdule_SaveAndCheckin '�����˳��Ҵ򿪱�������
            If SaveSchedule = True Then
                mblnCheckIn = True
                Unload Me
            End If

        Case conMenu_PacsSchdule_SaveAndQuit '�����˳�
            If SaveSchedule = True Then
                Unload Me
            End If
        
        Case conMenu_PacsSchdule_Quit       '�˳�
            Unload Me
            
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Delete     'ɾ��ԤԼ
            Control.Enabled = mblnISScheduled
            
        Case conMenu_PacsSchdule_Print      '��ӡԤԼ��
            Control.Enabled = mblnISScheduled
        
        Case conMenu_PacsSchdule_Save, conMenu_PacsSchdule_SaveAndCheckin, _
             conMenu_PacsSchdule_SaveAndQuit     '����ԤԼ,�����˳��Ҵ򿪱�������,�����˳�
            Control.Enabled = mblnNewSchedule
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picInfo.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picTimeTable.hwnd
    End If
End Sub

Private Sub dpCalendar_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If InStr(mstrSchRestDate, Format(Day, "YYYY-MM-DD")) > 0 Then
        Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
    End If
End Sub

Private Sub dpCalendar_MonthChanged()
    If mblnISScheduled = True Or (Format(dpCalendar.FirstVisibleDay, "YYYY-MM") < Format(Now, "YYYY-MM")) Then
        ChangeCalendar (mschDate)
    Else
        Call RefreshCalendar
    End If
End Sub

Private Sub dpCalendar_SelectionChanged()
    Dim dtDate As Date
    
    '���������ڣ�����ˢ��ʱ���
    dtDate = dpCalendar.Selection.Blocks(0).DateBegin
    If InStr(mstrSchRestDate, Format(dtDate, "YYYY-MM-DD")) > 0 Or mblnISScheduled = True Then
        If dtDate = Format(Now, "YYYY-MM-DD") And mblnISScheduled = False Then
            Call MsgBox("����ԤԼ�Ѿ����ˡ�", vbInformation, "���ԤԼ��ʾ")
        End If
        '���޷�ԤԼ�����ӣ���ѡ��
        ChangeCalendar (mschDate)
    Else
        mschDate = dtDate
    End If
    
    Call RefreshSchedule(False, True)
End Sub

Private Sub InitFaceScheme()
'------------------------------------------------
'���ܣ���ʼ�����沼��
'������ ��
'���أ� ��
'------------------------------------------------
    Dim Pane1 As Pane, Pane2 As Pane
    
    On Error GoTo err
    
    '����������ʾ����
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .options.HideClient = True
        .options.UseSplitterTracker = False 'ʵʱ�϶�
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
        dkpMain.options.DefaultPaneOptions = PaneNoCaption
    End With
    
    '�ȴ�ע����ȡԤ�����úõĴ��ڲ��֣�Ȼ�����������
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    '���ע����б���Ľ��沼��Pane�������ԣ������Ĭ�ϵ�Pane����
    If dkpMain.PanesCount <> 2 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, 350, 150, DockLeftOf)
        Pane1.title = "ԤԼ��Ϣ"
        Pane1.options = PaneNoCaption
        
        Set Pane2 = dkpMain.CreatePane(2, 650, 300, DockRightOf, Pane1)
        Pane2.title = "ԤԼʱ���"
        Pane2.options = PaneNoCaption
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    If mblnLoadDone = True Then
        If mblnNewSchedule = True Then
            If MsgBox("�Ƿ񱣴���ԤԼ��Ϣ��", vbYesNo, "���ԤԼ��ʾ") = vbYes Then
                Call SaveSchedule
            End If
        End If
        
        '�رմ����ʱ�򣬱�����沼��
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    End If
    
    Set mfrmParent = Nothing
    ' '�رմ���ʱ�����ͷţ��ᵼ��VB����
'    '�ͷ�DockingPane
'    For i = 1 To dkpMain.PanesCount
'        dkpMain.Panes(i).Handle = 0
'    Next i
'    dkpMain.CloseAll
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    frmInfo.Left = 30
    frmInfo.Top = 50
    frmInfo.Width = picInfo.ScaleWidth - 30
    
    lblInfo(0).Width = frmInfo.Width - lblInfo(0).Left - 50
    lblInfo(1).Width = lblInfo(0).Width
    lblInfo(2).Width = lblInfo(0).Width
    lblInfo(4).Width = lblInfo(0).Width
    lblInfo(5).Width = lblInfo(0).Width
    txtPhone.Width = frmInfo.Width - 1450
    txtNotice.Width = txtPhone.Width
    lblSchDevice.Width = txtPhone.Width
    cboSchDevice.Width = txtPhone.Width
    
    dpCalendar.Left = 0
    dpCalendar.Top = frmInfo.Height + 10
    dpCalendar.Width = frmInfo.Width
    
    vsfSchOther.Left = 0
    vsfSchOther.Top = dpCalendar.Top + dpCalendar.Height + 30
    vsfSchOther.Width = frmInfo.Width
    vsfSchOther.Height = picInfo.ScaleHeight - vsfSchOther.Top - 300
End Sub

Private Sub picTimeTable_Resize()
    On Error Resume Next
    
    frmTimeTable.Left = 0
    frmTimeTable.Top = 0
    frmTimeTable.Width = picTimeTable.ScaleWidth
    frmTimeTable.Height = picTimeTable.ScaleHeight - stbThis.Height
    
    schTimeTable.Left = 0
    schTimeTable.Top = 0
    schTimeTable.Width = frmTimeTable.Width
    schTimeTable.Height = frmTimeTable.Height
End Sub

Public Function ZlShowMe(ByVal strPrivs As String, ByVal lngOrderID As Long, ByVal strDeptIDs As String, _
    ByVal frmParent As Object, Optional ByRef blnCheckin As Boolean = False) As String
'------------------------------------------------
'���ܣ��򿪴���
'������ lngOrderID -- ҽ��ID
'       strDeptIDs -- ����ID��
'       frmParent -- ������
'       strPrivs -- �����ߵ�Ȩ��
'���أ�����ԤԼ���ҽ��ID
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mblnLoadDone = False
    mlngOrderID = lngOrderID
    Set mfrmParent = frmParent
    mstrDeptIDs = strDeptIDs
    mstrModifiedOrderID = ""
    mlngSchDeviceID = 0
    mblnCheckIn = False
    mstrPrivs = strPrivs
    mblnIsForceModify = CheckPopedom(mstrPrivs, "ǿ���޸�סԺ������Ϣ")
    
    '�����ȫ�����ң��Ȳ�ִ�п���ID�����û����ȡ��һ������
    If InStr(mstrDeptIDs, ",") > 0 Then
        strSQL = "select ִ�п���ID from ����ҽ����¼ where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ����ID", mlngOrderID)
        If rsTemp.EOF = False Then
            mlngDeptID = NVL(rsTemp!ִ�п���ID)
        Else
            mlngDeptID = Split(mstrDeptIDs, ",")(0)
        End If
    Else
        mlngDeptID = Val(mstrDeptIDs)
    End If
    
    '��ȡ����
    mblnExecFee = IIf(Val(zlDatabase.GetPara("ԤԼʱִ�з���", glngSys, 1292)) = 1, True, False)
    mblnAutoPrint = IIf(Val(zlDatabase.GetPara("����ԤԼ���Զ���ӡԤԼ��", glngSys, 1292)) = 1, True, False)
    
    '��ʼ�����沼��
    Call InitFaceScheme
    
    '����������
    Call InitCommandBar
    
    '�ȳ�ʼ��ʱ���ؼ�
    Call schTimeTable.Init(1)   '�����ĿԤԼ
    
    Call RestoreWinState(Me, App.ProductName)
    
    '������������
    dpCalendar.AskDayMetrics = True
    dpCalendar.ShowNonMonthDays = False
    mschDate = Format(Now, "YYYY-MM-DD")
    
    '��������
    If LoadData = False Then
        Unload Me
        Exit Function
    End If
    
    Call RefreshCalendar
    
    mblnLoadDone = True
    
    Me.Show 1, mfrmParent
    
    blnCheckin = mblnCheckIn
    ZlShowMe = mstrModifiedOrderID
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadSchDevice() As Boolean
'------------------------------------------------
'���ܣ�����ԤԼ�豸
'������
'���أ�True -- �ɹ���False -- ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim iSelRow As Integer
    
    On Error GoTo err
    
    iSelRow = -1
    mblnLoadingDevice = True
    
    strSQL = "Select  distinct a.id, a.�豸����, a.Ӱ�����, a.�豸˵��, a.�Ƿ�Ĭ��" _
            & " From Ӱ��ԤԼ�豸 A, ����ҽ����¼ B , Ӱ��ԤԼ��Ŀ c " _
            & " Where a.id = c.ԤԼ�豸id And b.������Ŀid = c.������Ŀid And a.�Ƿ����� = 1 " _
            & " And b.ID = [1]  And  a.����id In (" & mstrDeptIDs & ") order by �Ƿ�Ĭ�� desc"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ�豸", mlngOrderID)
    
    '�����ݿ��������
    If rsTemp.RecordCount = 1 Then
        lblSchDevice.Visible = True
        cboSchDevice.Visible = False
        lblSchDevice.Caption = rsTemp!�豸����
        mlngSchDeviceID = rsTemp!ID
    Else
        lblSchDevice.Visible = False
        cboSchDevice.Visible = True
        
        For i = 1 To rsTemp.RecordCount
            cboSchDevice.AddItem (rsTemp!�豸����)
            cboSchDevice.ItemData(cboSchDevice.NewIndex) = rsTemp!ID
            If rsTemp!ID = mlngSchDeviceID Then
                iSelRow = cboSchDevice.NewIndex
            End If
            rsTemp.MoveNext
        Next i
        If iSelRow <> -1 Then
            cboSchDevice.ListIndex = iSelRow
        ElseIf cboSchDevice.ListCount > 1 Then
            cboSchDevice.ListIndex = 0
            mlngSchDeviceID = cboSchDevice.ItemData(0)
        Else
            mlngSchDeviceID = 0
            Call MsgBoxD(Me, "û�п�����ԤԼ��Ӱ���豸���������ԤԼ�豸��", vbOKOnly, "���ԤԼ��ʾ")
            mblnLoadingDevice = False
            Exit Function
        End If
    End If
    mblnLoadingDevice = False
    
    LoadSchDevice = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnLoadingDevice = False
End Function

Private Sub LoadSchOther()
'------------------------------------------------
'���ܣ����ػ����������豸�ϵ�ԤԼ��Ϣ
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    strSQL = "Select a.id, a.ҽ��id, a.ԤԼ�豸����, b.ҽ������, a.ԤԼ��ʼʱ��, " _
        & " a.ԤԼ����ʱ�� From Ӱ��ԤԼ��¼ A, ����ҽ����¼ B, ����ҽ������ C " _
        & " Where a.ҽ��ID = b.ID And b.ID = c.ҽ��ID And c.ִ��״̬ = 0 And b.����id = " _
        & " (Select f.����id From ����ҽ����¼ F Where f.ID = [1]) And a.ҽ��id <> [1] order by ԤԼ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�������豸�ϵ�ԤԼ", mlngOrderID)
    
    With vsfSchOther
        .Rows = rsTemp.RecordCount + 2
        .Cols = 6
        .FixedRows = 2
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
        .ScrollBars = flexScrollBarBoth
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        .ExtendLastCol = True
        
        .ColWidth(col_SchOther_ID) = 50
        .ColWidth(col_SchOther_ҽ������) = 2000
        .ColWidth(col_SchOther_ԤԼ����) = 1200
        .ColWidth(col_SchOther_ԤԼ�豸����) = 1000
        .ColWidth(col_SchOther_ԤԼ��ʼʱ��) = 1000
        .ColWidth(col_SchOther_ԤԼ����ʱ��) = 1000
        
        '�ϲ���һ��
        .RowHeight(0) = 350
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        For i = 0 To 5
            .TextMatrix(0, i) = "�������ԤԼ��Ϣ"
        Next i
        .Cell(flexcpFontBold, 0, 0, 0, 5) = True
        
        '�ڶ�����ʾ����
        .TextMatrix(1, col_SchOther_ҽ������) = "ҽ������"
        .TextMatrix(1, col_SchOther_ԤԼ����) = "ԤԼ����"
        .TextMatrix(1, col_SchOther_ԤԼ�豸����) = "ԤԼ�豸"
        .TextMatrix(1, col_SchOther_ԤԼ��ʼʱ��) = "��ʼʱ��"
        .TextMatrix(1, col_SchOther_ԤԼ����ʱ��) = "����ʱ��"
        .RowHeight(1) = 300
        
        '�����ݿ��������
        i = 1
        While rsTemp.EOF = False
            If mlngOrderID <> rsTemp!ҽ��ID Then
                .TextMatrix(i + 1, col_SchOther_ID) = rsTemp!ID
                .TextMatrix(i + 1, col_SchOther_ҽ������) = rsTemp!ҽ������
                .TextMatrix(i + 1, col_SchOther_ԤԼ����) = Format(rsTemp!ԤԼ��ʼʱ��, "yyyy-mm-dd")
                .TextMatrix(i + 1, col_SchOther_ԤԼ�豸����) = rsTemp!ԤԼ�豸����
                .TextMatrix(i + 1, col_SchOther_ԤԼ��ʼʱ��) = Format(rsTemp!ԤԼ��ʼʱ��, "HH:MM")
                .TextMatrix(i + 1, col_SchOther_ԤԼ����ʱ��) = Format(rsTemp!ԤԼ����ʱ��, "HH:MM")
                i = i + 1
            End If
            rsTemp.MoveNext
        Wend
        
        '���غ�̨������
        .ColHidden(col_SchOther_ID) = True
        
    End With

    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function loadTimeTable() As Boolean
'------------------------------------------------
'���ܣ�����ˢ��ʱ������ݣ�����Ѿ���ԤԼ��Ϣ������ԤԼ��Ϣ���������ó�������ԤԼ�豸������
'������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim strSQL  As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSchDeviceID As Long
    Dim dtSchDate As Date
    
    On Error GoTo err
    
    If mblnISScheduled = True Then
        '�Ѿ�����ԤԼ��Ϣ��ֱ����ʾ����
        If schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID) = False Then
            Exit Function
        End If
        mblnNewSchedule = False
    Else
        mblnNewSchedule = True
        dtSchDate = mschDate
        If schTimeTable.NewSchedule(mlngSchDeviceID, mschDate, mlngOrderID, True) = False Then
            Exit Function
        End If
        If dtSchDate <> mschDate Then
            '������ڱ�����ˣ����޸��������������
            Call ChangeCalendar(mschDate)
        End If
    End If
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
    loadTimeTable = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData() As Boolean
'------------------------------------------------
'���ܣ����ش������������
'������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '���Ⱥ�˳��
    If LoadSchDevice = False Then
        Exit Function
    End If
    
    mblnISScheduled = False
    
    'ˢ�»��߻�����Ϣ
    Call RefreshSchInfo(True)
    
    '����ԤԼ����
    Call ChangeCalendar(mschDate)
    
    '��ȡȱʡ��������
    strSQL = "select ���� from �������� where ȱʡ��־=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡȱʡ��������")
    If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = NVL(rsTemp!����)
    
    Call LoadSchOther
    
    If loadTimeTable = False Then
        Exit Function
    End If
    
    LoadData = True

    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub schTimeTable_OnMenuSchedulePrint()
    Call PrintSchedule
End Sub

Private Sub schTimeTable_OnSchLabelModifed(ByVal iIndex As Integer)
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
    mblnNewSchedule = True
End Sub

Private Sub txtNotice_Change()
    If txtNotice.Locked = False Then
        txtNotice.ForeColor = vbRed
        mblnNewSchedule = True
    End If
End Sub

Private Sub txtNotice_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtNotice_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtNotice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        txtNotice.Locked = False
    End If
End Sub

Private Sub txtPhone_Change()
    If txtPhone.Locked = False Then
        txtPhone.ForeColor = vbRed
        mblnNewSchedule = True
    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPhone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And (mlngPatSource = 3 Or mblnIsForceModify = True) Then    'ֻ��������ֻ��Ų����޸�
        txtPhone.Locked = False
    End If
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'���ܣ�������ݺϷ���
'������
'���أ�True -- �Ϸ���False -- ���ݲ��Ϸ�����Ҫ�޸�
'------------------------------------------------
    On Error GoTo err
    
    '�ֻ��źϷ��Լ��
    If Trim(txtPhone.Text) <> "" Then
        If Not IDKind.IsMobileNo(Trim(txtPhone.Text)) Then
            MsgBox "[�ֻ���]��Ч,������¼�����ɾ����¼������!", vbInformation, gstrSysName
            If txtPhone.Enabled And txtPhone.Visible Then txtPhone.SetFocus: Exit Function
        End If
    End If
    
    ValidData = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveSchedule() As Boolean
'------------------------------------------------
'���ܣ�����ԤԼʱ��
'������
'���أ�True -- ����ɹ���False -- ����ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim lngSendNo As Long
    Dim lngState As Long
    Dim strStartTime As String
    Dim strEndTime As String
    
    On Error GoTo err
    
    SaveSchedule = False
    
    If ValidData() = False Then
        Exit Function
    End If
    
    If schTimeTable.Label��� <> 0 Then
        
        If schTimeTable.funSaveSchedule(schTimeTable.Label��ʼʱ��, schTimeTable.Label����ʱ��, mlngOrderID, _
                schTimeTable.Label����, schTimeTable.Label���, mlngSchDeviceID, schTimeTable.Label��ʼʱ���, _
                schTimeTable.Label����ʱ���, txtNotice.Text) = False Then
                
            Exit Function
        End If
        
        'ִ�з���
        If mblnExecFee = True Then
            strSQL = "select ���ͺ�,ִ�в���ID,ִ�й��� from ����ҽ������ where ҽ��ID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ִ����Ϣ", mlngOrderID)
            If rsTemp.EOF = False Then
                strSQL = "zl_Ӱ�����ִ��(" & mlngOrderID & "," & Val(rsTemp!���ͺ�) & "," & Val(NVL(rsTemp!ִ�й���, 0)) _
                    & ",null,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & Val(rsTemp!ִ�в���ID) & ")"
                zlDatabase.ExecuteProcedure strSQL, "ԤԼʱִ�з���"
            End If
        End If
        
        '���滼����ϵ��Ϣ
        If txtPhone.ForeColor = vbRed Then
            strSQL = "select b.������Դ,b.Ӥ��,a.����id,a.����,a.�Ա�,a.����,a.�ѱ�," _
                & " a.ҽ�Ƹ��ʽ,a.����,a.����״��,a.ְҵ,a.���֤��,a.��ͥ�绰, a.��ͥ��ַ�ʱ�," _
                & " a.��������,a.��ҳid,a.��ͥ��ַ from ������Ϣ a,����ҽ����¼ b where " _
                & " a.����id=b.����id and b.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���˻�����Ϣ", mlngOrderID)
            If rsTemp.EOF = False Then
                If NVL(rsTemp!Ӥ��, 0) <> 0 Then
                    strSQL = "select  Ӥ������ ,Ӥ���Ա� , ����ʱ��  from   ������������¼ " _
                        & " Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "��ѯӤ����Ϣ", CLng(rsTemp!����ID), CLng(NVL(rsTemp!��ҳID, 0)), CLng(NVL(rsTemp!Ӥ��, 0)))

                    If rsBaby.EOF = False Then
                        strSQL = "zl_Ӱ������Ϣ_�޸�(" & NVL(rsTemp!������Դ) & "," & mlngOrderID & "," _
                        & rsTemp!����ID & ",'" & NVL(rsBaby!Ӥ������) & "','" & NVL(rsBaby!Ӥ���Ա�) & "','" _
                        & NVL(rsTemp!����) & "','" & NVL(rsTemp!�ѱ�) & "','" & NVL(rsTemp!ҽ�Ƹ��ʽ) _
                        & "','" & NVL(rsTemp!����) & "','" & NVL(rsTemp!����״��) & "','" & NVL(rsTemp!ְҵ) _
                        & "','" & NVL(rsTemp!���֤��) & "','" & NVL(rsTemp!��ͥ��ַ) _
                        & "','" & NVL(rsTemp!��ͥ�绰) & "','" & NVL(rsTemp!��ͥ��ַ�ʱ�) & "'," _
                        & zlStr.To_Date(CDate(rsBaby!����ʱ��)) & "," & NVL(rsTemp!��ҳID, 0) & "," & NVL(rsTemp!Ӥ��) _
                        & ",'" & Trim(txtPhone.Text) & "')"
                    zlDatabase.ExecuteProcedure strSQL, "���滼����Ϣ"
                    End If
                End If
                strSQL = "zl_Ӱ������Ϣ_�޸�(" & NVL(rsTemp!������Դ) & "," & mlngOrderID & "," _
                    & rsTemp!����ID & ",'" & NVL(rsTemp!����) & "','" & NVL(rsTemp!�Ա�) & "','" _
                    & NVL(rsTemp!����) & "','" & NVL(rsTemp!�ѱ�) & "','" & NVL(rsTemp!ҽ�Ƹ��ʽ) _
                    & "','" & NVL(rsTemp!����) & "','" & NVL(rsTemp!����״��) & "','" & NVL(rsTemp!ְҵ) _
                    & "','" & NVL(rsTemp!���֤��) & "','" & NVL(rsTemp!��ͥ��ַ) _
                    & "','" & NVL(rsTemp!��ͥ�绰) & "','" & NVL(rsTemp!��ͥ��ַ�ʱ�) & "'," _
                    & zlStr.To_Date(CDate(rsTemp!��������)) & "," & NVL(rsTemp!��ҳID, 0) & ",0" _
                    & ",'" & Trim(txtPhone.Text) & "')"
                zlDatabase.ExecuteProcedure strSQL, "���滼����Ϣ"
            End If
        End If
        
        'ˢ��ԤԼ������Ϣ
        Call RefreshSchInfo(True)
        mblnNewSchedule = False
        
        '��¼�����˵�ҽ��ID
        mstrModifiedOrderID = CStr(mlngOrderID)
        SaveSchedule = True
        
        '�Զ���ӡԤԼ��
        If mblnAutoPrint = True Then
            Call PrintSchedule
        End If
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DelSchedule(lngOrderID As Long)
'------------------------------------------------
'���ܣ�ɾ��ԤԼ
'������ lngOrderID -- ҽ��ID
'���أ���
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "Zl_Ӱ��ԤԼ��¼_ɾ��(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure strSQL, "ɾ�����ԤԼ"
    
    Call schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, lngOrderID)
    mblnNewSchedule = False
    
    'ˢ��ԤԼ������Ϣ
    Call RefreshSchInfo(False)
    
    '��¼�����˵�ҽ��ID
    mstrModifiedOrderID = CStr(mlngOrderID)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshSchedule(blnRefreshBaseInfo As Boolean, blnAutoNew As Boolean)
'------------------------------------------------
'���ܣ�ˢ��ԤԼ��������½�ԤԼ״̬����ÿ��ˢ�¶��½�һ��ԤԼ��ǩ����������򵥴���ˢ��
'������ blnRefreshBaseInfo -- �Ƿ�ˢ�»��߻�����Ϣ
'       blnAutoNew -- �Ƿ��Զ�����
'���أ���
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    If mblnNewSchedule = True And blnAutoNew = True Then
        Call schTimeTable.NewSchedule(mlngSchDeviceID, mschDate, mlngOrderID, False)
        'ˢ��ԤԼ������Ϣ
        Call RefreshSchInfo(blnRefreshBaseInfo)
    Else
        Call schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ChangeCalendar(dtDate As Date)
'------------------------------------------------
'���ܣ��޸�ԤԼ����������
'������dtDate -- ����������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    dpCalendar.ClearSelection
    Call dpCalendar.Select(dtDate)
    dpCalendar.EnsureVisibleSelection
    If dpCalendar.Visible = True Then
        dpCalendar.SetFocus
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintSchedule()
'------------------------------------------------
'���ܣ���ӡ��ǰԤԼ��
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsReports As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnPrinted As Boolean
    Dim lngUniFmt As Long           'ͨ�ñ����ʽ���
    
    On Error GoTo err
    
    If mblnNewSchedule = True Then
        Call MsgBox("���ȱ���ԤԼ���ٴ�ӡԤԼ����", vbInformation, "���ԤԼ��ʾ")
        Exit Sub
    End If
    
    '��ӡԤԼ��
    If mlngOrderID <> 0 Then
        '���ȼ�鱨���Ƿ�ֻ��һ����ʽ
        strSQL = "Select a.ID,a.���,b.���,b.˵�� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
        Set rsReports = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ�������ʽ", "ZL1_INSIDE_1290_01")

        If rsReports.EOF = True Then
            Call MsgBox("����ZL1_INSIDE_1290_01�������ڣ�����ϵ����Ա��Ӵ˱���", vbInformation, "���ԤԼ��ʾ")
            Exit Sub
        End If
        '����ж����ʽ������������ĿID�����Ҷ�Ӧ�ı����ʽ����
        If rsReports.RecordCount > 1 Then
            strSQL = "Select a.���� From �����ļ��б� A, ��������Ӧ�� B, ����ҽ����¼ C " _
                & " Where c.������Ŀid = b.������Ŀid And decode(c.������Դ, 3, 1, c.������Դ) = b.Ӧ�ó��� " _
                & "And b.�����ļ�id = a.ID And c.ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ļ�����", mlngOrderID)
            
            If rsTemp.EOF = False Then
            While rsReports.EOF = False And blnPrinted = False
                If NVL(rsReports!˵��) = "ͨ�ü��ԤԼ��" Then
                    lngUniFmt = rsReports!���
                End If
                
                If NVL(rsReports!˵��) = NVL(rsTemp!����) Then
                    If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & mlngOrderID, "ReportFormat=" & rsReports!���, 2) = False Then
                        Call MsgBox("����ZL1_INSIDE_1290_01���У���ʽΪ��" & NVL(rsReports!˵��) & "�ı����򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
                    Else
                        '��ӡ���˳�ѭ��
                        blnPrinted = True
                    End If
                Else
                    rsReports.MoveNext
                End If
            Wend
            End If
            '���û�У�����ҡ�ͨ�ü��ԤԼ������������ӡ
            If blnPrinted = False Then
                If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & mlngOrderID, "ReportFormat=" & lngUniFmt, 2) = False Then
                    Call MsgBox("����ZL1_INSIDE_1290_01���У���ʽΪ����ͨ�ü��ԤԼ�����ı����򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
                Else
                    blnPrinted = True
                End If
            End If
        Else
            If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & mlngOrderID, 2) = False Then
                Call MsgBox("����ZL1_INSIDE_1290_01���򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
            Else
                blnPrinted = True
            End If
        End If
        
        'д���ӡ��¼
        strSQL = "Zl_Ӱ��ԤԼ��¼_��ӡ(" & mlngOrderID & ")"
        zlDatabase.ExecuteProcedure strSQL, "���ԤԼ����ӡ"
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyPatInfo()
'------------------------------------------------
'���ܣ��޸Ĳ��˻�����Ϣ���򿪡��޸���Ϣ���ڡ�
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "Select  a.����, b.ִ�й���, b.���ͺ�,b.ִ�в���id as ִ�п���ID" _
        & " From ����ҽ����¼ A, ����ҽ������ B " _
        & " Where a.id = b.ҽ��id  And a.id = [1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�޸���Ϣ����ѯ������Ϣ", mlngOrderID)
    
    With frmRISRequest
        .mstrPrivs = gstrPrivs
        .mlngModul = glngModul
        .mlngSendNo = rsTemp!���ͺ�
        .mlngAdviceId = mlngOrderID
        .mstrPatientName = NVL(rsTemp!����)
        .mintEditMode = IIf(NVL(rsTemp!ִ�й���, 0) > 1, 3, 1) '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = rsTemp!ִ�п���ID
        .mstrCur���� = "����"
        
        Call frmRISRequest.InitMvar(False)
        .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
    End With
    Call RefreshForm
    Call mfrmParent.RefreshList 'ˢ�¸�����
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshForm()
'------------------------------------------------
'���ܣ�ˢ�´�������
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Call LoadData
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshSchInfo(blnRefreshBaseInfo As Boolean)
'------------------------------------------------
'���ܣ�ˢ�»��ߵĻ�����Ϣ
'������ blnRefreshBaseInfo -- ˢ�»�����Ϣ
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    lblInfo(4).Caption = ""
    lblInfo(5).Caption = ""
        
    If blnRefreshBaseInfo = True Then
        lblInfo(1).Caption = ""
        lblInfo(2).Caption = ""
        txtPhone.Text = ""
        txtPhone.Locked = True
        txtPhone.ForeColor = vbWindowText
        txtNotice.Text = ""
        txtNotice.Locked = True
        txtNotice.ForeColor = vbWindowText
        mlngPatSource = 0
    End If
        
    If blnRefreshBaseInfo = True Then
        strSQL = "Select A.����, A.�Ա�, A.����, " _
            & " DECODE(A.������Դ, 2, 'סԺ', 1, '����', 4, '���', '����') As ������Դ����,A.������Դ, " _
            & " B.�ֻ��� ,Nvl(B.��ͥ��ַ, B.��ϵ�˵�ַ) ��ַ, A.ҽ������,Nvl(a.Ӥ��, 0) As Ӥ�� " _
            & " From ����ҽ����¼ A, ������Ϣ B " _
            & " Where a.ID = [1] And a.����ID = b.����ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ������Ϣ", mlngOrderID)
        
        If Not rsTemp.EOF Then
            '����Ӥ������
            If rsTemp!Ӥ�� <> 0 Then
                strSQL = "Select A.����ʱ��,Nvl(B.Ӥ������, A.���� || '֮��' || Trim(To_Char(B.���, '9'))) As Ӥ������, B.Ӥ���Ա�, B.����ʱ��" & vbNewLine & _
                                 "  From ����ҽ����¼ A, ������������¼ B " & vbNewLine & _
                                 "  Where a.����ID = b.����ID  And b.��� = [2] And a.ID = [1]"
                            
                Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡӤ����Ϣ", mlngOrderID, CLng(rsTemp!Ӥ��))
                
                lblInfo(0).Caption = "Ӥ��������" & rsBaby!Ӥ������ & "   �Ա�" & rsBaby!Ӥ���Ա�
                lblInfo(1).Caption = "����ʱ�䣺" & rsBaby!����ʱ�� & "   ��Դ��" & rsTemp!������Դ����
            Else
                lblInfo(0).Caption = "������" & rsTemp!���� & "   �Ա�" & rsTemp!�Ա�
                lblInfo(1).Caption = "���䣺" & rsTemp!���� & "   ��Դ��" & rsTemp!������Դ����
            End If
            
            lblInfo(2).Caption = rsTemp!ҽ������
            txtPhone.Text = NVL(rsTemp!�ֻ���)
            mlngPatSource = rsTemp!������Դ
        Else
            lblInfo(0).Caption = "������         �Ա�"
            lblInfo(1).Caption = "���䣺         ��Դ��"
            lblInfo(2).Caption = ""
            txtPhone.Text = ""
            mlngPatSource = 0
        End If
    End If
    
    strSQL = "select ԤԼ�豸ID,ԤԼ�豸����,ԤԼ��ʼʱ��,ԤԼ����ʱ��,���ע�� from Ӱ��ԤԼ��¼ where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ��¼", mlngOrderID)
    If Not rsTemp.EOF Then
        txtNotice.Text = NVL(rsTemp!���ע��)
        lblInfo(4).Caption = "ԤԼ����: " & Format(rsTemp!ԤԼ��ʼʱ��, "yyyy-mm-dd")
        lblInfo(5).Caption = "ԤԼʱ�䣺" & Format(rsTemp!ԤԼ��ʼʱ��, "HH:MM:SS") _
            & " - " & Format(rsTemp!ԤԼ����ʱ��, "HH:MM:SS")
        mschDate = Format(rsTemp!ԤԼ��ʼʱ��, "YYYY-MM-DD")
        mlngSchDeviceID = rsTemp!ԤԼ�豸ID
        
        '����ԤԼ�豸
        lblSchDevice.Caption = rsTemp!ԤԼ�豸����
        
        mblnISScheduled = True
    Else
        txtNotice.Text = ""
        lblInfo(4).Caption = "ԤԼ����: "
        lblInfo(5).Caption = "ԤԼʱ�䣺"
        mblnISScheduled = False
    End If
       
    lblSchDevice.Visible = IIf(cboSchDevice.ListCount > 0, False, True)
    cboSchDevice.Visible = IIf(cboSchDevice.ListCount > 0, True, False)
   
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub NewSchedule()
'------------------------------------------------
'���ܣ��½�ԤԼ
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select ID from Ӱ��ԤԼ��¼ where ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ��¼", mlngOrderID)
    
    If rsTemp.EOF = False Then
        If MsgBox("�½�ԤԼ֮ǰ�����Զ�ɾ��������ԭ�е�ԤԼ��Ϣ��" & vbCrLf & vbCrLf & "�Ƿ�ȷ��ɾ��ԭ�е�ԤԼ��Ϣ��", vbYesNo, "���ԤԼ��ʾ") = vbNo Then
            Exit Sub
        Else
            Call DelSchedule(mlngOrderID)
        End If
    End If
    
    mblnNewSchedule = True
    Call RefreshSchedule(False, True)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshCalendar()
'------------------------------------------------
'���ܣ�ˢ������
'������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    mstrSchRestDate = RefeshSchRestDay(mlngOrderID, mlngSchDeviceID, dpCalendar.LastVisibleDay)
    
    dpCalendar.RedrawControl
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
