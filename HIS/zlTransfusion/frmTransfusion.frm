VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmTransfusion 
   BackColor       =   &H8000000C&
   Caption         =   "������Һע�����"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmTransfusion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picReadyReceive 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6480
      Width           =   3615
      Begin zlIDKind.PatiIdentify ptiReadyReceive 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmTransfusion.frx":6852
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         ShowPropertySet =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
   End
   Begin VB.PictureBox picTmp 
      Height          =   255
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5550
      TabIndex        =   30
      Top             =   105
      Width           =   1905
   End
   Begin VB.Timer tmrAutoReady 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1710
      Top             =   225
   End
   Begin VB.PictureBox picRecord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   45
      ScaleHeight     =   1035
      ScaleWidth      =   3705
      TabIndex        =   17
      Top             =   5070
      Width           =   3705
      Begin XtremeReportControl.ReportControl rptRecord 
         Height          =   780
         Left            =   60
         TabIndex        =   18
         Top             =   75
         Width           =   3555
         _Version        =   589884
         _ExtentX        =   6271
         _ExtentY        =   1376
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   5115
      ScaleHeight     =   585
      ScaleWidth      =   7740
      TabIndex        =   0
      Top             =   465
      Width           =   7740
      Begin VB.Frame fraInfo 
         Height          =   645
         Left            =   15
         TabIndex        =   3
         Top             =   -60
         Width           =   7695
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   20
            Top             =   375
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   15
            Left            =   570
            TabIndex        =   19
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   13
            Left            =   4350
            TabIndex        =   16
            Top             =   375
            Width           =   90
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "���:"
            Height          =   180
            Index           =   12
            Left            =   3855
            TabIndex        =   15
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   11
            Left            =   2175
            TabIndex        =   14
            Top             =   375
            Width           =   1605
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   10
            Left            =   1695
            TabIndex        =   13
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   7
            Left            =   4350
            TabIndex        =   12
            Top             =   165
            Width           =   1290
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�:"
            Height          =   180
            Index           =   6
            Left            =   3855
            TabIndex        =   11
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   3285
            TabIndex        =   10
            Top             =   165
            Width           =   525
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   4
            Left            =   2775
            TabIndex        =   9
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   2175
            TabIndex        =   8
            Top             =   165
            Width           =   540
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�:"
            Height          =   180
            Index           =   2
            Left            =   1695
            TabIndex        =   7
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   6
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   165
            Width           =   450
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7455
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransfusion.frx":693C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15716
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   6240
      Left            =   5130
      TabIndex        =   2
      Top             =   1080
      Width           =   6600
      _Version        =   589884
      _ExtentX        =   11642
      _ExtentY        =   11007
      _StockProps     =   64
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   4485
      Left            =   165
      ScaleHeight     =   4485
      ScaleWidth      =   4890
      TabIndex        =   1
      Top             =   480
      Width           =   4890
      Begin VB.PictureBox picQueue0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   3780
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3555
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue0 
            Height          =   3270
            Left            =   195
            TabIndex        =   42
            Top             =   270
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.Frame fraWhere 
         Height          =   1300
         Left            =   45
         TabIndex        =   33
         Top             =   -45
         Width           =   4815
         Begin zlIDKind.IDKindNew idkSelect 
            Height          =   270
            Left            =   120
            TabIndex        =   39
            Top             =   900
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   476
            IDKindStr       =   "��|��������￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|0|0|0|0|0|;IC|IC����|1|0|0|0|0|;��|�����|0|0|0|0|0|"
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
            ShowPropertySet =   -1  'True
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            BackColor       =   -2147483633
            SaveRegType     =   4
         End
         Begin VB.ComboBox cboDate 
            Height          =   300
            ItemData        =   "frmTransfusion.frx":71CE
            Left            =   795
            List            =   "frmTransfusion.frx":71E1
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   540
            Width           =   2600
         End
         Begin VB.TextBox txtInfo 
            Height          =   270
            Left            =   1095
            TabIndex        =   40
            Top             =   900
            Width           =   2580
         End
         Begin VB.CommandButton cmdOk 
            Height          =   270
            Left            =   3405
            Picture         =   "frmTransfusion.frx":7213
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   540
            Width           =   315
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   795
            TabIndex        =   35
            Text            =   "cboDept"
            Top             =   195
            Width           =   2910
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʱ��(&T)"
            Height          =   180
            Left            =   135
            TabIndex        =   36
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lblB 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            Height          =   180
            Left            =   135
            TabIndex        =   34
            Top             =   255
            Width           =   630
         End
      End
      Begin VB.PictureBox picQueue7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   1080
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3585
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue7 
            Height          =   3270
            Left            =   150
            TabIndex        =   32
            Top             =   255
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picQueue6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   3285
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3255
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue6 
            Height          =   3270
            Left            =   135
            TabIndex        =   29
            Top             =   150
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picQueue5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   2295
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2865
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue5 
            Height          =   3270
            Left            =   210
            TabIndex        =   27
            Top             =   675
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtNo5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   45
            Top             =   225
            Width           =   1995
         End
         Begin VB.Label lblNo5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Һŵ��س���ɴ���"
            Height          =   180
            Left            =   15
            TabIndex        =   46
            Top             =   0
            Width           =   1980
         End
      End
      Begin VB.PictureBox picQueue1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   855
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2565
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue1 
            Height          =   3270
            Left            =   150
            TabIndex        =   25
            Top             =   630
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtNo1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   43
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label lblNo1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Һŵ��س������Һ"
            Height          =   180
            Left            =   165
            TabIndex        =   44
            Top             =   75
            Width           =   1980
         End
      End
      Begin VB.PictureBox picQueueAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   225
         ScaleHeight     =   2775
         ScaleWidth      =   3630
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2430
         Width           =   3660
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   3270
            Left            =   105
            TabIndex        =   23
            Top             =   420
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl tbcList 
         Height          =   960
         Left            =   225
         TabIndex        =   21
         Top             =   1335
         Width           =   3630
         _Version        =   589884
         _ExtentX        =   6403
         _ExtentY        =   1693
         _StockProps     =   64
      End
      Begin VB.Timer timRefresh 
         Interval        =   1000
         Left            =   3045
         Top             =   2745
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   2010
         Top             =   2625
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":DA65
               Key             =   "δִ��"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":DFFF
               Key             =   "��ִ��"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":E599
               Key             =   "�ܾ�ִ��"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":EB33
               Key             =   "����ִ��"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":F0CD
               Key             =   "�ѱ���"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":F667
               Key             =   "Calling"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTransfusion.frx":15EC9
      Left            =   675
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private patiList As New cPatients '�����б���
Private ObjOutNurse As New OutNurses '���ﻤʿ�б�
Private mobjPopupInfo As CommandBar

Private mstr�Һŵ� As String  '��ǰ�û�Ψһ��ʶ
Private mstrPrivs As String
Private mlngModul As Long

Private mlngPreDept As Long '�ϴο���ID
Private mstr��λ As String
Private mDateBegin As Date '��ʼʱ��
Private mdateEnd As Date  '����ʱ��
Private mintRefresh As Integer
Public mblnƤ����֤ As Boolean
'�Ӵ���
Private WithEvents mclsSeating As clsDockSeating
Attribute mclsSeating.VB_VarHelpID = -1
Private mfrmLeaveMedi As frmLeaveMedi 'ҩƷ�Ĵ�
Private mfrmRecord As frmRecord '��ִ����Ŀ
Public mobjRecord As ExecRecord

Private mcolSubForm As Collection 'subTab�Ӵ��弯��
'������Ŀ�б�
Private Enum rptCOL
    rptCOL_ִ�з��� = 0
    rptCOL_�ӵ�ʱ�� = 1
    rptCOL_��ҩ�� = 2
    rptCOL_�ӵ��� = 3

    rptCOL_��ʱ = 4
    rptCOL_��ϵ�� = 5
    rptCOL_���� = 6
    rptCOL_��ˮ�� = 7
End Enum

'Private mfrmActive As Form
Private mintRow As Integer    '��ǰ��,���ڽк�
Private Type SiblingRow
    PrivRowIndex As Integer
    PrivRow�Һŵ� As String
    PrivRow״̬ As String

    curRow�Һŵ� As String
    curRow״̬ As String
    curRowIndex As Integer

    nextRow�Һŵ� As String
    nextRow״̬ As String
    nextRowIndex As Integer
End Type

Private mintPatirow As Integer 'ˢ��ʱ,���¶�λ
Private mintRecordRow As Integer 'ˢ��ʱ,���¶�λ

Private mintFindType As Integer '�������� 0-���￨,1-�����,2-�Һŵ�,3-����,4-���֤,5-IC��

'Private mstrIDCard As String '����Զ�ˢ���������֤��
'Private WithEvents mobjIDCard As clsIDCard '���֤����
'Private mobjICCard As Object 'IC������
Private mstrQueueTab As String  '��ǰ����ҳ��
Private mintLastFind As Integer     '���Ҵ���

Private mblnLiquid  As Boolean  '�Ƿ�����Һ����
'Private mblnPuncture As Boolean '�Ƿ��д�������  ��/��׼�б����б����Ҽ���Ϊ�д�������
'Private mblnCall    As Boolean  '�Ƿ��к�������

Private mblnVisits As Boolean   '�Ƿ���Ѳ������
    
Private mobjSquareCard   As Object   'һ��ͨ���� add by 2011-08-23
Private mstrSquareCards As String    'һ��ͨ����
Private mintCards As Integer         'һ���ܿ�������
Private mstrPatiKey As String        '���˲�����Ϣ����
Private mblnReadCard As Boolean
Private mintPatiIdentify As String
Private mfrmTimeCall As Form      '�Ŷӽкŵ���ѭ����

Private Const MLNG_INFO As Long = 100000
Private Const MSTR_MODE As String = "��|�Һŵ�|0;��|���￨|1;��|�����|0;��|����|0;��|���֤��|0;IC|IC��|1"

Private Sub ShowPage()
    '����ϵͳ������ʾ��Һҳ��
    Dim i As Integer
    Dim strPara As String
    
    '85046
    'mblnLiquid = GetDeptInListPara("������Һ_��Һ�����б�", mlngPreDept)
    strPara = zlDatabase.GetPara("����Һ�����б�", glngSys, mlngModul, "")
    mblnLiquid = InStr("," & strPara & ",", "," & mlngPreDept & ",") > 0
    
'    mblnPuncture = True '���̹��ܱ����У�����д��ʼʱ�䣬��ʼ����Ա��
    ' GetDeptInListPara("������Һ_��׼�����б�", mlngPreDept) Or GetDeptInListPara("������Һ_�򵥴����б�", mlngPreDept)
    
    '85046 ����ȡ���ò����Ŀ���
    'mblnCall = GetDeptInListPara("������Һ_���п����б�", mlngPreDept)
    
    With Me.tbcList
        For i = 1 To .ItemCount - 1
            If (Not mblnLiquid And .Item(i).Tag = "����Һ") Or .Item(i).Tag = "��ִ��" Then
                .Item(i).Visible = False
            Else
                .Item(i).Visible = True
            End If
        Next
    End With
End Sub

Private Sub cboDate_Click()
    mdateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    If Trim(txtInfo.Text) = "" Then
        Select Case cboDate.ListIndex
            Case 1              '����
                mDateBegin = Format(mdateEnd - 1, "yyyy-MM-dd 00:00:00")
            Case 2              '������
                mDateBegin = Format(mdateEnd - 2, "yyyy-MM-dd 00:00:00")
            Case 3              '��һ��
                mDateBegin = Format(mdateEnd - 6, "yyyy-MM-dd 00:00:00")
            Case 4              '��ʮ��
                mDateBegin = Format(mdateEnd - 9, "yyyy-MM-dd 00:00:00")
            Case Else           '����
                mDateBegin = Format(mdateEnd, "yyyy-MM-dd 00:00:00")
        End Select
    Else
        mDateBegin = Format(mdateEnd - 364, "yyyy-MM-dd 00:00:00")
    End If
End Sub

Private Sub cboDept_Click()
    If cboDept.ListCount <= 0 Then Exit Sub
    If cboDept.ItemData(cboDept.ListIndex) = mlngPreDept Then Exit Sub
    mlngPreDept = cboDept.ItemData(cboDept.ListIndex)
    
    Call ShowPage
    Call ObjOutNurse.getOutNurse(mlngPreDept)  '��ʼ�������һ�ʿ�б�
    '��ʼ��patients��
    Call ShowLblInfo("")
    Call ShowReport
    mstr�Һŵ� = ""
    mintRow = 0
    ShowPatiList
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex

    '֧��¼����Ҽ��롢���ơ�����
    If KeyAscii = vbKeyReturn Then
        If Trim(cboDept.Text) = "" Then Exit Sub
        
        Dim intIndex As Integer
        Dim strText As String, strSQL As String
        Dim rsTmp As ADODB.Recordset
        Dim vRect As RECT
        Dim blnCanel As Boolean
        
        KeyAscii = 0
        intIndex = cboDept.ListIndex
        strText = UCase(cboDept.Text) & "%"
        If Val(zlDatabase.GetPara("����ƥ��", , "0")) = 1 Then
            strText = UCase(cboDept.Text) & "%"
        Else
            strText = "%" & UCase(cboDept.Text) & "%"
        End If
        
        If InStr(mstrPrivs, ";���п���;") > 0 Then
            strSQL = "Select /*+ Rule*/ Distinct a.Id, a.����, a.���� " & vbCr & _
                     "From ���ű� A, ��������˵�� B " & vbCr & _
                     "Where b.����id = a.Id And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.������� In (1, 3) " & vbCr & _
                     "  And b.�������� In ('����', '�ٴ�') And (a.վ�� = [2] Or a.վ�� Is Null) " & vbCr & _
                     "  And (A.���� Like [3] Or A.���� Like [3] Or A.���� Like [3]) " & vbCr & _
                     "Order By a.���� "
        Else
            strSQL = "Select /*+ Rule*/ Distinct a.Id, a.����, a.���� " & vbCr & _
                     "From ���ű� A, ��������˵�� B, ������Ա C " & vbCr & _
                     "Where b.����id = a.Id And a.Id = c.����id And c.��Աid = [1] " & vbCr & _
                     "  And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.������� In (1, 3) " & vbCr & _
                     "  And b.�������� In ('����', '�ٴ�') And (a.վ�� = [2] Or a.վ�� Is Null) " & vbCr & _
                     "  And (A.���� Like [3] Or A.���� Like [3] Or A.���� Like [3]) " & vbCr & _
                     "Order By a.���� "
        End If
        On Error GoTo errHandle
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", UserInfo.ID, zl9ComLib.gstrNodeNo, strText)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Call FindCboIndex(cboDept, rsTmp!ID)
            Else
                rsTmp.Close
                
                vRect = zlControl.GetControlRect(cboDept.hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ָ������", False, "", "ѡ�����", False, False, True, _
                                         vRect.Left, vRect.Top, cboDept.Height, blnCanel, True, True, _
                                         UserInfo.ID, zl9ComLib.gstrNodeNo, strText)
                If blnCanel = False Then
                    Call FindCboIndex(cboDept, rsTmp!ID)
                    rsTmp.Close
                Else
                    cboDept.ListIndex = Val(cboDept.Tag)
                End If
                
            End If
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
            rsTmp.Close
        End If
        
        Call zlCommFun.PressKey(vbKeyTab)
        
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim strStat As String
    Dim objPati As cPatient
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
'    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 99 '���ҷ�ʽ
'        mintFindType = Val(Right(Control.ID, 2)) - 1
'
'        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_FindType)
'        objControl.Caption = "����" & Split(Control.Caption, "(")(0) & "����"
'
'        cbsMain.RecalcLayout
'        txtFind.Tag = Control.Parameter
'        txtFind.Text = ""
'        txtFind.SetFocus
        
'    Case MLNG_INFO + 1 To MLNG_INFO + 99     'ָ����ȡ������Ϣ
'        lblBill.Caption = Split(Control.Caption, "(")(0)
'        lblBill.Caption = lblBill.Caption & "��"
'        lblBill.Tag = Control.ID - MLNG_INFO
        
        
    Case conMenu_View_Find '����
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '��ʱ��Ҫ��λһ��
            If txtFind.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            txtFind.SetFocus
        End If
        
'    Case conMenu_View_FindNext '������һ��
'        If txtFind.Text = "" And mstrIDCard = "" Then
'            txtFind.SetFocus
'        Else
'            Call ExecuteFindPati(True, IIf(txtFind.Text = "", mstrIDCard, ""))
'        End If
'    Case conMenu_View_ReadIC '��IC��
'        If Not mobjICCard Is Nothing Then
'            txtFind.Text = mobjICCard.Read_Card(Me)
'            If txtFind.Text <> "" Then Call ExecuteFindPati
'        End If
    Case conMenu_View_Refresh 'ˢ��
        Call cmdOk_Click
        
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '�۵�������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend 'չ��������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    
    Case conMenu_Manage_ThingAdd
        '�ӵ�
        Call thingAdd
    Case conMenu_File_Parameter
        '��������
         Call ParameterSetup
    Case conMenu_File_RoomSet
        '����̨����
        frmPunctureDeskSet.ShowMe mlngPreDept
    Case conMenu_Manage_Call
        '����
        Call Calling(2)
        If mstr�Һŵ� <> "" Then Call CallOnePlay(mstr�Һŵ�)
'    Case conMenu_Manage_CallNext
'        '��һ��
'        Call Calling(1)
'    Case conMenu_Manage_CallPrevious
'        '��һ��
'        Call Calling(-1)
    Case conMenu_Manage_Up
        '����
        Call rptQueueMove(-1)
    Case conMenu_Manage_Down
        '����
        Call rptQueueMove(1)
        
    Case conMenu_Manage_Discard
        '����
        If UpdateState("2-����") Then
            Set objPati = patiList.Item(mstr�Һŵ�)
            SaveOperLog mlngPreDept, objPati, QUEUE, "���Ų���"
        End If
    Case conMenu_Manage_Recall
        '�ٻ�
        If UpdateState("7-ִ����") Then
            Set objPati = patiList.Item(mstr�Һŵ�)
            SaveOperLog mlngPreDept, objPati, QUEUE, "�ٻز���"
        End If
    Case conMenu_Manage_Untread
        '�˺�
        If UpdateState("3-�˺�") Then
            Set objPati = patiList.Item(mstr�Һŵ�)
            SaveOperLog mlngPreDept, objPati, QUEUE, "�˺Ų���"
        End If
    Case conMenu_Manage_TagEnd
        '��Ϊȫ��
        If UpdateState("4-����") Then
            Set objPati = patiList.Item(mstr�Һŵ�)
            SaveOperLog mlngPreDept, objPati, QUEUE, "��������"
        End If
    Case conMenu_Edit_Transf_Liquid
        '��Һ
        Call LiquidAndPlay
    Case conMenu_Edit_Transf_Puncture
        '����
        Call Puncture
    Case conMenu_Edit_Bed_Modify
        '����״̬
        strStat = patiList.Item(mstr�Һŵ�).�Ŷ�״̬
        If frmChangeStat.ShowMe(strStat, mblnLiquid) Then
            If UpdateState(strStat) Then
                Set objPati = patiList.Item(mstr�Һŵ�)
                SaveOperLog mlngPreDept, objPati, QUEUE, "����״̬Ϊ" & strStat
            End If
        End If
    Case conMenu_Queue_Setup    '���в�������
        Call QueueSetup(Me)
    Case conMenu_View_Show      '�鿴��־
        Call frmTransfusionLog.ShowMe(mlngPreDept, mstr�Һŵ�)
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If mstr�Һŵ� <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "����ID=" & patiList.Item(mstr�Һŵ�).����ID)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        ElseIf Me.tbcSub.Selected.Tag = "��λ����" Then
            Call mclsSeating.zlExecuteCommandBars(Control)
        ElseIf tbcSub.Selected.Tag = "ִ����Ŀ" Then
            Call mfrmRecord.zlExecuteCommandBars(Control)
        ElseIf tbcSub.Selected.Tag = "ҩƷ�Ĵ�" Then
            Call mfrmLeaveMedi.zlExecuteCommandBars(Control, Me)
        End If
    End Select
    
    'ˢ��TabControl
    On Error Resume Next
    Select Case Control.ID
        Case conMenu_Edit_Transf_Liquid, conMenu_Edit_Transf_Puncture, conMenu_Edit_Bed_Modify, _
            conMenu_Manage_Discard, conMenu_Manage_Recall, conMenu_Manage_Untread, conMenu_Manage_TagEnd
            If Not tbcList.Selected Is Nothing Then Call tbcList_SelectedChanged(tbcList.Selected)
    End Select
    Err.Clear
End Sub
Private Sub LiquidAndPlay()
    '��Һ������
    Dim strStat As String, strErr As String, i As Integer
    Dim objPati As cPatient
    
    strStat = Liquid(mlngPreDept, mstr�Һŵ�, patiList, strErr)
    If strErr <> "" Then
        MsgBox strErr, vbInformation, Me.Caption
        Exit Sub
    End If
    If strStat <> "" Then
        If Not mobjRecord Is Nothing Then
            For i = 1 To mobjRecord.Count
                Call mobjRecord.Item(i).SaveDispenseUser(1, zlDatabase.Currentdate, UserInfo.����)
            Next
            
            Call ShowLblInfo(mstr�Һŵ�)
        End If
        
        Set objPati = patiList.Item(mstr�Һŵ�)
        If UpdateState(strStat) Then
            SaveOperLog mlngPreDept, objPati, QUEUE, "��Һ�����״̬Ϊ" & strStat
        Else
            SaveOperLog mlngPreDept, objPati, QUEUE, "��Һ��δ����״̬"
        End If
    End If
    If strStat = "5-������" Then
        Call CallPlay(mstr�Һŵ�)
    End If
End Sub
Private Sub Puncture()
    '���̲���
    Dim strSQL As String, i As Integer, Y As Integer
    Dim dateS As Date, dateE As Date, blnExitFor As Boolean, strGroupKey As String
    Dim intOneOrTow As Integer, rsTmp As ADODB.Recordset, lngTaiID As Long, strNextNo As String
    Dim objPati As cPatient
    
    On Error GoTo hErr
    
    If mstr�Һŵ� <> "" Then
        Set objPati = patiList.Item(mstr�Һŵ�)
        If objPati Is Nothing Then
            MsgBox "��ǰ�������ҵ����ܾ����̣�", vbInformation, gstrSysName
            Exit Sub
        End If
                 
        strSQL = "ZL_���ﴩ��̨_Puncture(" & mlngPreDept & "," & objPati.����ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If UpdateState("7-ִ����") Then
            SaveOperLog mlngPreDept, objPati, QUEUE, "���̺����״̬Ϊ7-ִ����"
        Else
            SaveOperLog mlngPreDept, objPati, QUEUE, "���̺�δ����״̬"
        End If
        
        '׼�����д������ˡ�ͨ�����˿��ҡ�����̨���ҵ�����������ID�������ҵ����˵ĹҺŵ�����ҳID
        strSQL = "Select a.����id, a.�Һŵ�, a.��ҳid " & vbNewLine & _
                 "From �ŶӼ�¼ A, ���ﴩ��̨ B " & vbNewLine & _
                 "Where a.����id = b.����id And a.����id = b.��������id And a.����id = [1] And a.����̨ = [2] and a.״̬ = 5 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ׼�����еĵȴ�����", mlngPreDept, objPati.����̨)
        If rsTmp.EOF = False Then
            If Val(zlCommFun.NVL(rsTmp!��ҳid)) <= 0 Then
                '����
                strNextNo = zlCommFun.NVL(rsTmp!�Һŵ�)
            Else
                '��������
                strNextNo = zlStr.FormatString("[1]_[2]", rsTmp!����ID, rsTmp!��ҳid)
            End If
        Else
            strNextNo = ""
        End If
        rsTmp.Close
        
        '����׼�����̵Ĳ���
        If strNextNo <> "" Then
            Call CallPlay(strNextNo)
        End If
        
        dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dateE = Format(dateS, "yyyy-MM-dd 23:59:59")
        If Not mobjRecord Is Nothing Then
            blnExitFor = False
            For i = 1 To mobjRecord.Count
                For Y = 1 To mobjRecord.Item(i).Count
                     If mobjRecord.Item(i).Item(Y).ִ�з��� = "1-��Һ" And _
                        mobjRecord.Item(i).Item(Y).��� = 1 And _
                        mobjRecord.Item(i).ִ��ʱ�� >= dateS And mobjRecord.Item(i).ִ��ʱ�� <= dateE And _
                        mobjRecord.Item(i).Item(Y).ִ���� = "" Then
                            '�����һ��ҩ��û����ִ���ˣ�����
                            dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                            strGroupKey = mobjRecord.Item(i).Item(Y).ִ��ҽ��ID & "_" & mobjRecord.Item(i).Item(Y).���ͺ�
                            Call mobjRecord.Item(i).ExecStart(1, strGroupKey, dateS, UserInfo.����)
                            
                            '
                            Call ExecComplt(CStr(mobjRecord.Item(i).��ˮ��), strGroupKey)
                            blnExitFor = True
                            Exit For
                     End If
                Next
                If blnExitFor Then Exit For
            Next
        End If
    Else
        MsgBox "��ѡ��һ����¼����ִ�д˲���!", vbQuestion, Me.Caption
    End If
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CallPlay(ByVal strNO As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    '--- ˳��
     
'    If mblnCall Then Exit Sub                       '�����Ҳ��ں����б���
    Call QueueCall(strNO, mlngPreDept, patiList.Item(strNO))
    
End Sub

Private Sub CallOnePlay(ByVal strNO As String)
    '�����ĺ��к���ʾ
    
    'ȡ�����߼�����Ϊ����������Һ���û����ã�̨ʽ���˶����Բ�������ָ������
    'If Not mblnCall Then Exit Sub                       '�����Ҳ��ں����б���
    
    '0-�������ҽ�������д�������
    Dim lngNo As Long
    Dim bln���� As Boolean
    
    lngNo = Get��ˮ��()
    If lngNo > 0 Then
        If Not mobjRecord.Item(CStr(lngNo)) Is Nothing Then
            bln���� = (mobjRecord.Item(CStr(lngNo)).ִ�з��� Like "0*")
            If bln���� Then
                '�������ҽ�������д���̨��������
                Call QueueOnePlay(strNO, "�롢" & patiList.Item(strNO).���� & "����������", lngNo)
            End If
        End If
    End If
    If bln���� = False Then
        Call QueueOnePlay(strNO, "�롢" & patiList.Item(strNO).���� & "������" & patiList.Item(strNO).����̨ & "�Ŵ�����Һ", lngNo)
    End If
    
End Sub

Private Sub thingAdd(Optional ByVal bytType As Byte = 0)
'���ܣ��ӵ�����
'������
'  bytType�� 0-����ӵ���ť��ʽ��1-�Զ����ýӵ������ӵ����ˣ�

    Dim strJZK As String, strName As String
    Dim lngDeptID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '��鵱ǰ���Ҵ���̨����
    lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    strSQL = "Select Count(1) Rec from ���ﴩ��̨ where ����id = [1] and ��Ч = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ﴩ��̨", lngDeptID)
    lngDeptID = rsTemp!Rec
    rsTemp.Close
    If lngDeptID <= 0 Then
        MsgBox "��ǰ����δ���ô���̨������Ч�Ĵ���̨��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�ӵ�
    strJZK = lblinfo(15).Caption    '����
    strName = lblinfo(1).Caption    '����
    
    If bytType = 1 Then
        Call frmReady.ShowIncepBill(bytType, mlngPreDept, cboDept.List(cboDept.ListIndex), mstr��λ, mDateBegin, mdateEnd, patiList, _
                    ObjOutNurse, Me, , , , Me.ptiReadyReceive)
    Else
        Call frmReady.ShowIncepBill(bytType, mlngPreDept, cboDept.List(cboDept.ListIndex), mstr��λ, mDateBegin, mdateEnd, patiList, _
                    ObjOutNurse, Me, Me.txtInfo.Text, strJZK, strName)
    End If
        
    '�����ѱ仯��ˢ����ʾ        '������λ��ʾ
    mlngPreDept = -1
    mdateEnd = CDate(0)
    Call cboDept_Click

    
'    If mstr�Һŵ� = "" Then Exit Sub
'    With rptPati.SelectedRows(0)
'        If Not .GroupRow Then
'            If mstr�Һŵ� <> "" And InStr("1-����Һ,0-δ�ӵ�", .Record(col_�Ŷ�״̬).Value) > 0 Then
'                '-- ����,��������,��λ,��ʼ����,��������,patient��,����
'                If frmReady.InceptBill(mlngPreDept, cboDept.List(cboDept.ListIndex), mstr��λ, mDateBegin, mdateEnd, _
'                                       patiList.Item(mstr�Һŵ�), patiList.mSeatings, ObjOutNurse, Me) Then
'                'ˢ����λ
'                    mlngPreDept = -1
'                    mdateEnd = CDate(0)
'                    Call cboDept_Click
'                    Call rptPati_SelectionChanged
'                End If
'            End If
'        End If
'    End With

End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub

    '�Ҽ������˵�
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType '���ҷ�ʽ
'        With CommandBar.Controls
'            If .Count = 0 Then
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "���￨(&1)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "�����(&2)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "���ݺ�(&3)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "��  ��(&4)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 5, "���֤(&5)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 6, "�ɣÿ�(&6)"
'            End If
'        End With
        
    Case Else
        If tbcSub.Selected.Tag = "��λ����" Then
           Call mclsSeating.zlPopupCommandBars(CommandBar)
        End If
        
        If tbcSub.Selected.Tag = "ִ����Ŀ" Then
            Call mfrmRecord.zlPopupCommandBars(CommandBar)
        End If
        
        If tbcSub.Selected.Tag = "ҩƷ�Ĵ�" Then
            Call mfrmLeaveMedi.zlPopupCommandBars(CommandBar)
        End If
        
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.picInfo
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With

    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = picInfo.Top + picInfo.Height
        .Height = lngBottom - .Top - stbThis.Height
    End With

End Sub

Private Function GetRowState(objRpt As ReportControl, blnQueue As Boolean) As SiblingRow
    'ȡָ��RPT�ؼ���������״̬
    'blnQueue = ȡ������ ,����ȡ״̬��
    Dim intCurRow As Integer
    If blnQueue Then
        intCurRow = mintRow
        If intCurRow > 0 Then
            GetRowState = SiblingRowState(objRpt, intCurRow)
        End If
    Else
        If objRpt.SelectedRows.Count > 0 Then
            intCurRow = objRpt.SelectedRows(0).Index
            If intCurRow >= 0 Then
                GetRowState = SiblingRowState(objRpt, intCurRow)
            End If
        End If
    End If
End Function
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Dim TCurRowState As SiblingRow  '�ƶ���
    Dim TQueueRowState As SiblingRow
    Dim blnEnabled As Boolean
    Dim lng��ˮ�� As Long, intItem As Integer
    
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        TCurRowState = GetRowState(rptQueue0, False)
        TQueueRowState = GetRowState(rptQueue0, True)
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        'ȡ��ǰ�м�������״̬
        TCurRowState = GetRowState(rptQueue1, False)
        TQueueRowState = GetRowState(rptQueue1, True)
    ElseIf tbcList.Selected.Tag = "������" Then
        TCurRowState = GetRowState(rptQueue5, False)
        TQueueRowState = GetRowState(rptQueue5, True)
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        TCurRowState = GetRowState(rptQueue6, False)
        TQueueRowState = GetRowState(rptQueue6, True)
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        TCurRowState = GetRowState(rptQueue7, False)
        TQueueRowState = GetRowState(rptQueue7, True)
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        TCurRowState = GetRowState(rptPati, False)
        TQueueRowState = GetRowState(rptPati, True)
    End If
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = rptPati.SelectedRows(0).Expanded
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '�۵�/չ����
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
'    Case conMenu_View_FindType '���ҷ�ʽ
'        If Control.Parent Is cbsMain.ActiveMenuBar Then
'            If mintFindType <= 5 Then
'                Control.Caption = "��" & Decode(mintFindType, 0, "���￨", 1, "�����", 2, "�Һŵ�", 3, "����", 4, "���֤", 5, "�ɣÿ�") & "����"
'            Else
'                Control.Caption = ""
'            End If
'        End If
'        txtFind.PasswordChar = IIf(mintFindType = 0 And gblnCardHide, "*", "")
'    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 99
'        '���ҷ�ʽ
'        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
'    Case conMenu_View_ReadIC '��IC��
'        Control.Visible = mintFindType = 5
    Case conMenu_Manage_ThingAdd    '�ӵ�
        Control.Enabled = InStr(mstrPrivs, ";" & "ҽ���ӵ�" & ";") <> 0 'And InStr("3-�˺�", TCurRowState.curRow״̬) <= 0 And TCurRowState.curRow״̬ <> ""
    Case conMenu_Edit_Transf_Liquid
        '��Һ
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr("1-����Һ,0-δ�ӵ�", TCurRowState.curRow״̬) > 0
    Case conMenu_Manage_Call, conMenu_Edit_Transf_Puncture
        '���У�����
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And (TCurRowState.curRow״̬ = "5-������" Or Val(TCurRowState.curRow״̬) = 7) And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
        
'    Case conMenu_Manage_Call
'        '����
'        Control.Enabled = TQueueRowState.curRow�Һŵ� <> "" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
'    Case conMenu_Manage_CallNext
'        '��һ��
'        Control.Enabled = TQueueRowState.nextRow�Һŵ� <> "" And TQueueRowState.nextRow״̬ = "1-����Һ" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
'    Case conMenu_Manage_CallPrevious
'        '��һ��
'        Control.Enabled = TQueueRowState.PrivRow�Һŵ� <> "" And TQueueRowState.PrivRow״̬ = "1-����Һ" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_Reset
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_Up
        '����
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And TCurRowState.PrivRow״̬ = "5-������" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_Down
        '����
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And TCurRowState.nextRow״̬ = "5-������" And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_Discard
        '����
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow״̬) & ",") <= 0 And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
        If Control.Enabled Then
            lng��ˮ�� = Get��ˮ��()
            If lng��ˮ�� > 0 Then
                If Not mobjRecord.Item(CStr(lng��ˮ��)) Is Nothing Then
                    For intItem = 1 To mobjRecord.Item(CStr(lng��ˮ��)).Count
                        If mobjRecord.Item(CStr(lng��ˮ��)).Item(intItem).ִ��״̬ = 1 Then
                            Control.Enabled = False
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Case conMenu_Edit_Bed_Modify
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr(mstrPrivs, ";����״̬;") > 0
    Case conMenu_Manage_Recall
        '�ٻ�
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr("2-����,4-����", TCurRowState.curRow״̬) > 0 And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_TagEnd
        '��Ϊ����
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow״̬) & ",") <= 0 And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
    Case conMenu_Manage_Untread
        '�˺�
        Control.Enabled = TCurRowState.curRow�Һŵ� <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow״̬) & ",") <= 0 And InStr(mstrPrivs, ";" & "�Ŷӹ���" & ";") <> 0
'        TCurRowState.
        If Control.Enabled Then
            lng��ˮ�� = Get��ˮ��()
            If lng��ˮ�� > 0 Then
                If Not mobjRecord.Item(CStr(lng��ˮ��)) Is Nothing Then
                For intItem = 1 To mobjRecord.Item(CStr(lng��ˮ��)).Count
                    If mobjRecord.Item(CStr(lng��ˮ��)).Item(intItem).ִ��״̬ = 1 Then
                        Control.Enabled = False
                        Exit For
                    End If
                Next
                End If
            End If
        End If
    Case conMenu_File_RoomSet   '����̨����
        Control.Enabled = InStr(mstrPrivs, ";��λ����;") > 0
        
'    Case MLNG_INFO + 1 To MLNG_INFO + 99     'ָ����ȡ������Ϣ
'        Control.Checked = (Val(lblBill.Tag) = Control.ID - MLNG_INFO)
    Case Else
        If Me.tbcSub.Selected.Tag = "��λ����" Then
            Call mclsSeating.zlUpdateCommandBars(Control)
        End If
        If Me.tbcSub.Selected.Tag = "ִ����Ŀ" Then
            Call mfrmRecord.zlUpdateCommandBars(Control)
        End If
        If Me.tbcSub.Selected.Tag = "ҩƷ�Ĵ�" Then
            Call mfrmLeaveMedi.zlUpdateCommandBars(Control)
        End If

    End Select
End Sub

Private Sub chkInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjPopupInfo Is Nothing And Button = vbRightButton Then mobjPopupInfo.ShowPopup
End Sub

Private Sub cmdOk_Click()
    '����ʱ��
    Dim datBegin As Date, datEnd As Date
    
'    If DateDiff("d", dtpBegin.Value, dtpEnd.Value) > 7 Then
'        If MsgBox("ָ�������ڼ������7�죬���ܻ�Ӱ���ѯ�ٶȣ��Ƿ������", vbOKCancel + vbDefaultButton2, Me.Caption) = vbCancel Then Exit Sub
'    End If
    
    mdateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    If Trim(txtInfo.Text) = "" Then
        Select Case cboDate.ListIndex
            Case 1              '����
                mDateBegin = Format(mdateEnd - 1, "yyyy-MM-dd 00:00:00")
            Case 2              '������
                mDateBegin = Format(mdateEnd - 2, "yyyy-MM-dd 00:00:00")
            Case 3              '��һ��
                mDateBegin = Format(mdateEnd - 6, "yyyy-MM-dd 00:00:00")
            Case 4              '��ʮ��
                mDateBegin = Format(mdateEnd - 9, "yyyy-MM-dd 00:00:00")
            Case Else           '����
                mDateBegin = Format(mdateEnd, "yyyy-MM-dd 00:00:00")
        End Select
    Else
        mDateBegin = Format(mdateEnd - 364, "yyyy-MM-dd 00:00:00")
    End If
    
'    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
'        mdateEnd = CDate(0)  '��ʾȡ��ǰʱ��
'    Else
'        mdateEnd = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
'    End If
    
    If Not Me.idkSelect.GetCurCard Is Nothing Then
        If Me.idkSelect.GetCurCard.���� = "�Һŵ�" Then
            '�Һŵ��Զ���������
            txtInfo.Text = zlCommFun.GetFullNO(txtInfo.Text, 12)
        End If
    End If

    'ˢ��
    mlngPreDept = 0
    Call cboDept_Click

End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Title Like "�����б�*" Then
        Item.Handle = picLeft.hwnd
    ElseIf Item.Title = "�ӵ�����" Then
        Item.Handle = picRecord.hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picReadyReceive.hwnd
    End If
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub

Private Function ShowPar() As String
    Dim strPar As String, strType As String, i As Integer
    strPar = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
    Next
    strType = Mid(strType, 2)
    strType = Replace(strType, "0", "����")
    strType = Replace(strType, "1", "��Һ")
    strType = Replace(strType, "2", "ע��")
    strType = Replace(strType, "3", "Ƥ��")
    
    ShowPar = "�����б�(" & strType & ")"
End Function

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim arrVal As Variant
    Dim i As Integer
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)

    '����/��ʼ��һ��ͨ����
'    mintCards = 0
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle) Then
        Set mobjSquareCard = Nothing
        MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
    Else
        mstrSquareCards = mobjSquareCard.zlGetIDKindStr(mstrSquareCards)
'        If mstrSquareCards <> "" Then
'            arrVal = Split(mstrSquareCards, ";")
'            mintCards = UBound(arrVal) + 1
'        End If
    End If

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call initMenus
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 250, 400, DockLeftOf, Nothing)
    objPane.Title = ShowPar
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 200
    objPane.MaxTrackSize.Width = 500

    Set objPane = Me.dkpMain.CreatePane(2, 250, 400, DockBottomOf, dkpMain.FindPane(1))
    objPane.Title = "�ӵ�����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(3, 250, 60, DockBottomOf, dkpMain.FindPane(2))
    objPane.Title = "���ӵ�����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.MaxTrackSize.Height = 60
    objPane.MinTrackSize.Height = 60

    picLeft.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)

    'TabControl
    '-----------------------------------------------------
    Set mclsSeating = New clsDockSeating
    Set mfrmRecord = New frmRecord
    Set mfrmLeaveMedi = New frmLeaveMedi

    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsSeating.zlGetForm, "_��λ����"
    mcolSubForm.Add mfrmRecord, "_ִ����Ŀ"
    mcolSubForm.Add mfrmLeaveMedi, "_ҩƷ�Ĵ�"

    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��

        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��Һע��", "")
        
        '�ָ��ϴ�ѡ��Ĳ�����Ϣ
        mstrPatiKey = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ȡ������Ϣ", "")
        
        '���ӵ�������Ϣ
        mintPatiIdentify = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "���ӵ�������Ϣ", "1"))
        If Val(mintPatiIdentify) <= 0 Then
            mintPatiIdentify = 1
        End If
        
        ''If InStr(mstrPrivs, ";" & "��λ����" & ";") <> 0 Or InStr(mstrPrivs, "��λ����" & ";") <> 0 Then
        '.InsertItem(intIdx, "��λ����", mcolSubForm("_��λ����").hwnd, 0).Tag = "��λ����": intIdx = intIdx + 1
        .InsertItem(intIdx, "��λ����", picTmp.hwnd, 0).Tag = "��λ����": intIdx = intIdx + 1
        ''End If

        If InStr(mstrPrivs, ";" & "ҽ��ִ��" & ";") <> 0 Then
            '.InsertItem(intIdx, "ִ����Ŀ", mcolSubForm("_ִ����Ŀ").hwnd, 0).Tag = "ִ����Ŀ": intIdx = intIdx + 1
            .InsertItem(intIdx, "ִ����Ŀ", picTmp.hwnd, 0).Tag = "ִ����Ŀ": intIdx = intIdx + 1
        End If

        If InStr(mstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0 Then
            '.InsertItem(intIdx, "ҩƷ�Ĵ�", mcolSubForm("_ҩƷ�Ĵ�").hwnd, 0).Tag = "ҩƷ�Ĵ�": intIdx = intIdx + 1
            .InsertItem(intIdx, "ҩƷ�Ĵ�", picTmp.hwnd, 0).Tag = "ҩƷ�Ĵ�": intIdx = intIdx + 1
        End If

        If .ItemCount = 0 Then
            MsgBox "��û��ʹ����Һע������Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If

        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = strTab
            .Item(intIdx).Selected = True
            If intIdx = 0 Then tbcSub_SelectedChanged .Item(0)
        Else
            .Item(0).Selected = True
            tbcSub_SelectedChanged .Item(0)
        End If
        Call SubWinDefCommandBar(.Selected) '��ʼˢ�¶���һ�β˵�����ť
    End With

    '2012-08-23 �����б��ҳ
    Call TabListInit
    
    'ҽ�����ҳ�ʼ��
    '----------------------------------------------------
    If patiList.DeptToCbo(cboDept, mstrPrivs) = False Then
        MsgBox "��ʼ��ҽ������ʧ��,����ʹ�ñ�ϵͳ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If

    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "���п���") > 0 Then
            MsgBox "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Else
            MsgBox "û�з�������������,����ʹ�ñ�ϵͳ��", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If

    '
    '����ʱ��
    '----------------------------
    cboDate.ListIndex = 0
    
    
    '��������IDKindNew�ؼ�
    idkSelect.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE, txtInfo
    idkSelect.IDKind = 1
    For i = 1 To idkSelect.ListCount
        If idkSelect.Cards(i).���� = mstrPatiKey Then
            idkSelect.IDKind = i
            Exit For
        End If
    Next
    
    ptiReadyReceive.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE
    ptiReadyReceive.IDKindIDX = mintPatiIdentify
    
    '�����б��ʼ
    '--------------------
    Call InitReport
    Call cmdOk_Click
    
    '����ָ�:�������ִ��
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    '��ʼ�������˿�
    'TransUdpSock.SockSend
    
'    Set mobjIDCard = New clsIDCard
'    On Error Resume Next
'    Set mobjICCard = CreateObject("zlICCard.clsICCard")
'    On Error GoTo 0
    
    Call SetTimer '�����Զ�ˢ��
    
    
    mblnƤ����֤ = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, 1264)) <> 0
    
    Call QueueInit
    Call mdlQueueManage.QueueInit
    If Val(zlDatabase.GetPara("�ƶ�����", glngSys, 1264)) = 1 Then
        Set mfrmTimeCall = mdlQueueManage.QueueTimeCall
        If Not mfrmTimeCall Is Nothing Then
            mfrmTimeCall.Show , Me
        End If
    End If
    
    '���35��ǰ�Ĳ�����־
    Dim strSQL As String
    strSQL = "zl_������Һ������־_ClearOld"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
End Sub

Private Sub TabListInit()
    Dim intIdx As Integer, strTab As String
    With Me.tbcList
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
                
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        
        rptQueue0.Tag = "0"
        .InsertItem(intIdx, "δ�ӵ�", picQueue0.hwnd, 0).Tag = "δ�ӵ�": .Item(intIdx).Visible = False:   intIdx = intIdx + 1   '0-δ�Ŷ�
        rptQueue1.Tag = "1"
        .InsertItem(intIdx, "����Һ", picQueue1.hwnd, 0).Tag = "����Һ":  intIdx = intIdx + 1  '1 -�ѽӵ�/����Һ���ӵ�����ݲ�������׼��/��Һ���̡������Ƿ���д��
        rptQueue5.Tag = "5"
        .InsertItem(intIdx, "������", picQueue5.hwnd, 0).Tag = "������": intIdx = intIdx + 1  '5 -������,��Ҫ���У���Һ��ע�䣩
        rptQueue6.Tag = "6"
        .InsertItem(intIdx, "��ִ��", picQueue6.hwnd, 0).Tag = "��ִ��": .Item(intIdx).Visible = False:  intIdx = intIdx + 1 '6 -��ִ��,������У�Ƥ�ԣ����ƣ�
        rptQueue7.Tag = "7"
        .InsertItem(intIdx, "ִ����", picQueue7.hwnd, 0).Tag = "ִ����": intIdx = intIdx + 1  '7 -ִ����,
        rptPati.Tag = "2,3,4"
        .InsertItem(intIdx, "�ѽ���", picQueueAll.hwnd, 0).Tag = "�ѽ���": intIdx = intIdx + 1 ' 2 -����,3���˺�,4�������
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����б�", "")
        
        For intIdx = 0 To .ItemCount - 1
            If tbcList(intIdx).Visible And tbcList(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= .ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '���⼤���¼�
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            If mblnLiquid Then
                .Item(1).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
            Else
                .Item(4).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
            End If
        End If
        
    End With
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Height <= 10000 Then Height = 10000
    If Width <= 10000 Then Width = 10000
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMode As String, strIndex As String

    strMode = idkSelect.GetCurCard.����
    strIndex = ptiReadyReceive.IDKindIDX
    
    Set patiList = Nothing
    mstrPrivs = ""
    mlngModul = 0
    mstr�Һŵ� = ""
    mlngPreDept = 0
    mintRow = 0
    mDateBegin = CDate(0)
    mdateEnd = CDate(0)
    
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name _
              , "��ȡ������Ϣ", strMode)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name _
              , "���ӵ�������Ϣ", strIndex)

    On Error Resume Next
    
'    mstrIDCard = ""
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled False
'        Set mobjIDCard = Nothing
'    End If
'    Set mobjICCard = Nothing
    
    Unload mfrmLeaveMedi
    Unload mfrmRecord
    
    Set mclsSeating = Nothing
    Set mobjSquareCard = Nothing
    Set mfrmTimeCall = Nothing
    
    Call QueueUnload
    
End Sub

Private Sub lblBill_Click()
    If Not mobjPopupInfo Is Nothing Then mobjPopupInfo.ShowPopup
End Sub

Private Sub idkSelect_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtInfo.Enabled And txtInfo.Visible Then
        txtInfo.Text = ""
        txtInfo.SetFocus
    End If
End Sub

Private Sub idkSelect_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtInfo.Text = objPatiInfor.����
    mblnReadCard = True
    Call txtInfo_KeyPress(0)
End Sub

Private Sub mclsSeating_RequestRefresh()
    '���ܣ���λ�Ӵ���Ҫ��ˢ��
    mlngPreDept = -1
    Call cboDept_Click
End Sub

Private Sub mclsSeating_StatusTextUpdate(ByVal Text As String)
    '��ǰѡ�е���λ��
    If InStr(Text, "_") > 0 Then
        If patiList.mSeatings.Item(Text).����ID = 0 Then
            mstr��λ = Text
        Else
            Dim objPati As cPatient
            For Each objPati In patiList
                If objPati.��λ�� = Mid(Text, InStr(Text, "_") + 1) And objPati.����ID = patiList.mSeatings.Item(Text).����ID Then
                    mstr�Һŵ� = objPati.Key
                    Call ShowLblInfo(mstr�Һŵ�)
                End If
            Next
        End If
    End If
End Sub

'Private Sub optDate_Click(Index As Integer)
'
'    Dim curDate As Date
'    curDate = zldatabase.Currentdate
'    dtpEnd.MaxDate = Format(curDate, "yyyy-MM-dd 23:59:59"): dtpBegin.MaxDate = curDate
'
'    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59:59")
'    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
'
'    Select Case Index
'    Case 1 '����
'        curDate = curDate - 1
'    Case 2 '�������
'        curDate = curDate - 2
'    Case 3 '���һ��
'        curDate = curDate - 6
'    End Select
'    dtpBegin.Value = CDate(Format(curDate, "yyyy-MM-dd 00:00:00"))
'    optDate(Index).Value = True
'
'End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Left = 0
    fraInfo.Top = -90
    fraInfo.Width = picInfo.ScaleWidth
    fraInfo.Height = picInfo.Height + 90
    
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    
    fraWhere.Top = 0
    fraWhere.Left = 15
    fraWhere.Width = picLeft.ScaleWidth - fraWhere.Left
    
    cboDept.Width = fraWhere.Width - cboDept.Left - 135
    cmdOk.Left = fraWhere.Width - cmdOk.Width - 135
    cboDate.Width = fraWhere.Width - cboDate.Left - cmdOk.Width - 150
    txtInfo.Width = fraWhere.Width - txtInfo.Left - 135
    
    tbcList.Left = picLeft.ScaleLeft
    tbcList.Top = fraWhere.Top + fraWhere.Height
    tbcList.Width = picLeft.ScaleWidth - tbcList.Left
    tbcList.Height = picLeft.ScaleHeight - tbcList.Top
    

End Sub

Private Sub ShowPatiList()
    Dim curDate As Date, objRpt As ReportControl
    Dim datBegin As Date, datEnd As Date
    Dim strOrderCol As String
    Dim arrOrderCol As Variant, arrEle As Variant
    Dim i As Integer
    Dim strInfo As String, strTmp As String, strCard As String
    Dim strReserve As String

    '��ѯʱ���
    curDate = zlDatabase.Currentdate
    If mDateBegin = CDate(0) Then
        mDateBegin = CDate(Format(curDate, "yyyy-mm-dd 00:00:00"))
    End If
    datBegin = mDateBegin

    If mdateEnd = CDate(0) Then
        mdateEnd = CDate(Format(curDate, "yyyy-mm-dd 23:59:59"))
    End If
    datEnd = mdateEnd
    
    strCard = idkSelect.GetCurCard.����
    
    'ָ����ȡ������Ϣ
    If Trim(txtInfo.Text) <> "" Then
        'һ��ͨ�Ŀ���ȡ�������
        strTmp = GetSquareCardInfo(mstrSquareCards, strCard, enuCardProperty.�����ID)
        '׼������
        Select Case strCard
            Case "���￨"
                strInfo = "1"
            Case "�����"
                strInfo = "2"
            Case "�Һŵ�"
                strInfo = "3"
            Case "����"
                strInfo = "4"
            Case "���֤��", "�������֤"
                strInfo = "5"
            Case Else
                strInfo = "6"
        End Select
        strInfo = strInfo & "|" & Trim(txtInfo.Text) & "|" & strTmp
    End If

    '��ʾ�����б�
    Call patiList.FetchPatients(mlngPreDept, datBegin, datEnd, , strInfo, , , mobjSquareCard)
    Call PlugInFunc
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.TransfusionShowPatiList(glngSys, 1264, mlngPreDept, datBegin, datEnd, strReserve)
        Call zlPlugInErrH(Err, "TransfusionShowPatiList")
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0
    End If
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    
    '���������
    If objRpt.SortOrder.Count > 0 Then
        For i = 0 To objRpt.SortOrder.Count - 1
            strOrderCol = strOrderCol & objRpt.SortOrder.Column(i).Index & ";" & IIf(objRpt.SortOrder.Column(i).SortAscending, 1, 0)
            If i < objRpt.SortOrder.Count - 1 Then
                strOrderCol = strOrderCol & "|"
            End If
        Next
    End If
    Call patiList.initObjRpt(objRpt, img16)
    '�ָ�������
    If strOrderCol <> "" Then
        arrOrderCol = Split(strOrderCol, "|")
        objRpt.SortOrder.DeleteAll
        For i = LBound(arrOrderCol) To UBound(arrOrderCol)
            arrEle = Split(arrOrderCol(i), ";")
            objRpt.SortOrder.Add objRpt.Columns(arrEle(0))
            objRpt.SortOrder(i).SortAscending = (arrEle(1) = 1)
        Next
    End If
    Set arrEle = Nothing
    Set arrOrderCol = Nothing
    
    Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
   
    If mintPatirow > 0 And mintPatirow < objRpt.Rows.Count Then
        If Not objRpt.Rows(mintPatirow).GroupRow Then
            Call objRpt.SelectedRows.Add(objRpt.Rows(mintPatirow))
            objRpt.Rows(mintPatirow).Selected = True
            Call RptSelectChanged(objRpt)
        End If
    End If
    '��ʾ��ǰ������
    Call Calling(0)
    
    Call SubWinRefreshData(tbcSub.Selected)

End Sub

Private Sub picQueue0_Resize()
    Call PicQueueResize(picQueue0, rptQueue0)
End Sub

Private Sub picQueue1_Resize()
    Call PicQueueResize(picQueue1, rptQueue1, lblNo1, txtNo1)
End Sub

Private Sub PicQueueResize(objPic As PictureBox, objRpt As ReportControl, Optional objLbl As Label, Optional objTxt As TextBox)
    Dim panTmp As Pane
    
    On Error Resume Next
    
    With objPic
        objRpt.Left = objPic.ScaleLeft
        objRpt.Top = objPic.ScaleTop
        objRpt.Width = objPic.ScaleWidth
        objRpt.Height = objPic.ScaleHeight
        Set panTmp = dkpMain.FindPane(2)    '�õ��ӵ�����Pane
        If Not panTmp Is Nothing Then
            If panTmp.Closed Or panTmp.Hidden Then objRpt.Height = objPic.ScaleHeight - 350
        End If
        Set panTmp = Nothing
    End With
    
    If Not objLbl Is Nothing Then
        With objRpt
            objLbl.Left = .Left + 15
            objLbl.Top = .Top + 15
            objTxt.Left = .Left + 15
            objTxt.Top = objLbl.Top + objLbl.Height + 15
            objTxt.Width = .Width - 30
            
            .Top = objTxt.Top + objTxt.Height + 15
            .Height = .Height - .Top
            
        End With
    End If
End Sub

Private Sub picQueue5_Resize()
    Call PicQueueResize(picQueue5, rptQueue5, lblNo5, txtNo5)
End Sub

Private Sub picQueue6_Resize()
    Call PicQueueResize(picQueue6, rptQueue6)
End Sub

Private Sub picQueue7_Resize()
    Call PicQueueResize(picQueue7, rptQueue7)
End Sub

Private Sub picQueueAll_Resize()
    'ԭ�ȵ��б���Ϊ��ʷ��¼�Ĳ�������
    Call PicQueueResize(picQueueAll, rptPati)
    On Error Resume Next
'    lblDate.Left = picQueueAll.ScaleLeft + 25
'    lblDate.Top = picQueueAll.ScaleTop + 15
'
'    cmdOk.Top = lblDate.Top + lblDate.Height + 25
'    cmdOk.Left = picQueueAll.ScaleWidth - cmdOk.Width - 15
'
'    dtpBegin.Top = lblDate.Top + lblDate.Height + 15
'    dtpBegin.Left = picQueueAll.ScaleLeft + 15
'    dtpBegin.Width = (cmdOk.Left - 30) / 2
'    DtpEnd.Top = dtpBegin.Top
'    DtpEnd.Left = dtpBegin.Left + dtpBegin.Width + 15
'    DtpEnd.Width = dtpBegin.Width
'
'    optDate(0).Left = picQueueAll.ScaleLeft + 15
'    optDate(0).Top = dtpBegin.Top + dtpBegin.Height + 15
'
'    optDate(1).Left = optDate(0).Left + optDate(0).Width
'    optDate(1).Top = optDate(0).Top
'
'    optDate(2).Left = optDate(1).Left + optDate(1).Width
'    optDate(2).Top = optDate(0).Top
'
'    optDate(3).Left = optDate(2).Left + optDate(2).Width
'    optDate(3).Top = optDate(0).Top
'
'    optDate(0).Value = 1
    
'    rptPati.Left = picQueueAll.ScaleLeft
'    rptPati.Top = optDate(0).Top + optDate(0).Height + 30
'    rptPati.Width = picQueueAll.ScaleWidth
'    rptPati.Height = picQueueAll.ScaleHeight - rptPati.Top

End Sub

Private Sub picReadyReceive_Resize()
    On Error Resume Next
    ptiReadyReceive.Width = picReadyReceive.ScaleWidth - ptiReadyReceive.Left * 2
End Sub

Private Sub picRecord_Resize()
    On Error Resume Next
    rptRecord.Top = 0
    rptRecord.Left = 0
    rptRecord.Width = picRecord.ScaleWidth
    rptRecord.Height = picRecord.ScaleHeight - Me.stbThis.Height
End Sub

Private Sub ptiReadyReceive_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    If Not objHisPati Is Nothing Then
        ptiReadyReceive.Tag = CLng(objHisPati.����ID)
    Else
        ptiReadyReceive.Tag = ""
    End If
    blnCancel = True    '¼����Ϣ�󲻸Ļ�Ϊ��������
End Sub

Private Sub ptiReadyReceive_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    ptiReadyReceive.Text = ""
    If ptiReadyReceive.Visible And ptiReadyReceive.Enabled Then ptiReadyReceive.SetFocus
End Sub

Private Function FindReadyReceivePati(ByVal ptiVar As PatiIdentify) As Boolean
'���ܣ������ӵ����˵ĵ���
'������
'  ptiVar��PatiIdentify�ؼ�
'���أ�False�ѽӹ������޵��ɽӣ�True�е�δ�ӹ�

    Dim strSQL As String, strPati As String, strPar As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim lngPatiId As Long
    
    '������Ϣ
    Select Case ptiReadyReceive.GetCurCard.����
    Case "�Һŵ�"
        strPati = " And a.�Һŵ� = [5] "
    Case "�����"
        strPati = " And c.����� = [5] "
    Case "����"
        strPati = " And c.���� = [5] "
    Case "���֤��", "�������֤"
        strPati = " And c.���֤�� = [5] "
    Case Else
        If Val(ptiVar.Tag) > 0 Then
            strPati = " And c.����id = [6] "
            lngPatiId = Val(ptiVar.Tag)
        Else
            strPati = " And c.IC���� = [5] "
        End If
    End Select
    
    '��������
    strTemp = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strPar = strPar & IIf(Val(Split(strTemp, ",")(i)) = 1, "," & i, "")
    Next
    
    On Error GoTo hErr
    
    strSQL = "Select a.ҽ��id, a.���ͺ�, Max(a.��������) ��������, Sum(Nvl(b.��������, 0)) �ѽ����� " & vbNewLine & _
             "From (Select a.����id, b.ҽ��id, b.���ͺ�, b.�������� " & vbNewLine & _
             "  From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���˹Һż�¼ D1, ������ҳ D2, ������ĿĿ¼ E, ���ű� F " & vbNewLine & _
             "  Where a.Id = b.ҽ��id And a.����id = c.����id And a.�Һŵ� = D1.No(+) And a.����id = D2.����id(+) And a.��ҳid = D2.��ҳid(+) " & vbNewLine & _
             "    And a.������Ŀid = e.Id And a.ִ�п���id = f.Id And a.������Դ In (1, 2) And Decode(D2.��������(+), -1, 1, D2.��������(+)) = 1 " & vbNewLine & _
             "    And b.ִ�в���id = [1] And b.����ʱ�� Between [2] And [3] And D1.��¼����(+) = 1 And D1.��¼״̬(+) = 1 " & vbNewLine & _
             "    And Instr([4], Nvl(e.ִ�з���, 0)) > 0 " & vbNewLine & _
             strPati & vbNewLine & _
             ") A, ����ҽ��ִ�� B " & vbNewLine & _
             "Where a.ҽ��id = b.ҽ��id(+) And a.���ͺ� = b.���ͺ�(+) " & vbNewLine & _
             "Group By a.ҽ��id, a.���ͺ� " & vbNewLine & _
             "Having Max(a.��������) - Sum(Nvl(b.��������, 0)) > 0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ӵ����˵ĵ���", mlngPreDept, mDateBegin, mdateEnd, strPar, ptiVar.Text, lngPatiId)
    If rsTemp.RecordCount > 0 Then
        '�е�δ�ӹ�
        FindReadyReceivePati = True
    End If
    rsTemp.Close
            
    Exit Function
    
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ptiReadyReceive_KeyPress(KeyAscii As Integer)
    Dim strCard As String
    Dim lngID As Long
    
    strCard = ptiReadyReceive.Cards(ptiReadyReceive.IDKindIDX).����

    If KeyAscii = 13 Then
        '��ȡ����ID
        If Not mobjSquareCard Is Nothing Then
            If ptiReadyReceive.IDKindIDX = 2 Or ptiReadyReceive.IDKindIDX >= 6 Then
                Call mobjSquareCard.zlGetPatiID(ptiReadyReceive.IDKindIDX - 1, ptiReadyReceive.Text, , lngID)
                ptiReadyReceive.Tag = CLng(lngID)
            End If
        End If
        
        '�Һŵ��Զ����뵥��
        If strCard = "�Һŵ�" Then
            ptiReadyReceive.Text = zlCommFun.GetFullNO(ptiReadyReceive.Text, 12)
        End If
    
        Call zlControl.TxtSelAll(ptiReadyReceive)
        
        '�����д�Ĳ��˴��ӵ�����
        If FindReadyReceivePati(ptiReadyReceive) Then
            '�����ݾ͵��ýӵ�����
            Call thingAdd(Val("1-�Զ��ӵ�"))
        Else
            '������
            MsgBox "δ�ҵ����ӵ����˵ĵ��ݣ�", vbInformation, gstrSysName
            ptiReadyReceive.SetFocus
        End If

    Else
        Select Case strCard
            Case "�����"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "�Һŵ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (ptiReadyReceive.Text = "" Or ptiReadyReceive.SelLength = Len(ptiReadyReceive.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "���֤��", "�������֤"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptPati, Button, Shift, X, Y)
End Sub

Private Sub rptMouseMove(objRpt As ReportControl, Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And (Abs(X) > 220 Or Abs(Y) > 220) Then
        If objRpt.SelectedRows.Count > 0 Then
            If Not objRpt.SelectedRows(0).GroupRow Then
                If objRpt.SelectedRows(0).Record(col_ͼ��).Value = "" Then
                    Set objRpt.DragIcon = img16.ListImages("δִ��").Picture
                    objRpt.Drag vbBeginDrag
                End If
            End If
        End If
    End If
End Sub

Private Sub rptMouseUp(objRpt As ReportControl, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        If objRpt.Records.Count <= 0 Then Exit Sub
        If Not objRpt.SelectedRows(0).GroupRow Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        Else
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ViewPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub
Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptPati, Button, Shift, X, Y)
End Sub


Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not Row.GroupRow Then
        If Row.Record(col_�Ŷ�״̬).Value = "1-����Һ" Then
            Call Calling(0)
        End If
    End If
End Sub

Public Sub rptPati_SelectionChanged()
    Call RptSelectChanged(rptPati)
End Sub

Private Sub RptSelectChanged(objRpt)

    Dim i As Integer
    
    If objRpt.SelectedRows.Count = 0 Then
        If objRpt.Rows.Count > 1 Then
            '�м�¼,ȡ�ڸ��Ƿ�����,����ǰ��
            For i = 1 To objRpt.Rows.Count - 1
                If Not objRpt.Rows(i).GroupRow Then
                    objRpt.Rows(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
    
'    Call ShowLblInfo("")
'    Call ShowReport

    If objRpt.SelectedRows.Count = 0 Then Exit Sub  '���������
    mintPatirow = objRpt.SelectedRows(0).Index
    
    With objRpt.SelectedRows(0)
        mfrmRecord.��ˮ�� = Get��ˮ��
        mfrmRecord.�༭ = 0
        mfrmRecord.�޸Ĺ� = False
        mfrmRecord.��Key = ""

        If Not .GroupRow Then
            mstr�Һŵ� = .Record(col_key).Value
            Call ShowLblInfo(mstr�Һŵ�)
            Call SubWinRefreshData(tbcSub.Selected)
        Else
            '������չ�����۵���ı�����,����ȡһ��calling��������
            Call Calling(0)
            mstr�Һŵ� = ""
            Call ShowLblInfo(mstr�Һŵ�)
            Call SubWinRefreshData(tbcSub.Selected)
        End If

    End With
End Sub
Private Sub initMenus()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim i As Integer
    Dim strTmp As String, strDefName As String
    Dim arrCard As Variant

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_RoomSet, "����̨����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Queue_Setup, "��������(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "������־(&L)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "ִ��(&E)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
    
        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Call, "�к�(&J)")
        
'        With objPopup.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "��һλ(&N)", -1, False)
'            Set objControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "��һλ(&P)", -1, False)
'        End With

        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Reset, "����˳��(&R)"): objControl.BeginGroup = True
        'With objPopup.CommandBar.Controls
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Up, "����(&U)", -1, False)
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Down, "����(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Discard, "����(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Recall, "�ٻ�(&R)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Untread, "�˺�(&U)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_TagEnd, "����(&E)", -1, False): objControl.BeginGroup = True
        'End With
        
         
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "�ӵ�(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Liquid, "��Һ")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Call, "����(&J)", -1, False)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Puncture, "����")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_Modify, "����״̬")
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "���ҷ�ʽ(&Y)"): objPopup.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "���˹���(&O)"): objControl.BeginGroup = True
'
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����

    End With

'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
'    objMenu.ID = conMenu_ToolPopup
'    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)")
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
'            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
'        End With
'    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With

'���˾�ȷ��ȡ���ˣ�����ȡ��
    '���������⴦��
    '-----------------------------------------------------
'   ���˵��Ҳ�Ĳ��� �����￨�Ų��ң�֧��ˢ��
'    With cbsMain.ActiveMenuBar.Controls
'        Set objPopup = .Add(xtpControlPopup, conMenu_View_FindType, "����")
'        objPopup.ID = conMenu_View_FindType
'        objPopup.Flags = xtpFlagRightAlign
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
'
'        objCustom.Handle = txtFind.hwnd
'
'        objCustom.Flags = xtpFlagRightAlign
'
'        Set objControl = .Add(xtpControlButton, conMenu_View_ReadIC, "����")
'        objControl.Flags = xtpFlagRightAlign
'    End With
    txtFind.Visible = False
    
    
'ȡ��ԭ��ʽѡ������Ϣ������IDKindNew�ؼ�����
'    '����ʽ�˵�
'    strTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name _
'                         , "��ȡ������Ϣ", "1"))
'    If Val(strTmp) = 0 Then
'        lblBill.Tag = "1"
'    Else
'        i = Val(strTmp)
'        If i > UBound(Split(mstrSquareCards, ";")) + 1 + 6 Then
'            lblBill.Tag = "1"
'        Else
'            lblBill.Tag = strTmp
'        End If
'    End If
'
'    Set mobjPopupInfo = cbsMain.Add("ָ����Ϣ", xtpBarPopup)
'    With mobjPopupInfo.Controls
'        .Add xtpControlButton, MLNG_INFO + 1, "���￨(&1)"
'        .Add xtpControlButton, MLNG_INFO + 2, "�����(&2)"
'        .Add xtpControlButton, MLNG_INFO + 3, "���ݺ�(&3)"
'        .Add xtpControlButton, MLNG_INFO + 4, "��  ��(&4)"
'        .Add xtpControlButton, MLNG_INFO + 5, "���֤(&5)"
'        .Add xtpControlButton, MLNG_INFO + 6, "�ɣÿ�(&6)"
'        'һ��ͨ�Ŀ�
'        If mstrSquareCards <> "" Then
'            arrCard = Split(mstrSquareCards, ";")
'            For i = LBound(arrCard) To UBound(arrCard)
'                strTmp = Split(arrCard(i), "|")(enuCardProperty.ȫ��)
'                If Val(lblBill.Tag) = i + 7 Then
'                    strDefName = strTmp
'                End If
'                If InStr(";���￨;�����;���ݺ�;����;���֤;IC��;�ɣÿ�;", ";" & strTmp & ";") = 0 Then
'                    .Add xtpControlButton, MLNG_INFO + 7 + i, strTmp & "(&" & i + 7 & ")"
'                End If
'            Next
'        End If
'    End With
'    Select Case Val(lblBill.Tag)
'        Case 1
'            lblBill.Caption = "���￨"
'        Case 2
'            lblBill.Caption = "�����"
'        Case 3
'            lblBill.Caption = "���ݺ�"
'        Case 4
'            lblBill.Caption = "��  ��"
'        Case 5
'            lblBill.Caption = "���֤"
'        Case 6
'            lblBill.Caption = "�ɣÿ�"
'        Case Else
'            If strDefName = "" Then
'                'Ĭ��Ϊ���￨
'                lblBill.Caption = "���￨"
'            Else
'                lblBill.Caption = strDefName
'            End If
'    End Select
'    lblBill.Caption = lblBill.Caption & "��"

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����

        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "�ӵ�"): objControl.BeginGroup = True: objControl.ToolTipText = "�ӵ�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Liquid, "��Һ")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Call, "����"): objControl.ToolTipText = "���е�ǰ��Ա"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Puncture, "����")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "��һλ"): objControl.BeginGroup = True: objControl.ToolTipText = "������һλ"
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "��һλ"):: objControl.ToolTipText = "������һλ"

        'Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Manage_Reset, "�ź�", objControl.Index + 1)
        'objPopup.ID = conMenu_Manage_Reset: objPopup.BeginGroup = True
        'With objPopup.CommandBar.Controls
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Up, "����"): objControl.BeginGroup = True: objControl.ToolTipText = "�Ŷ�˳������"
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Down, "����"): objControl.ToolTipText = "�Ŷ�˳������"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Discard, "����"): objControl.ToolTipText = "�����Ŷ�����"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_TagEnd, "����"): objControl.ToolTipText = "���Ϊ��������"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Recall, "�ٻ�"): objControl.ToolTipText = "�����Ŷ�����"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Untread, "�˺�"): objControl.ToolTipText = "�˳��Ŷ�����"
        'End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend          'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse   '�۵�������
        .Add 0, vbKeyF12, conMenu_File_Parameter            '��������
        
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add 0, vbKeyF3, conMenu_Manage_Call                '����
        .Add FCONTROL, vbKeyPageUp, conMenu_Manage_Up       '����
        .Add FCONTROL, vbKeyPageDown, conMenu_Manage_Down   '����
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '��ӡ����
        .AddHiddenCommand conMenu_File_Excel            '�����Excel
    End With

    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub


Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ�½��������
'������
'  objItem��tbcTab�ؼ���Item����

    Dim objPati As cPatient
    Dim strOutNurse As String, objNurse As OutNurse, lng����ID As Long
    For Each objNurse In ObjOutNurse
        strOutNurse = strOutNurse & "|" & objNurse.����
    Next
    If Mid(strOutNurse, 1, 1) = "|" Then strOutNurse = Mid(strOutNurse, 2)
    Call cbsMain_Resize
    Select Case objItem.Caption
    Case "��λ����"
        lng����ID = 0
        Set objPati = Nothing
        
        If mstr�Һŵ� <> "" Then
            If Not patiList.Item(mstr�Һŵ�) Is Nothing Then
                If patiList.Item(mstr�Һŵ�).��λ�� = "��" Or patiList.Item(mstr�Һŵ�).��λ�� = "" Then
                    lng����ID = patiList.Item(mstr�Һŵ�).����ID
                    Set objPati = patiList.Item(mstr�Һŵ�)
                End If
            End If
        End If
        Call mclsSeating.zlRefresh(patiList.mSeatings, lng����ID, objPati)
    Case "ִ����Ŀ"
        Set objPati = Nothing
        
        If mstr�Һŵ� <> "" Then
            If Not patiList.Item(mstr�Һŵ�) Is Nothing Then
                lng����ID = patiList.Item(mstr�Һŵ�).����ID
                Set objPati = patiList.Item(mstr�Һŵ�)
            End If
        End If
        Call mfrmRecord.zlRefresh(mobjRecord, objPati)
    Case "ҩƷ�Ĵ�"
        If mstr�Һŵ� <> "" Then
            If Not patiList.Item(mstr�Һŵ�) Is Nothing Then
                mfrmLeaveMedi.dateBeging = mDateBegin
                mfrmLeaveMedi.DateEnd = mdateEnd
                mfrmLeaveMedi.����ID = patiList.Item(mstr�Һŵ�).����ID
                mfrmLeaveMedi.�Һŵ� = mstr�Һŵ�
                mfrmLeaveMedi.���� = lblinfo(1)
                mfrmLeaveMedi.�Ա� = lblinfo(3)
                mfrmLeaveMedi.���� = lblinfo(5)
                mfrmLeaveMedi.����ID = mlngPreDept
                mfrmLeaveMedi.���� = cboDept.List(cboDept.ListIndex)
                Call mfrmLeaveMedi.zlRefresh
            End If
        End If
    End Select
End Sub

'Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
''���ܣ����֤ʶ��ɹ��󼤻�
'    mstrIDCard = strID
'    If mintFindType = 4 Then
'        txtFind.Text = mstrIDCard
'    Else
'        txtFind.Text = "" '�������(Ŀǰ�������������²��ܼ���)��
'    End If
'    Call ExecuteFindPati(False, mstrIDCard)
'End Sub

Private Sub rptQueue0_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue0, Button, Shift, X, Y)
End Sub

Private Sub rptQueue0_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue0, Button, Shift, X, Y)
End Sub

Private Sub rptQueue0_SelectionChanged()
    Call RptSelectChanged(rptQueue0)
End Sub

Private Sub rptQueue1_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue1, Button, Shift, X, Y)
End Sub

Private Sub rptQueue1_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue1, Button, Shift, X, Y)
End Sub

Private Sub rptQueue1_SelectionChanged()
    Call RptSelectChanged(rptQueue1)
End Sub

Private Sub rptQueue5_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue5, Button, Shift, X, Y)
End Sub

Private Sub rptQueue5_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue5, Button, Shift, X, Y)
End Sub

Private Sub rptQueue5_SelectionChanged()
    Call RptSelectChanged(rptQueue5)
End Sub

Private Sub rptQueue6_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue6, Button, Shift, X, Y)
End Sub

Private Sub rptQueue6_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue6, Button, Shift, X, Y)
End Sub

Private Sub rptQueue6_SelectionChanged()
    Call RptSelectChanged(rptQueue6)
End Sub

Private Sub rptQueue7_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue7, Button, Shift, X, Y)
End Sub

Private Sub rptQueue7_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue7, Button, Shift, X, Y)
End Sub

Private Sub rptQueue7_SelectionChanged()
    Call RptSelectChanged(rptQueue7)
End Sub

Private Sub rptRecord_GotFocus()
    mfrmRecord.��ˮ�� = Get��ˮ��
End Sub

Private Sub rptRecord_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    Dim objRpt As ReportControl
    
    If Button = 2 Then
        If tbcList.Selected.Tag = "δ�ӵ�" Then
            Set objRpt = Me.rptQueue0
        ElseIf tbcList.Selected.Tag = "����Һ" Then
            Set objRpt = Me.rptQueue1
        ElseIf tbcList.Selected.Tag = "������" Then
            Set objRpt = Me.rptQueue5
        ElseIf tbcList.Selected.Tag = "��ִ��" Then
            Set objRpt = Me.rptQueue6
        ElseIf tbcList.Selected.Tag = "ִ����" Then
            Set objRpt = Me.rptQueue7
        ElseIf tbcList.Selected.Tag = "�ѽ���" Then
            Set objRpt = Me.rptPati
        End If
    
        Call SubWinRefreshData(tbcSub.Selected)     'ˢ��
    
        If objRpt.SelectedRows.Count <= 0 Then Exit Sub
        If Not objRpt.SelectedRows(0).GroupRow Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        Else
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ViewPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    Else
        Call SubWinRefreshData(tbcSub.Selected)     'ˢ��
    End If
    
End Sub

Public Sub rptRecord_SelectionChanged()
    Dim i As Integer
    
    mfrmRecord.��ˮ�� = 0
    
    If rptRecord.SelectedRows.Count = 0 Then
        If mintRecordRow > 0 And mintRecordRow < rptRecord.Rows.Count Then
            If Not rptRecord.Rows(mintRecordRow).GroupRow Then
                Call rptRecord.SelectedRows.Add(rptRecord.Rows(mintRecordRow))
                rptRecord.Rows(mintRecordRow).Selected = True
            End If
        End If
    End If

    If rptRecord.SelectedRows.Count = 0 Then
        If rptRecord.Rows.Count > 1 Then
            '�м�¼,ȡ�ڸ��Ƿ�����,����ǰ��
            For i = 1 To rptRecord.Rows.Count - 1
                If Not rptRecord.Rows(i).GroupRow Then
                    rptRecord.Rows(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If

    If rptRecord.SelectedRows.Count = 0 Then Exit Sub '����û��ѡ�����,���˳�
    
    mintRecordRow = rptRecord.SelectedRows(0).Index
    If mfrmRecord.�༭ = 0 Then
        '���ģʽ
        mfrmRecord.��ˮ�� = Get��ˮ��
        Call mfrmRecord.ShowVsList(mfrmRecord.��ˮ��)
        Call mfrmRecord.KernalRefresh
    Else
        '�޸�ģʽ
        mfrmRecord.��ˮ�� = Get��ˮ��
        If mfrmRecord.��ˮ�� <> mfrmRecord.�༭ Then
            MsgBox "�뽫��ǰ��¼���޸����֮����������������", vbExclamation, gstrSysName
            mfrmRecord.��ˮ�� = mfrmRecord.�༭
            Exit Sub
        End If
        If Not mfrmRecord.�޸Ĺ� Then
            '�й��޸ģ���ˢ�¡�
            
            Call mfrmRecord.ShowVsList(mfrmRecord.��ˮ��)
            Call mfrmRecord.KernalRefresh
            
        End If
    End If
    mfrmRecord.��Key = ""
End Sub

Private Sub tbcList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim objRpt As ReportControl
    
    If Item.Tag = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����б�", Item.Tag
    mstrQueueTab = Item.Tag
    If Item.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf Item.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf Item.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf Item.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf Item.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf Item.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    If Not objRpt Is Nothing Then
        Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
        'ѡ��һ��
        Call ˢ��(1)
    End If
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    'If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    If Item.Tag = "" Then Exit Sub
    
    '����ѡ��
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��Һע��", Item.Tag
   
    On Error GoTo errHandle
    Screen.MousePointer = vbHourglass
    If picTmp.hwnd = Item.Handle Then
        Dim objItem As TabControlItem
        Dim intIndex As Integer
        intIndex = Item.Index
        Select Case Item.Tag
            Case "��λ����"
                Set objItem = tbcSub.InsertItem(intIndex, "��λ����", mcolSubForm("_��λ����").hwnd, 0)
                objItem.Tag = "��λ����"
            Case "ִ����Ŀ"
                Set objItem = tbcSub.InsertItem(intIndex, "ִ����Ŀ", mcolSubForm("_ִ����Ŀ").hwnd, 0)
                objItem.Tag = "ִ����Ŀ"
            Case "ҩƷ�Ĵ�"
                Set objItem = tbcSub.InsertItem(intIndex, "ҩƷ�Ĵ�", mcolSubForm("_ҩƷ�Ĵ�").hwnd, 0)
                objItem.Tag = "ҩƷ�Ĵ�"
        End Select
        If Not objItem Is Nothing Then
            objItem.Selected = True
            tbcSub.RemoveItem intIndex + 1
        
            Call SubWinDefCommandBar(objItem)
            'ˢ���Ӵ�������
            Call SubWinRefreshData(objItem)
        
        End If
    Else
        Call SubWinDefCommandBar(Item)
        'ˢ���Ӵ�������
        Call SubWinRefreshData(Item)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ParameterSetup()
    Dim strRoom As String

    frmTransfusionSetup.mstrPrivs = mstrPrivs
    frmTransfusionSetup.mlng����ID = cboDept.ItemData(cboDept.ListIndex)
    frmTransfusionSetup.Show vbModal, Me
    If frmTransfusionSetup.mblnOk Then
        'Ƥ����֤���
        mblnƤ����֤ = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, 1264)) <> 0
        
        '�����Զ�ˢ��
        Call SetTimer
        '�����б�
        Me.dkpMain.FindPane(1).Title = ShowPar
        Call cmdOk_Click
    End If
    timRefresh.Enabled = mintRefresh <> 0
    
End Sub

Private Sub ShowLblInfo(ByVal str�Һŵ� As String)
    Dim objPati As cPatient
    Dim dateS As Date, dateE As Date
    On Error GoTo hNoPati
    
    If str�Һŵ� = "" Then
        GoTo hNoPati
    Else
        Set objPati = patiList.Item(str�Һŵ�)
        
        If Not objPati Is Nothing Then
            lblinfo(1) = objPati.����
            lblinfo(3) = objPati.�Ա�
            lblinfo(5) = objPati.����
            lblinfo(7) = objPati.�ѱ�
            lblinfo(11) = objPati.���˿���
            lblinfo(13) = objPati.�������
            lblinfo(15) = objPati.���￨��
                       
            Set mobjRecord = New ExecRecord
            dateS = objPati.�Һ�ʱ��
            dateE = zlDatabase.Currentdate
            Call mobjRecord.GetExecGroups(objPati, mlngPreDept, 1, dateS, dateE)
         
        Else
            GoTo hNoPati
        End If

    End If
    Exit Sub
hNoPati:
    lblinfo(1) = ""
    lblinfo(3) = ""
    lblinfo(5) = ""
    lblinfo(7) = ""
    lblinfo(11) = ""
    lblinfo(13) = ""
    lblinfo(15) = ""
    Set mobjRecord = Nothing '�����Ŀ
        
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long

    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).STYLE
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)

    Me.Caption = "������Һע����� - " & objItem.Caption

    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next

    '���������¼���
    Call initMenus

    '�Ӵ������¼���

    Call dkpMain.FindPane(2).Close
    Select Case objItem.Tag
    Case "ִ����Ŀ"
        Call mfrmRecord.zlDefCommandBars(Me, Me.cbsMain)
        Call dkpMain.ShowPane(2)
    Case "��λ����"
        Call mclsSeating.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "ҩƷ�Ĵ�"
        Call mfrmLeaveMedi.zlDefCommandBars(Me, Me.cbsMain)
    End Select

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.STYLE = bytStyle
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next

    '�������RecalcLayout����������
    Call LockWindowUpdate(0)

    'Set mfrmActive = mcolSubForm("_" & zlCommFun.NVL(objItem.Tag))
End Sub

Private Sub Calling(ByVal intRow As Integer)
    '��ʾ��ǰ����λ��
    Dim intFister As Integer
    Dim objRpt As ReportControl
    Dim blnInitRow As Boolean, i As Integer

    '-- ����ȡ��ǰ������
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    
    blnInitRow = True
    For i = 1 To objRpt.Rows.Count - 1
        If Not objRpt.Rows(i).GroupRow Then
            If objRpt.Rows(i).Record(col_�Ŷ�״̬).Value = "5-������" Then
                If intFister = 0 Then intFister = i
                If objRpt.Rows(i).Record(col_calling).Icon = 5 Then
                    mintRow = i
                    blnInitRow = False
                    Exit For
                End If
            End If
        End If
    Next
    If blnInitRow = True And intFister > 0 Then
        If mintRow <= 0 Or mintRow > objRpt.Rows.Count Then mintRow = intFister
    ElseIf intFister = 0 Then
        mintRow = 0
    End If
'
'    '-- ������ʾ
'    If mintRow + intRow > 0 And mintRow + intRow <= rptPati.Rows.Count Then
'        If Not rptPati.Rows(mintRow + intRow).GroupRow Then
'            If rptPati.Rows(mintRow + intRow).Record(col_�Ŷ�״̬).Value = "1-����Һ" Then
'                rptPati.Rows(mintRow).Record(col_calling).Icon = 6
'                rptPati.Rows(mintRow + intRow).Record(col_calling).Icon = 5
'                mintRow = mintRow + intRow
'                rptPati.Redraw
'            End If
'        End If
'    End If

    Dim objPati As cPatient
    With objRpt
        If .SelectedRows.Count <= 0 Then Exit Sub
        If .SelectedRows(0).GroupRow Then Exit Sub
        If .SelectedRows(0).Record(col_�Ŷ�״̬).Value <> "5-������" Then Exit Sub
 
        If patiList.Item(mstr�Һŵ�).SetCallTag(mlngPreDept) Then
            For Each objPati In patiList
                objPati.���б�־ = 0
            Next
            patiList.Item(mstr�Һŵ�).���б�־ = 1
            'Call patiList.PatiListRefresh(rptPati)
            .Rows(mintRow).Record(col_calling).Icon = 6
            .SelectedRows(0).Record(col_calling).Icon = 5
            .Redraw
        End If
        
    End With

End Sub

Private Function SiblingRowState(objRpt As ReportControl, ByVal intRow As Integer) As SiblingRow
    'ȡ�����е�״̬
    With SiblingRowState
        If intRow + 1 < objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow + 1).GroupRow Then
                .nextRow�Һŵ� = objRpt.Rows(intRow + 1).Record(col_key).Value
                .nextRow״̬ = objRpt.Rows(intRow + 1).Record(col_�Ŷ�״̬).Value
                .nextRowIndex = intRow + 1
            End If
        End If

        If intRow - 1 >= 0 And intRow <= objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow - 1).GroupRow Then
                .PrivRow�Һŵ� = objRpt.Rows(intRow - 1).Record(col_key).Value
                .PrivRow״̬ = objRpt.Rows(intRow - 1).Record(col_�Ŷ�״̬).Value
                .PrivRowIndex = intRow - 1
            End If
        End If

        If intRow >= 0 And intRow < objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow).GroupRow Then
                .curRow�Һŵ� = objRpt.Rows(intRow).Record(col_key).Value
                .curRow״̬ = objRpt.Rows(intRow).Record(col_�Ŷ�״̬).Value
                .curRowIndex = intRow
            End If
        End If
    End With
End Function

Private Sub rptQueueMove(ByVal intRow As Integer)
    '�ƶ�λ��
    Dim icurRow As Integer, lngTmp As Long
    Dim TcurrowStat As SiblingRow, TobjRowStat As SiblingRow
    Dim objRpt As ReportControl
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    If objRpt.SelectedRows.Count > 0 Then
        icurRow = objRpt.SelectedRows(0).Index

        TcurrowStat = SiblingRowState(objRpt, icurRow)  'ȡ������״̬
        TobjRowStat = SiblingRowState(objRpt, icurRow + intRow)

        If (icurRow + intRow > 0) And ((icurRow + intRow) < objRpt.Rows.Count) Then

            lngTmp = Val(patiList.Item(TcurrowStat.curRow�Һŵ�).��Ȩ��)
            patiList.Item(TcurrowStat.curRow�Һŵ�).��Ȩ�� = Val(patiList.Item(TobjRowStat.curRow�Һŵ�).��Ȩ��)
            patiList.Item(TobjRowStat.curRow�Һŵ�).��Ȩ�� = lngTmp

            Call patiList.Item(TcurrowStat.curRow�Һŵ�).UpdateSequence(mlngPreDept)
            Call patiList.Item(TobjRowStat.curRow�Һŵ�).UpdateSequence(mlngPreDept)

            Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
            objRpt.Rows(TobjRowStat.curRowIndex).Selected = True
            
            If mintRow = icurRow Then
                If (mintRow + intRow) > 0 And ((mintRow + intRow) < objRpt.Rows.Count) Then
                    mintRow = mintRow + intRow
                End If
            Else
                If icurRow + intRow = mintRow Then mintRow = mintRow - intRow
            End If
            'Call Calling(0)
        End If
    End If
End Sub

Public Function UpdateState(ByVal strState) As Boolean
    '�޸��Ŷӵ�״̬
    Dim objRpt As ReportControl
    
    UpdateState = False
    If InStr("2-����,1-����Һ,3-�˺�,4-����,5-������,6-��ִ��,7-ִ����", strState) > 0 Then

        If mstr�Һŵ� <> "" Then
            If tbcList.Selected.Tag = "δ�ӵ�" Then
                Set objRpt = Me.rptQueue0
            ElseIf tbcList.Selected.Tag = "����Һ" Then
                Set objRpt = Me.rptQueue1
            ElseIf tbcList.Selected.Tag = "������" Then
                Set objRpt = Me.rptQueue5
            ElseIf tbcList.Selected.Tag = "��ִ��" Then
                Set objRpt = Me.rptQueue6
            ElseIf tbcList.Selected.Tag = "ִ����" Then
                Set objRpt = Me.rptQueue7
            ElseIf tbcList.Selected.Tag = "�ѽ���" Then
                Set objRpt = Me.rptPati
            End If
                    
            If patiList.Item(mstr�Һŵ�).UpdateState(strState, mlngPreDept) Then
                UpdateState = True
                Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
                'Call Calling(0)
                If Val(strState) >= 2 And Val(strState) <= 4 Then ClearSeat
            End If
        End If
    End If
End Function

'Private Sub TransUdpSock_inSockString()
'    '����ģ���յ�����Ϣ
'    If TransUdpSock.Infos(TransUdpSock.Infos.Count).����ģ�� = con_����ģ�� And _
'       TransUdpSock.����IP = TransUdpSock.Infos(TransUdpSock.Infos.Count).����IP Then
'        '�������͵�
'        Call MsgBox("�յ���ģ�鷢�͵���Ϣ:" & TransUdpSock.Infos)
'
'    End If
'End Sub


Private Sub ClearSeat()
    
    'ִ�������λ����
    If patiList.Item(mstr�Һŵ�).��λ�� <> "" Then
        Dim objSeat As Seating
        For Each objSeat In patiList.mSeatings
            If objSeat.����ID = patiList.Item(mstr�Һŵ�).����ID And patiList.Item(mstr�Һŵ�).��λ�� = objSeat.��� Then
                Call patiList.mSeatings.Clear(objSeat.��� & "_" & objSeat.���)
                'ˢ����λ
                Call mclsSeating.zlRefresh(patiList.mSeatings, patiList.Item(mstr�Һŵ�).����ID, patiList.Item(mstr�Һŵ�))
            End If
        Next
    End If
    
End Sub
Private Sub InitReport()
    '����LOadʱ����һ��
    Dim objCol As ReportColumn
    With rptRecord
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������

        Set objCol = .Columns.Add(rptCOL.rptCOL_ִ�з���, "ִ�з���", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(rptCOL.rptCOL_�ӵ�ʱ��, "�ӵ�ʱ��", 80, True)
        Set objCol = .Columns.Add(rptCOL.rptCOL_��ҩ��, "��ҩ��", 60, True)
        Set objCol = .Columns.Add(rptCOL.rptCOL_�ӵ���, "�ӵ���", 60, True)

        '����������
        Set objCol = .Columns.Add(rptCOL_��ʱ, "��ʱ", 0, False)
        Set objCol = .Columns.Add(rptCOL_��ϵ��, "��ϵ��", 0, False)
        Set objCol = .Columns.Add(rptCOL_����, "����", 0, False)
        Set objCol = .Columns.Add(rptCOL_��ˮ��, "��ˮ��", 0, False)


        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = rptCOL_ִ�з���
            If objCol.Width = 0 Then objCol.Visible = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."

        End With

        .PreviewMode = True

        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList img16

        .GroupsOrder.Add .Columns(rptCOL_ִ�з���)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����

        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(rptCOL_�ӵ�ʱ��)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(rptCOL_��ˮ��)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Public Sub ShowReport()
    Dim i As Integer
    '�������
    rptRecord.Records.DeleteAll

    If mobjRecord Is Nothing Then
        rptRecord.Populate '������ʾ
        Exit Sub
    End If
    '��ʾ����
    With mobjRecord
        For i = 1 To mobjRecord.Count
            Call AddRecord(mobjRecord.Item(i))
        Next
    End With
    rptRecord.Populate
    Call rptRecord_SelectionChanged
End Sub

Public Function Get��ˮ��() As Long
    'ȡ��ǰѡ���е���ˮ��
    Get��ˮ�� = 0
    On Error GoTo errHandle
    With rptRecord
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then
                If .Columns.Count > rptCOL_��ˮ�� Then
                    Get��ˮ�� = .SelectedRows(0).Record(rptCOL_��ˮ��).Value
                End If
            End If
        End If
    End With
    Exit Function
errHandle:
    Get��ˮ�� = 0
    If Err.Number = 5 Then
        Exit Function
    Else
        If ErrCenter = 1 Then
            Resume
        End If
    End If
End Function

Private Sub AddRecord(ByVal objExecRecord As ExecutiveGroup)
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim intIcon As Integer
    With objExecRecord
        Set objRecord = rptRecord.Records.Add
        Call Add_rptItem(objRecord, .ִ�з���)
        Call Add_rptItem(objRecord, Format(.ִ��ʱ��, "MM-dd hh:mm"))
        Call Add_rptItem(objRecord, .��ҩ��)
        Call Add_rptItem(objRecord, .�ӵ���)

        Call Add_rptItem(objRecord, IIf(.�ܺ�ʱ = 0, "", .�ܺ�ʱ))
        Call Add_rptItem(objRecord, IIf(.��ϵ�� = 0, "", .��ϵ��))
        Call Add_rptItem(objRecord, .����)
       
        
        Call Add_rptItem(objRecord, .��ˮ��)
        Select Case Val(Mid(.ִ�з���, 1, 1))

        Case 1
            '��Һ
            objRecord.PreviewText = "����:" & .���� & _
                                    IIf(.��ϵ�� = 0, "", " ��ϵ��:" & .��ϵ��) & _
                                    IIf(.�ܺ�ʱ = 0, "", " ��ʱ:" & .�ܺ�ʱ)
        Case 2
            'ע��
            objRecord.PreviewText = "����:" & .����
        Case 3
            'Ƥ��
            objRecord.PreviewText = IIf(.�ܺ�ʱ = 0, "", " ��ʱ:" & .�ܺ�ʱ)
        Case Else
            '����
        End Select
    End With
End Sub

Private Function Add_rptItem(ByRef objRecord As ReportRecord, ByVal strValues As String) As ReportRecordItem
    Set Add_rptItem = objRecord.AddItem(strValues)
    Add_rptItem.Caption = strValues

End Function

Public Sub �����ӵ�(ByVal lng��ˮ��)
    '----------
    '�Ӵ������
    Dim lngRow As Long, lngDeptID As Long, objRpt As ReportControl
    Dim lngErrNo As Long
    Dim objPati As cPatient
    
    lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call mobjRecord.Item(CStr(lng��ˮ��)).Undo(lng��ˮ��, lngDeptID, lngErrNo)
    If lngErrNo <> 0 Then Exit Sub
    
    Set objPati = patiList.Item(mstr�Һŵ�)
    
    SaveOperLog lngDeptID, objPati, MEDICAL, "��ˮ��Ϊ" & lng��ˮ�� & "��ҽ��ִ���˳�������"
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    
    If objRpt.SelectedRows.Count > 0 Then lngRow = objRpt.SelectedRows(0).Index
    
    
    rptRecord.SelectedRows(0).Record.DeleteAll
    Call mobjRecord.Remove(CStr(lng��ˮ��))
    Call ˢ��(lngRow)
    rptRecord.Populate
End Sub

Public Sub ˢ��(Optional ByVal lngRow As Long)
    '�Ӵ������
    Dim objRpt As ReportControl
    mlngPreDept = -1
    mdateEnd = CDate(0)
    Call cmdOk_Click
    If tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "������" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = Me.rptPati
    End If
    
    If lngRow > 0 And lngRow <= objRpt.Rows.Count Then
        objRpt.SetFocus
        Call objRpt.SelectedRows.Add(objRpt.Rows(lngRow))
        objRpt.Rows(lngRow).Selected = True
        'Call RptSelectChanged(objRpt)
    End If
End Sub

Public Sub ����״̬��(ByVal strText As String)
    Me.stbThis.Panels(2).Text = strText
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String)
    '���ܣ�����(��һ��)����
    '������blnNext=�Ƿ������һ��
    '      strIDCard=����ֵʱ����ʾ�̶������֤�Ų���

    
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Dim objRpt As ReportControl, strNO As String
'    '��������ʽ���Һ��Զ�ˢ���֤�ļ���������ȡ��
'    If strIDCard = "" And txtFind.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mintFindType = 2 Then
        txtFind.Text = GetFullNO(txtFind.Text, 12)  '12���Һ��վݺ�
    End If
    Call zlControl.TxtSelAll(txtFind)
          
    
    '------------------------------------------------------------------------
    Dim objPati As cPatient, intFind As Integer, iCount As Integer
    
    If blnReStart Then mintLastFind = 0
    
    If blnNext Then
        intFind = mintLastFind + 1
    Else
        intFind = 1
    End If
    strNO = ""
    
    For Each objPati In patiList
        If strIDCard <> "" Then '���֤�Զ�ʶ��ǿ������
            If UCase(objPati.���֤��) = UCase(strIDCard) Then
                iCount = iCount + 1
                strNO = objPati.�Һŵ�
            End If
        Else
            If mintFindType = 0 Then '���￨
                If UCase(objPati.���￨��) = UCase(txtFind.Text) Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 1 Then '�����
                If UCase(objPati.�����) = UCase(txtFind.Text) Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 2 Then '���ݺ�
                If UCase(objPati.�Һŵ�) = UCase(txtFind.Text) Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 3 Then '����
                If UCase(objPati.����) Like "*" & UCase(txtFind.Text) & "*" Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 4 Then '���֤
                If UCase(objPati.���֤��) = UCase(txtFind.Text) Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 5 Then 'IC��
                If UCase(objPati.IC����) = UCase(txtFind.Text) Then
                    strNO = objPati.�Һŵ�
                    iCount = iCount + 1
                End If
            End If
        End If
        
        If iCount = intFind Then Exit For
    Next
    If iCount > 0 And iCount = intFind Then
        mintLastFind = iCount
    Else
        strNO = ""
    End If
    If strNO = "" Then
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
'    If Val(patiList(strNo).�Ŷ�״̬) = 0 Then
'        Set objRpt = Me.rptQueue0
'        tbcList.Item(0).Selected = True
'    Else
    If Val(patiList(strNO).�Ŷ�״̬) = 1 Then
        Set objRpt = Me.rptQueue1
        tbcList.Item(1).Selected = True
    ElseIf Val(patiList(strNO).�Ŷ�״̬) >= 2 And Val(patiList(strNO).�Ŷ�״̬) <= 4 Then
        Set objRpt = Me.rptPati
        tbcList.Item(5).Selected = True
    ElseIf Val(patiList(strNO).�Ŷ�״̬) = 5 Then
        Set objRpt = Me.rptQueue5
        tbcList.Item(2).Selected = True
    ElseIf Val(patiList(strNO).�Ŷ�״̬) = 6 Then
        Set objRpt = Me.rptQueue6
        tbcList.Item(3).Selected = True
    ElseIf Val(patiList(strNO).�Ŷ�״̬) = 7 Or Val(patiList(strNO).�Ŷ�״̬) = 0 Then
        Set objRpt = Me.rptQueue7
        tbcList.Item(4).Selected = True
    End If
    '------------------------------------------------------------------------
            
    '��ʼ������
    
    i = 0 'ReportControl����������0��ʼ
    '���Ҳ���
    
    For i = i To objRpt.Rows.Count - 1
        With objRpt.Rows(i)
            If Not .GroupRow Then
                If .Record(col_�Һŵ�).Value = strNO Then Exit For
            End If
        End With
    Next

    If i <= objRpt.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set objRpt.FocusedRow = objRpt.Rows(i)
        objRpt.SetFocus
        tmrAutoReady.Enabled = True
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If

End Sub

Private Sub timRefresh_Timer()
    Static lngSecond As Long
    
    If mintRefresh = 0 Then Exit Sub
    lngSecond = lngSecond + 1 '����
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call ˢ��
    End If
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("ҽ��ˢ�¼��", glngSys, 1264))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '�̶�Ϊ1����
        timRefresh.Enabled = True
    End If
End Sub

Private Sub tmrAutoReady_Timer()
    '�ҵ������Զ��ӵ� 2012-05-14
    
    tmrAutoReady.Enabled = False
    If Val(zlDatabase.GetPara("������Һ�Զ��ӵ�", glngSys, 1264)) <> 0 Then
        If cbsMain.FindControl(, conMenu_Manage_ThingAdd).Enabled Then
            Call thingAdd           '�Զ��ӵ�
        End If
    End If
End Sub

'Private Sub txtFind_Change()
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtFind.Text = "" And Me.ActiveControl Is txtFind
'    End If
'End Sub

'Private Sub txtFind_GotFocus()
'    If txtFind.Tag = "" Then
'        Call zlControl.TxtSelAll(txtFind)
'    End If
'    txtFind.Tag = ""
'
'    If Not mobjIDCard Is Nothing Then
'        If txtFind.Text = "" Then mobjIDCard.SetEnabled True
'    End If
'End Sub

'Private Sub txtFind_KeyPress(KeyAscii As Integer)
'    '���س�
'    Dim blnCard As Boolean
'
'    '�Ƿ�ˢ�����
'    blnCard = mintFindType = 0 And KeyAscii <> 8 And Len(txtFind.Text) = gbytCardLen - 1 And txtFind.SelLength <> Len(txtFind.Text)
'    If blnCard Or KeyAscii = 13 Then
'        If KeyAscii <> 13 Then
'            txtFind.Text = txtFind.Text & Chr(KeyAscii)
'            txtFind.SelStart = Len(txtFind.Text)
'        End If
'        KeyAscii = 0
'        Call ExecuteFindPati
'    Else
'        Select Case mintFindType
'            Case 0 '���￨
'                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
'                    KeyAscii = 0
'                Else
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                End If
'            Case 1 '�����
'                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'            Case 2 '�Һŵ�
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                If Not (txtFind.Text = "" Or txtFind.SelLength = Len(txtFind.Text)) _
'                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
'                    KeyAscii = 0
'                End If
'            Case 3 '����
'            Case 4 '���֤
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'            Case 5 'IC��
'        End Select
'    End If
'
'End Sub
'
'Private Sub txtFind_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Public Sub ExecuteTest(ByVal str��ˮ�� As String, strGroupKey As String)
    '��дƤ�Խ��
    Dim strResult As String
    Dim cnNew As ADODB.Connection, strUserName As String, lngDeptID As Long, blnEnd As Boolean, objGroup As Group, i As Integer
    Dim str��� As String
    Dim lngErrNo As Long
    
    'If InStr(",(+),(-),����,", "," & str��� & ",") <= 0 Then Exit Sub
    
    If Val(Me.mobjRecord.Item(str��ˮ��).ִ�з���) = 3 And Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��״̬ <> 1 Then
        'δ��ɵ�Ƥ��
        If Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).Ƥ�Խ�� <> "" Then
            If MsgBox("�ò������ı�ԭ�е�Ƥ�Խ�����Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If mblnƤ����֤ And cnNew Is Nothing Then
            Set cnNew = New ADODB.Connection
            strUserName = zlDatabase.UserIdentify(Me, "��дƤ�Խ��ǰ�������������û�����������������֤��", glngSys, 1263, "ȷ��ִ�����", cnNew)
            If strUserName = "" Then Exit Sub
        End If
        
        lngDeptID = Me.cboDept.ItemData(Me.cboDept.ListIndex)
        Call Me.mobjRecord.Item(str��ˮ��).Update(str��ˮ��, strGroupKey, lngDeptID, lngErrNo)
        
        If lngErrNo <> 0 Then
'            lngErrNo_Out = lngErrNo
            Exit Sub
        End If
        
        Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ���� = IIf(strUserName = "", UserInfo.����, strUserName)
        
        Dim objPati As cPatient
        Dim objexecGroup As ExecutiveGroup
        Set objexecGroup = Me.mobjRecord.Item(str��ˮ��)
        If objexecGroup.ExecuteTestFinish(strGroupKey, Me, mobjSquareCard, str���) Then
            Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).Ƥ�Խ�� = str���
            Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��״̬ = 1
            Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ���� = IIf(strUserName = "", UserInfo.����, strUserName)
            
            Set objPati = patiList.Item(mstr�Һŵ�)
            SaveOperLog mlngPreDept, objPati, MEDICAL, "��ˮ��" & str��ˮ�� & ",ҽ��ID,���ͺ�" & strGroupKey & "�ļ�¼,��дƤ�Խ��Ϊ" & str���
            
            '---- ����ˮ���µ�����ҽ����ִ���꣬������������
            blnEnd = True
            For i = 1 To Me.mobjRecord.Count
                For Each objGroup In Me.mobjRecord.Item(i)
                    If objGroup.ִ��״̬ <> 1 Then blnEnd = False
                Next
            Next
            If blnEnd Then
                If Me.UpdateState("4-����") Then
                    SaveOperLog mlngPreDept, objPati, QUEUE, "��дƤ�Խ�������״̬Ϊ4-����"
                Else
                    SaveOperLog mlngPreDept, objPati, QUEUE, "��дƤ�Խ����δ����״̬"
                End If
            End If
        End If
        
    End If
End Sub

Public Sub ExecComplt(ByVal str��ˮ�� As String, ByVal strGroupKey As String)
    '��Һҽ����ɹ���
    
    Dim intִ��״̬ As Integer, lngDeptID As Long, blnEnd As Boolean, objGroup As Group, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim objPati As cPatient
    
    '0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��

    intִ��״̬ = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��״̬
    lngDeptID = Me.cboDept.ItemData(Me.cboDept.ListIndex)
    
    If intִ��״̬ <> 1 And intִ��״̬ <> 2 And Val(Me.mobjRecord.Item(str��ˮ��).ִ�з���) <> 3 Then
        
        If Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).�������� = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).��ִ������ + Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).�������� Then
            
            '�����ǰ�����һ�Σ����� ǰ���λ���δ��д����Ա�ļ�¼������ɡ�
            If Not CheckComplt(str��ˮ��, strGroupKey) Then Exit Sub
                
            If Me.mobjRecord.Item(str��ˮ��).ExecuteFinish(strGroupKey, lngDeptID, , Me, mobjSquareCard) Then
                Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��״̬ = 1
                
                Set objPati = patiList.Item(mstr�Һŵ�)
                SaveOperLog mlngPreDept, objPati, MEDICAL, "ҽ��ִ����ɲ���,��ˮ��" & str��ˮ�� & ",ҽ��ID,���ͺ�" & strGroupKey
                
                '---- ����ҽ����ִ���꣬������������
                blnEnd = True
                For i = 1 To Me.mobjRecord.Count
                    For Each objGroup In Me.mobjRecord.Item(i)
                        If Not (objGroup.ִ��ҽ��ID = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��ҽ��ID And objGroup.���ͺ� = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).���ͺ� And objGroup.�ϴ�ִ��ʱ�� = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).�ϴ�ִ��ʱ��) Then
                            If objGroup.ִ��״̬ <> 1 Then blnEnd = False
                        End If
                    Next
                Next
                If blnEnd Then
                    If Me.UpdateState("4-����") Then
                        SaveOperLog mlngPreDept, objPati, QUEUE, "ҽ����ɺ����״̬Ϊ4-����"
                        
                        '2012-07-17 ���֮���Զ������λռ�� 51193����
                        strSQL = "select ��� from ��λ״����¼ where ����ID=[1] and ����ID=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ִ�����", patiList.Item(mstr�Һŵ�).����ID, lngDeptID)
                        If Not rsTmp.EOF Then
                            strSQL = "Zl_��λ״����¼_Clear(" & lngDeptID & ",'" & rsTmp!��� & "')"
                            Call zlDatabase.ExecuteProcedure(strSQL, "ִ�����")
                            SaveOperLog mlngPreDept, objPati, SEAT, "ҽ����ɺ����ռ����λ" & rsTmp!���
                        End If
                    Else
                        SaveOperLog mlngPreDept, objPati, QUEUE, "ҽ����ɺ�δ����״̬"
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Public Function ExecStart(ByVal str��ˮ�� As String, ByVal strGroupKey As String, Optional ByVal blnUndo As Boolean = False) As Boolean
'��ʼ/������ʼ���ܣ���Ҫ�������޸�ִ��ʱ�䣬��д/���ִ����
    
    Dim intִ��״̬ As Integer, strOper As String, strDate As String
    Dim objOutNur As OutNurse, strOutNurs As String, strSQL As String
    Dim objPati As cPatient
    
    Set objPati = patiList.Item(mstr�Һŵ�)
    If blnUndo Then
        '������ʼ
        Call Me.mobjRecord.Item(str��ˮ��).ExecStart(2, strGroupKey, Now, "")
        SaveOperLog mlngPreDept, objPati, MEDICAL, "ҽ��������ʼ��������ˮ��" & str��ˮ�� & "��ҽ��ID�ͷ��ͺ�Ϊ" & strGroupKey
        ExecStart = True
    Else
        '��ʼ
        intִ��״̬ = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ��״̬
        ExecStart = False
        If Not (intִ��״̬ >= 1 And intִ��״̬ <= 2) Then
            strOper = Me.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ����
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            
            For Each objOutNur In ObjOutNurse
                strOutNurs = strOutNurs & "|" & objOutNur.����
            Next
            If Mid(strOutNurs, 1, 1) = "|" Then strOutNurs = Mid(strOutNurs, 2)
            
            If frmRecordStart.ShowSelect(strOutNurs, strDate, strOper) Then
                Call Me.mobjRecord.Item(str��ˮ��).ExecStart(1, strGroupKey, CDate(strDate), strOper)
                
                SaveOperLog mlngPreDept, objPati, MEDICAL, "ҽ����ʼ������ִ������Ϊ" & strOper & "����ˮ��Ϊ" & str��ˮ�� & "��ҽ��ID�ͷ��ͺ�Ϊ" & strGroupKey
                
                Set objPati = patiList.Item(mstr�Һŵ�)
                If Not objPati Is Nothing Then
                    'strSQL = "Zl_�ŶӼ�¼_Startend(1," & mlngPreDept & ",'" & mstr�Һŵ� & "',to_date('" & strDate & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
                    strSQL = "Zl_�ŶӼ�¼_Startend(1," & _
                                    mlngPreDept & "," & _
                                    objPati.����ID & "," & _
                                    IIf(objPati.������Դ = 1, "Null", "'" & objPati.�Һŵ� & "'") & "," & _
                                    IIf(objPati.������Դ = 1, objPati.����ID, "Null") & "," & _
                                    IIf(strOper = "", "Null", "'" & strOper & "'") & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    ExecStart = True
                End If
            End If
        End If
    End If
End Function

Private Sub txtInfo_Change()
    idkSelect.SetAutoReadCard Trim(txtInfo.Text) = ""
End Sub

'Private Sub txtInfo_Change()
'    If txtInfo.Enabled = False Then Exit Sub
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtInfo.Text = "" And Me.ActiveControl Is txtInfo
'    End If
'    If txtInfo.Text = "" Then Call cboDate_Click
'End Sub

Private Sub txtInfo_GotFocus()
    Call zlControl.TxtSelAll(txtInfo)
    idkSelect.SetAutoReadCard Trim(txtInfo.Text) = ""
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
'���س�
    Dim strCard As String
    
    strCard = idkSelect.Cards(idkSelect.IDKind).����

    If mblnReadCard Or KeyAscii = 13 Then
        Call cmdOk_Click
        Call zlControl.TxtSelAll(txtInfo)
        
        '��λҳǩ
        If patiList.Count >= 1 Then
            Call SelectTabItem
        End If

    Else
        Select Case strCard
            Case "�����"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "�Һŵ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtInfo.Text = "" Or txtInfo.SelLength = Len(txtInfo.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "���֤��", "�������֤"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
    mblnReadCard = False
End Sub

Private Sub txtInfo_LostFocus()
    idkSelect.SetAutoReadCard False
End Sub

'Private Sub txtInfo_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Private Sub txtNo1_KeyPress(KeyAscii As Integer)
    '�ҵ��Һŵ���ִ����Һ
    If KeyAscii = vbKeyReturn Then
        Call FindAndExe(rptQueue1, txtNo1, 1)
    End If
End Sub

Private Sub FindAndExe(objRpt As ReportControl, ByVal strNoIn As String, ByVal intLiquidOrPut As Integer)

    '���Ҳ�ִ�� ��Һ �� ����
    Dim strNO As String, i As Integer
    strNO = GetFullNO(strNoIn, 12)  '12���Һ��վݺ�
    
    For i = i To objRpt.Rows.Count - 1
        With objRpt.Rows(i)
            If Not .GroupRow Then
                If .Record(col_�Һŵ�).Value = strNO Then Exit For
            End If
        End With
    Next

    If i <= objRpt.Rows.Count - 1 Then
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set objRpt.FocusedRow = objRpt.Rows(i)
        objRpt.SetFocus
        If intLiquidOrPut = 1 Then
            '��Һ������
            Call LiquidAndPlay
        Else
            '����
            Call Puncture
        End If
    Else
        MsgBox "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtNo5_KeyPress(KeyAscii As Integer)
    '�ҵ��Һŵ���ִ����Һ
    If KeyAscii = vbKeyReturn Then
        Call FindAndExe(rptQueue5, txtNo5, 2)
    End If
End Sub

Private Function CheckComplt(ByVal str��ˮ�� As String, ByVal strGroupKey As String) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngID As Long, lngSend As Long  'ҽ��ID,���ͺ�
    Dim objPati As cPatient
    
    On Error GoTo hErr
    
    CheckComplt = True
    lngID = Val(Split(strGroupKey, "_")(0))
    lngSend = Val(Split(strGroupKey, "_")(1))
    

    '�Ƿ���δ��д����Ա�ļ�¼����������ɡ�
    strSQL = "select ִ��ʱ�� from ����ҽ��ִ�� where ִ���� Is Null and ҽ��ID=[1] and ���ͺ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID, lngSend)
    If Not rsTmp.EOF Then
        CheckComplt = False
        Set objPati = patiList.Item(mstr�Һŵ�)
        SaveOperLog mlngPreDept, objPati, MEDICAL, "ҽ�����ͺ�Ϊ" & strGroupKey & "��" & Format(rsTmp!ִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "�ļ�¼��δ��ʼ"
    End If
    Exit Function
hErr:
    CheckComplt = False
End Function

Private Sub FindCboIndex(ByVal objCbo As Object, ByVal lngData As Long, Optional ByVal blnKeep As Boolean)
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not blnKeep Then objCbo.ListIndex = -1
End Sub

Private Sub SelectTabItem()
'���ܣ����Ҷ�λ�����˵�ǰ״̬ҳǩ

    Dim i As Integer
    Dim strTag As String

    strTag = patiList.Item(1).�Ŷ�״̬
    If InStr(strTag, "-") > 0 Then strTag = Mid(strTag, InStr(strTag, "-") + 1)
    
    'ͳһ���ƴ����Ŷ�״̬=��������ҳǩ����=�ѽ���
    If strTag = "����" Then strTag = "�ѽ���"

'    Dim strSQL As String
'    Dim lngID As Long
'    Dim rsTmp As ADODB.Recordset
'    Dim strTag As String
'
'    lngID = patiList.Item(1).����ID
'
    On Error GoTo errHandle
'
'    strSQL = "Select decode(a.״̬, 1, '����Һ', 5, '������', 6, 'ִ����', 7, '�ѽ���', '') ״̬ " & _
'             "From �ŶӼ�¼ A, ������Ϣ B Where a.����id = b.����id And b.����id = [1] "
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "��ȡ�����ŶӼ�¼��״̬", lngID)
'    If rsTmp.EOF = False Then strTag = zlcommfun.NVL(rsTmp!״̬)
'    rsTmp.Close
'
    For i = 1 To tbcList.ItemCount - 1
        If tbcList.Item(i).Tag = strTag And tbcList.Item(i).Visible Then
            tbcList.Item(i).Selected = True
            Exit For
        End If
    Next

    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Function FuncThingAudit(ByVal str��ˮ�� As String, strGroupKey As String) As Boolean
'���ܣ��˶�
    Dim strSQL As String
    Dim str�˶��� As String

    With mobjRecord.Item(CStr(str��ˮ��)).Item(strGroupKey)
        If .�˶��� <> "" Then
            MsgBox "��ҽ�����Ѿ��˶ԣ������ٴκ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        If .ִ���� = "" Then
            MsgBox "��ҽ����δ�Ǽ�ִ���ˣ����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        str�˶��� = zlDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", glngSys, 1263, "ִ������Ǽ�", , True)
        If str�˶��� = "" Then Exit Function
        
        If str�˶��� = .ִ���� Then
            MsgBox "ִ���˲��ܺ��������ͬ�����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If

        On Error GoTo errH
        strSQL = "Zl_����ҽ���˶�_Insert(" & Val(.ִ��ҽ��ID) & "," & Val(.���ͺ�) & ",'" & str�˶��� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "ҽ���˶�")
        .�˶��� = str�˶���
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function FuncThingDelAudit(ByVal str��ˮ�� As String, strGroupKey As String) As Boolean
'���ܣ�ȡ���˶�
    Dim strSQL As String
    Dim str�˶��� As String

    With mobjRecord.Item(CStr(str��ˮ��)).Item(strGroupKey)
        If .�˶��� = "" Then
            MsgBox "��ҽ����δ���к˶ԣ�����ȡ����", vbInformation, gstrSysName
            Exit Function
        End If

        If .�˶��� <> UserInfo.���� Then
            str�˶��� = zlDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", glngSys, 1263, "ִ������Ǽ�", , True)
            If str�˶��� = "" Then Exit Function
            If str�˶��� <> .�˶��� Then
                MsgBox "ֻ��ȡ���Լ��˶Ե�ҽ������ǰҽ���˶�����""" & .�˶��� & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        
        strSQL = "Zl_����ҽ���˶�_Delete(" & Val(.ִ��ҽ��ID) & "," & Val(.���ͺ�) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "ȡ��ҽ���˶�")
        .�˶��� = ""
        FuncThingDelAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

