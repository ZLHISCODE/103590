VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExamineKSS 
   Caption         =   "������ҩ���"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15165
   Icon            =   "frmExamineKSS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   12480
      TabIndex        =   38
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   40
         Top             =   -10
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "סԺ"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   39
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "ʹ�ó���"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame fraPati 
      Caption         =   "������Ϣ"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   10
      Top             =   600
      Width           =   11295
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   697
         Width           =   4815
      End
      Begin VB.PictureBox picInShow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   8445
         TabIndex        =   11
         Top             =   360
         Width           =   8450
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   16
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   14
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblCaption 
            Caption         =   "��Ժʱ�䣺"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblCaption 
            Caption         =   "���ţ�"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   20
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "����ȼ���"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   19
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblCaption 
            Caption         =   "������"
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   18
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "���أ�"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   22
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblCaption 
         Caption         =   "��ϣ�"
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         Caption         =   "����ҩ�"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         Caption         =   "���䣺"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "�Ա�"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox picUnAudited 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4440
      ScaleHeight     =   5895
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   2640
      Width           =   9735
      Begin VB.PictureBox picDateY 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   42
         Top             =   0
         Width           =   9375
         Begin VB.ComboBox cboDateY 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdFindY 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   45
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   121176067
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   46
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   121176067
            CurrentDate     =   40256
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2460
            TabIndex        =   48
            Top             =   90
            Width           =   1890
         End
         Begin VB.Label lblDateY 
            Caption         =   "����ʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   75
            Width           =   735
         End
      End
      Begin VB.PictureBox picDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   30
         Top             =   120
         Width           =   9375
         Begin VB.CommandButton cmdFind 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   30
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   33
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   121176067
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   34
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   121176067
            CurrentDate     =   40256
         End
         Begin VB.Label lblDate 
            Caption         =   "���ʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2460
            TabIndex        =   35
            Top             =   90
            Width           =   1890
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAudit 
         Height          =   4860
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   8505
         _cx             =   15002
         _cy             =   8572
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmExamineKSS.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7335
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   11355
      _Version        =   589884
      _ExtentX        =   20029
      _ExtentY        =   12938
      _StockProps     =   64
   End
   Begin VB.Frame fraDoctor 
      Caption         =   "ҽ��"
      ForeColor       =   &H000040C0&
      Height          =   8775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   7140
         Left            =   100
         TabIndex        =   2
         Top             =   1500
         Width           =   3330
         _Version        =   589884
         _ExtentX        =   5874
         _ExtentY        =   12594
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chkIsShowAll 
         Caption         =   "ֻ��ʾ�������ҽ��"
         Height          =   180
         Left            =   1080
         TabIndex        =   37
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   788
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   315
         TabIndex        =   5
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   315
         TabIndex        =   4
         Top             =   420
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9870
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   635
      SimpleText      =   $"frmExamineKSS.frx":68ED
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExamineKSS.frx":6934
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21669
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
   Begin MSComctlLib.ImageList img16 
      Left            =   600
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
            Picture         =   "frmExamineKSS.frx":71C8
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":DA2A
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":1428C
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":14826
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgAdvice 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":14DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":1535A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineKSS.frx":158F4
            Key             =   "ǩ��"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmExamineKSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCodeType As Long         '0-ƴ��,1-���
Private mobjBar As CommandBar
Private mobjPopup As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean
Private mobjESign As Object '����ǩ���ӿڲ���
Private mblnTeam As Boolean '��С��������  ϵͳ��������ҽ��С����п���ҩ�����

Private mlngFindNum As Long
Private mstrChangeRows As String   '��¼�޸ĵ���
Private mstrǩ��IDs As String      'ȡ����˵�ʱ���¼һ�����˴������ǩ��ID
Private mblnTmp As Boolean
Private mrsDefine As ADODB.Recordset
Private mclsMipModule As zl9ComLib.clsMipModule
Private Enum Enum_Dor
    COL_��ԱID = 0
    col_���� = 1
    COL_רҵְ�� = 2
    COL_������ҩȨ�� = 3
    COL_ƴ������ = 4
    COL_��ʼ��� = 5
    COL_�������� = 6
    COL_��������ID = 7
End Enum

Private Enum Enum_Advice
    col_ѡ�� = 0
    COL_ȡ��ѡ�� = 1
    COL_���˵�� = 2
    COL_���ʱ�� = 3
    COL_�������� = 4
    COL_ҽ������ = 5
    col_��ҩĿ�� = 6
    col_��ҩ���� = 7
    col_��Ч = 8
'�ü��ģʽ�����������͵���������������ҽ�����ݺϲ�
    COL_���� = 9
    COL_���� = 10
    COL_Ƶ�� = 11
    col_��ҩ;�� = 12
    COL_��ʼʱ�� = 13
    COL_��ֹʱ�� = 14
    col_ִ��ʱ�䷽�� = 15
'������
    col_ҽ��ID = 16
    col_���ID = 17
    col_�Ա� = 18
    col_���� = 19
    COL_���� = 20
    COL_��Ժʱ�� = 21
    col_���� = 22
    col_������ = 23
    COL_���� = 24
    col_����ȼ� = 25
    col_����Id = 26
    col_��ҳID = 27
    col_�Һŵ� = 28
    COL_��ID = 29
    COL_������� = 30
    COL_������Դ = 31
    col_�Һŵ��� = 32
    COL_ǩ��id = 33
    COL_ҽ��״̬ = 34
    
    COL_����� = 35
    col_סԺ�� = 36
    COL_��ǰ���� = 37
    COL_����ҽ�� = 38
    COL_����ʱ�� = 39
    COL_��������ID = 40
    COL_��Ժ����ID = 41
    COL_��ǰ����ID = 42
End Enum

Private Enum enum_Info
    info_��Ժʱ�� = 0
    info_�Ա� = 1
    info_���� = 2
    info_���� = 3
    info_����ȼ� = 4
    info_���� = 5
    info_��� = 6
    info_���� = 7
End Enum

Public Function ShowMe(frmParent As Object, Optional ByRef ojbMip As Object)
'�����ӿ�
    On Error Resume Next
    
    If Not ojbMip Is Nothing Then Set mclsMipModule = ojbMip
    
    Call frmExamineKSS.Show(0, frmParent)

End Function

Private Sub cboDateY_Click()
    Dim curDate As Date
    
    dtpTimeY(0).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    dtpTimeY(1).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTimeY(0).MaxDate = curDate
    dtpTimeY(1).MaxDate = curDate
    cmdFindY.Visible = False
    
    Select Case cboDateY.ListIndex
    Case 0 '����
        dtpTimeY(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '�������
        dtpTimeY(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '�������
        dtpTimeY(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '���һ��
        dtpTimeY(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '���һ��
        dtpTimeY(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 'ָ  ��
        If Me.Visible Then dtpTimeY(0).SetFocus
        cmdFindY.Visible = True
    End Select
    
    If cboDateY.ListIndex <> cboDateY.ListCount - 1 Then
        If Me.Visible Then Call LoadAdvice
    End If
End Sub

Private Sub cboDept_Click()
    Call LoadDoc
End Sub

Private Sub LoadDoc()
'����Ȩ�ޱȲ���Ա�͵�ҽ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    Dim datBegint As Date
    Dim datEnd As Date
 
    If cboDept.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 11
    stbThis.Panels(2).Text = "��ѡ��һλ����ҽ����"
    
    If tbcSub.Selected.Tag = "�����" Then
        datBegint = CDate(dtpTimeY(0).Value)
        datEnd = CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct ����ҽ�� From ����ҽ����¼ F Where  f.���״̬ = 1 And f.ҽ��״̬<>4 and (f.ҽ��״̬=1 or f.ҽ��״̬>2 and f.������־=1) And F.����ʱ�� Between [4] And [5] And f.������� In ('5','6')) F"
    Else
        datBegint = CDate(dtpTime(0).Value)
        datEnd = CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct f.����ҽ�� From ����ҽ����¼ F,����ҽ��״̬ G Where f.id=g.ҽ��id and G.�������� in (11,12)" & _
            " And G.����ʱ�� Between [4] And [5] And f.������� In ('5','6')) F"
    End If
    
    strSQL = "Select DISTINCT a.Id, A.�Ա�," & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "-Null as ����ID,Null as ��������,", "b.����ID,e.���� as ��������,") & _
            " a.����,a.רҵ����ְ��,Decode(c.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','��ʹ��Ȩ��') as ������ҩȨ��, Upper(zlSpellCode(a.����)) As ƴ������, Upper(Zlwbcode(a.����)) As ��ʼ���" & _
            " From ��Ա�� A, ������Ա B, ��Ա����ҩ��Ȩ�� C, ��Ա����˵�� D,���ű� E" & _
            IIf(chkIsShowAll.Value, strTmp, "") & _
            " Where c.��Աid(+) = a.Id And a.Id = b.��Աid And e.ID=b.����ID And d.��Աid = a.Id And C.����=[3] And d.��Ա���� = 'ҽ��'" & _
            " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (((c.��¼״̬ = 1 And c.���� <[1]) Or (c.��Աid Is Null))) " & _
            IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.����id=[2]") & _
            IIf(chkIsShowAll.Value, " And  f.����ҽ�� = a.���� ", "")
    
    If mblnTeam Then
        If Val(cboDept.ItemData(cboDept.ListIndex)) = -1 Then
            strSQL = "select k.id,k.�Ա�,f.id as ����id,f.����||','||m.���� as ��������,k.����,k.רҵ����ְ��,k.������ҩȨ��,k.ƴ������,k.��ʼ���" & _
                " from �ٴ�ҽ��С�� m,ҽ��С����Ա n,���ű� f,(" & strSQL & ") k" & vbNewLine & _
                " where m.id=n.С��id and n.��Աid=k.id and m.����id=f.id and Exists (select 1 from ҽ��С����Ա b where m.id=b.С��id and b.��Աid=[6])" & _
                " And (m.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or m.����ʱ�� Is Null)"
        Else
            strSQL = "select k.id,k.�Ա�,k.����id,k.��������||','||m.���� as ��������,k.����,k.רҵ����ְ��,k.������ҩȨ��,k.ƴ������,k.��ʼ���" & _
                " from �ٴ�ҽ��С�� m,ҽ��С����Ա n,(" & strSQL & ") k" & vbNewLine & _
                " where m.id=n.С��id and n.��Աid=k.id and Exists (select 1 from ҽ��С����Ա b where m.id=b.С��id and b.��Աid=[6])" & _
                " And (m.����ʱ��=To_Date('3000-01-01', 'YYYY-MM-DD') Or m.����ʱ�� Is Null) And m.����ID=[2]"
        End If
    End If
    
    On Error GoTo errH
    
    rptDoc.Records.DeleteAll
    vsAudit.Rows = 1: vsAudit.AddItem ""
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)), IIf(optOccasion(0).Value, 1, 2), datBegint, datEnd, UserInfo.ID)
    
    With rptDoc
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!�Ա� & "" = "Ů", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!רҵ����ְ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!������ҩȨ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!ƴ������ & "")
            Set objItem = objRecord.AddItem(rsTmp!��ʼ��� & "")
            Set objItem = objRecord.AddItem(rsTmp!�������� & "")
            Set objItem = objRecord.AddItem(Val(rsTmp!����ID & ""))
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    mlngFindNum = 0
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptDoc.Visible = False Then Exit Sub
    If rptDoc.Records.Count > 0 Then
        If rptDoc.SelectedRows.Count = 0 Then Exit Sub
        strSubhead = rptDoc.SelectedRows(0).Record(col_����).Value & "ҽ������嵥"
    Else
        Exit Sub
    End If
    
    '���ô�ӡ��������
    Set objPrint.Body = Me.vsAudit
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Cancle()
'���ܣ�ȡ������
    Dim i As Long
    With vsAudit
        If MsgBox("�����޸ĵ�����δ���棬�Ƿ������", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            If tbcSub.Selected.Tag = "�����" Then
                Call LoadAdvice(True)
            Else
                Call LoadAdvice
            End If
            mblnIsUpdate = False
            mstrChangeRows = ""
        End If
    End With
End Sub

Private Sub SaveAudit()
'���ܣ����������Ϣ
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strDate As String
    
    With vsAudit
        If .EditText <> "" Then .TextMatrix(.Row, .Col) = .EditText
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        If tbcSub.Selected.Tag = "�����" Then
            For i = 1 To .Rows - 1
                'һ�����˵���һ��
                If RowInͬһ����(i, lngBegin, lngEnd, vsAudit) Then
                    Call SaveAuditOnePati(lngBegin, lngEnd, strDate)
                    i = lngEnd
                Else
                    Call SaveAuditOnePati(i, i, strDate)
                End If
            Next
            Call LoadAdvice
        Else
            Call SaveAuditUpdate
            Call LoadAdvice(True)
        End If
        mstrChangeRows = ""
        mblnIsUpdate = False
    End With
End Sub

Private Sub SaveAuditUpdate()
'���ܣ��޸������δͨ�������˵��
    Dim i As Long
    Dim strSQL As String
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strDate As String
    Dim varArr As Variant
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    If mstrChangeRows <> "" Then
        varArr = Split(mstrChangeRows, ",")
        With vsAudit
            If .EditText <> "" Then .TextMatrix(.Row, .Col) = .EditText
            For i = 0 To UBound(varArr)
                If .TextMatrix(Val(varArr(i)), col_ҽ��ID) <> "" And Val(varArr(i)) <> 0 Then
                    strSQL = "Zl_������ҩ���_Update(" & Val(.TextMatrix(Val(varArr(i)), col_ҽ��ID)) & "," & strDate & ",'" & .TextMatrix(Val(varArr(i)), COL_���˵��) & "')"
                    colsql.Add strSQL, "C" & colsql.Count + 1
                End If
            Next
        End With
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colsql.Count
            Call zlDatabase.ExecuteProcedure(CStr(colsql(i)), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, ByVal strDate As String)
'���ܣ����������Ϣ
'�������ӵڼ��п�ʼ�����ڼ��н�����ͬһ�����ˣ�
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String
    Dim strSource As String, strSign As String
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim strSignSQL As String
    Dim lngMsgRow As Long
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .TextMatrix(i, col_ҽ��ID) = "" Then Exit Sub
            If Val(.Cell(flexcpData, i, col_ѡ��) & "") <> "0" Then
                If Not RowInһ����ҩ(i, lngGroupBegin, lngGroupEnd, vsAudit) Then
                    strSQL = Val(.TextMatrix(i, col_ҽ��ID)) & "|" & "Zl_������ҩ���_Audit(" & Val(.TextMatrix(i, col_ҽ��ID)) & "," & Val(.Cell(flexcpData, i, col_ѡ��) & "") & "," & _
                            "'" & UserInfo.���� & "'," & strDate & ",'" & .TextMatrix(i, COL_���˵��) & "'"
                    colsql.Add strSQL, "C" & colsql.Count + 1
                    If Val(.Cell(flexcpData, i, col_ѡ��) & "") = 1 Then
                        strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                    End If
                Else
                    'һ��ҩƷ
                    For j = lngGroupBegin To lngGroupEnd
                        strSQL = Val(.TextMatrix(j, col_ҽ��ID)) & "|" & "Zl_������ҩ���_Audit(" & Val(.TextMatrix(j, col_ҽ��ID)) & "," & Val(.Cell(flexcpData, i, col_ѡ��) & "") & "," & _
                            "'" & UserInfo.���� & "'," & strDate & ",'" & .TextMatrix(j, COL_���˵��) & "'"
                        colsql.Add strSQL, "C" & colsql.Count + 1
                        If Val(.Cell(flexcpData, j, col_ѡ��) & "") = 1 Then
                            strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(j, col_ҽ��ID)
                        End If
                    Next
                    i = lngGroupEnd
                End If
                '��ҩ��ʽ
                strSQL = Val(.TextMatrix(i, col_���ID)) & "|" & "Zl_������ҩ���_Audit(" & Val(.TextMatrix(i, col_���ID)) & "," & Val(.Cell(flexcpData, i, col_ѡ��) & "") & "," & _
                        "'" & UserInfo.���� & "'," & strDate & ",''"
                colsql.Add strSQL, "C" & colsql.Count + 1
                If Val(.Cell(flexcpData, i, col_ѡ��) & "") = 1 Then
                    strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_���ID)
                End If
                If Val(.Cell(flexcpData, i, col_ѡ��) & "") = 1 Then
                    lngMsgRow = i
                End If
            End If
        Next
        '��ȡǩ��ҽ��Դ��
        If gintCA <> 0 And strIDs <> "" Then
            If Val(.TextMatrix(lngBegin, COL_������Դ)) = 0 Then Exit Sub
            '��������˰����ҿ��Ƶ���ǩ��ʱ�������е���ǩ�����ơ�
            If Mid(gstrESign, Val(.TextMatrix(lngBegin, COL_������Դ)), 1) = "1" And CheckSign(Val(.TextMatrix(lngBegin, COL_������Դ)) - 1, -1, , , , , mobjESign) Then
                If mobjESign Is Nothing Then
                    On Error Resume Next
                    Set mobjESign = CreateObject("zl9ESign.clsESign")
                    err.Clear: On Error GoTo 0
                    If Not mobjESign Is Nothing Then
                        Call mobjESign.Initialize(gcnOracle, glngSys)
                    Else
                        MsgBox "����ǩ������δ����ȷ��װ����˲������ܼ�����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                intRule = ReadAdviceSignSource(11, Val(.TextMatrix(lngBegin, col_����Id)), IIf(Val(.TextMatrix(lngBegin, COL_������Դ)) = 1 _
                        , .TextMatrix(lngBegin, col_�Һŵ���), Val(.TextMatrix(lngBegin, col_��ҳID))), strIDs, 0, False, strSource)
                If intRule = 0 Then Screen.MousePointer = 0: Exit Sub
                If strSource = "" Then
                    Screen.MousePointer = 0
                    MsgBox "���ܶ�ȡ��Ҫ��˵���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
        
                strSign = mobjESign.signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
                If strSign <> "" Then
                    If strTimeStamp <> "" Then
                        strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        strTimeStamp = "NULL"
                    End If
                    lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
                    strSignSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",11," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
                Else
                    Screen.MousePointer = 0: Exit Sub
                End If
            End If
        Else
            Set mobjESign = Nothing
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colsql.Count
            strSQL = Mid(colsql("C" & i), InStr(colsql("C" & i), "|") + 1) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
        If strSignSQL <> "" Then
            Call zlDatabase.ExecuteProcedure(strSignSQL, Me.Caption)
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If lngMsgRow <> 0 Then
        '����ҽ�����´���Ϣ
        With vsAudit
            If Val(.TextMatrix(lngMsgRow, COL_������Դ)) = 2 Then
                If HaveOperateAdvice(Val(.TextMatrix(lngMsgRow, col_����Id)), Val(.TextMatrix(lngMsgRow, col_��ҳID)), 0) Then
                i = IIf(.TextMatrix(lngMsgRow, col_��Ч) = "����", 1, 0)
                Call ZLHIS_CIS_001(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_����Id)), .TextMatrix(lngMsgRow, COL_��������), .TextMatrix(lngMsgRow, col_סԺ��), "", 2, Val(.TextMatrix(lngMsgRow, col_��ҳID)), _
                    Val(.TextMatrix(lngMsgRow, COL_��ǰ����ID)), "", Val(.TextMatrix(lngMsgRow, COL_��Ժ����ID)), "", "", .TextMatrix(lngMsgRow, COL_��ǰ����), Val(.TextMatrix(lngMsgRow, col_ҽ��ID)), 0, i, 5, 2, _
                    .TextMatrix(lngMsgRow, COL_����ҽ��), .TextMatrix(lngMsgRow, COL_����ʱ��), .TextMatrix(lngMsgRow, COL_��������ID), "")
                End If
            End If
        End With
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTime(0).MaxDate = curDate
    dtpTime(1).MaxDate = curDate
    cmdFind.Visible = False
    
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '�������
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '�������
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '���һ��
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '���һ��
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 'ָ  ��
        If Me.Visible Then dtpTime(0).SetFocus
        cmdFind.Visible = True
    End Select
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub CancleAudit()
'ȡ�����
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnIsCheck As Boolean
    
    With vsAudit
        '�ж��Ƿ��й�ѡ�ģ��й�ѡ���Թ�ѡΪ׼
        If MsgBox("ȡ����˵�ҽ�����ڴ������������ˣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, COL_ȡ��ѡ��))) = 1 Then Exit For
        Next
        blnIsCheck = i < .Rows
        
        If blnIsCheck Then
            For i = i To .Rows - 1
                If Abs(Val(.TextMatrix(i, COL_ȡ��ѡ��))) = 1 Then
                    If RowInͬһ����(i, lngBegin, lngEnd, vsAudit) Then
                        Call CancleAuditOnePati(lngBegin, lngEnd)
                        i = lngEnd
                    Else
                        Call CancleAuditOnePati(i, i)
                    End If
                End If
            Next
        Else
            If .Row = 0 Then Exit Sub
            If gintCA > 0 Then
                If RowInͬһ����(.Row, lngBegin, lngEnd, vsAudit) And Val(.TextMatrix(.Row, COL_ǩ��id)) <> 0 Then
                    '�����ѡ�����������õݹ飬ֱ�Ӵ����ѡ����ǩ��IDһ����ҽ��
                    Call CancleAuditOnePati(lngBegin, lngEnd, Not blnIsCheck, Val(.TextMatrix(.Row, COL_ǩ��id)), False)
                Else
                    Call CancleAuditOnePati(.Row, .Row, Not blnIsCheck)
                End If
            Else
                Call CancleAuditOnePati(.Row, .Row, Not blnIsCheck)
            End If
        End If
        Call LoadAdvice(True)
    End With
End Sub

Private Sub CancleAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, Optional ByVal blnIsNoCheck As Boolean, _
        Optional ByVal lngǩ��ID_IN As Long, Optional ByVal blnIsRecursive As Boolean = True)
'���ܣ�ȡ�����
'������lngBegin�ӵڼ��п�ʼ��lngEnd���ڼ��н�����ͬһ�����ˣ�
'     blnIsNoCheck=û�й�ѡ����ѡ����Ϊ׼ȡ�����
'     lngǩ��ID_IN�����ڵݹ���ã������һ��ѭ���з�����ǩ��ID<>0����ݹ���ñ����������������ǩ��ID���룬
'    ���뵽�ַ���mstrǩ��IDs��ڶ��ν�������ǩ��ID��ҽ��,����ٷ����봫���ǩ��ID��һ���������ֲ����ַ���mstrǩ��IDs�У���Ϊ�µģ����ٵݹ���á�
'    blnIsRecursive:�Ƿ�ݹ飬Ĭ��ΪҪ�ݹ�
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String, blnTrans As Boolean
    Dim strSource As String, strSign As String
    Dim lng֤��ID As Long, lngǩ��ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    
    With vsAudit
        If gintCA > 0 Then
            For i = lngBegin To lngEnd
                If Abs(Val(.TextMatrix(i, COL_ȡ��ѡ��))) = 1 Or blnIsNoCheck Then
                    If Not RowInһ����ҩ(i, lngGroupBegin, lngGroupEnd, vsAudit) Then
                        If Val(.TextMatrix(i, COL_ǩ��id)) <> lngǩ��ID_IN Then
                            If lngǩ��ID = 0 And InStr("," & mstrǩ��IDs & ",", "," & Val(.TextMatrix(i, COL_ǩ��id)) & ",") = 0 Then
                                lngǩ��ID = Val(.TextMatrix(i, COL_ǩ��id))
                            End If
                        Else
                            strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        End If
                    Else
                        'һ��ҩƷ
                        For j = lngGroupBegin To lngGroupEnd
                            If Val(.TextMatrix(j, COL_ǩ��id)) <> lngǩ��ID_IN Then
                                If lngǩ��ID = 0 And InStr("," & mstrǩ��IDs & ",", "," & Val(.TextMatrix(j, COL_ǩ��id)) & ",") = 0 Then
                                    lngǩ��ID = Val(.TextMatrix(j, COL_ǩ��id))
                                End If
                            Else
                                strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(j, col_ҽ��ID)
                            End If
                        Next
                        i = lngGroupEnd
                    End If
                    '��ҩ��ʽ
                    If Val(.TextMatrix(i, COL_ǩ��id)) <> lngǩ��ID_IN Then
                        If lngǩ��ID = 0 And InStr("," & mstrǩ��IDs & ",", "," & Val(.TextMatrix(i, COL_ǩ��id)) & ",") Then
                            lngǩ��ID = Val(.TextMatrix(i, COL_ǩ��id))
                        End If
                    Else
                        strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_���ID)
                    End If
                End If
            Next
            
            If lngǩ��ID_IN <> 0 Then
                strSign = "zl_ҽ��ǩ����¼_Delete(" & lngǩ��ID_IN & ")"
            End If
            '����ܷ����ǩ��
            If strSign <> "" Then
                If mobjESign Is Nothing Then
                    On Error Resume Next
                    Set mobjESign = CreateObject("zl9ESign.clsESign")
                    err.Clear: On Error GoTo 0
                    If Not mobjESign Is Nothing Then
                        Call mobjESign.Initialize(gcnOracle, glngSys)
                    End If
                End If
                If mobjESign Is Nothing Then
                    If gintCA = 0 Then
                        MsgBox "ϵͳû�����õ���ǩ����֤���ģ����˲������ܼ�����", vbInformation, gstrSysName
                    Else
                        MsgBox "����ǩ������δ����ȷ��װ�����˲������ܼ�����", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    If Not mobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
                End If
            End If
        Else
            For i = lngBegin To lngEnd
                If Abs(Val(.TextMatrix(i, COL_ȡ��ѡ��))) = 1 Or blnIsNoCheck Then
                    If Not RowInһ����ҩ(i, lngGroupBegin, lngGroupEnd, vsAudit) Then
                        strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                    Else
                        'һ��ҩƷ
                        For j = lngGroupBegin To lngGroupEnd
                            strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(j, col_ҽ��ID)
                        Next
                        i = lngGroupEnd
                    End If
                    '��ҩ��ʽ
                    strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_���ID)
                End If
            Next
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    'ȡ��ǩ��
    If gintCA > 0 And strSign <> "" Then
        Call zlDatabase.ExecuteProcedure(strSign, Me.Caption)
    End If
    'ȡ�����
    If strIDs <> "" Then
        strSQL = "Zl_������ҩ���_Cancel('" & strIDs & "','" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    'ִ������ж��Ƿ�����ǩ��ID��ҽ����Ȼ��ݹ����
    If blnIsRecursive Then
        If lngǩ��ID <> 0 Then
            mstrǩ��IDs = mstrǩ��IDs & "," & lngǩ��ID
            Call CancleAuditOnePati(lngBegin, lngEnd, blnIsNoCheck, lngǩ��ID)
        End If
    End If
    mstrǩ��IDs = "0"
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext And Control.ID <> conMenu_Edit_Audit Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    Case conMenu_Edit_Audit '�Ҽ��������
        For i = vsAudit.FixedRows To vsAudit.Rows - 1
            If Val(vsAudit.Cell(flexcpData, i, col_ѡ��)) <> 0 Then Exit For
        Next
        If i = vsAudit.Rows Then Call AuditStateCheck(1)
        If i < vsAudit.Rows And Val(vsAudit.Cell(flexcpData, vsAudit.RowSel, col_ѡ��)) = 0 Then
            If MsgBox("������˲���ֻ�����ѹ�ѡ��ҽ�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
        End If
        Call SaveAudit
    Case conMenu_Edit_Untread     'ȡ��
        Call Cancle
    Case conMenu_Edit_Save        '����
        Call SaveAudit
    Case conMenu_Edit_AdviceUnAudit 'ȡ�����
        Call CancleAudit
    Case conMenu_Tool_Archive '���Ӳ�������
        If vsAudit.Row = 0 Or vsAudit.TextMatrix(1, col_ҽ��ID) = "" Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_����Id)), IIf(Val(vsAudit.TextMatrix(vsAudit.Row, COL_������Դ)) = 2, Val(vsAudit.TextMatrix(vsAudit.Row, col_��ҳID)) _
                , Val(vsAudit.TextMatrix(vsAudit.Row, col_�Һŵ�))))
    Case conMenu_View_Find '����
        txtFind.SetFocus '��ʱ��Ҫ��λһ��
        If txtFind.Text <> "" Then
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_FindNext '������һ��
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_View_Refresh 'ˢ��
        If tbcSub.Selected.Tag = "�����" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptDoc.SelectedRows.Count = 0 Or vsAudit.Row <= 0 Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex))
            Else
                With vsAudit
                    If .TextMatrix(.Row, COL_������Դ) = "2" Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex), "�����=" & rptDoc.SelectedRows(0).Record(col_����).Value, _
                            "����ID=" & .TextMatrix(.Row, col_����Id), "��ҳID=" & .TextMatrix(.Row, col_��ҳID), "ҽ��ID=" & .TextMatrix(.Row, col_ҽ��ID))
                    Else
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex), "�����=" & rptDoc.SelectedRows(0).Record(col_����).Value, _
                            "����ID=" & .TextMatrix(.Row, col_����Id), "�Һŵ�=" & .TextMatrix(.Row, col_�Һŵ�), "ҽ��ID=" & .TextMatrix(.Row, col_ҽ��ID))
                    End If
                End With
            End If
        End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With fraDoctor
        .Top = lngTop
        .Left = lngLeft + 100
        .Height = lngBottom - lngTop - stbThis.Height
    End With
    rptDoc.Height = fraDoctor.Height - 1600
    
    With fraPati
        .Top = fraDoctor.Top
        .Left = fraDoctor.Left + fraDoctor.Width + 45
        .Width = lngRight - fraDoctor.Width - 200
    End With
    
    With tbcSub
        .Top = fraPati.Top + fraPati.Height + 45
        .Left = fraPati.Left
        .Height = fraDoctor.Height - fraPati.Height - 45
        .Width = fraPati.Width + 50
    End With
    
    Me.Refresh
End Sub

Private Function CheckKssJuris() As Boolean
'����û��Ƿ��п�����Ȩ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnTmp As Boolean
    Dim lngLevel As Long
    
    strSQL = "Select NVL(MAX(����),0) as ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬ = 1 And ��Աid =[1] And ����=[2]"

    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, IIf(optOccasion(0).Value, 1, 2))
    
    lngLevel = Val(rsTmp!���� & "")
    '����סԺ���ﶼû��Ȩ�޲���ʾ
    If lngLevel = 0 Then
        blnTmp = True
        If Me.Visible = False Then
            strSQL = "Select NVL(MAX(����),0) as ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬ = 1 And ��Աid =[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
            If Val(rsTmp!���� & "") <> 0 Then blnTmp = False
        End If
        If blnTmp Then
            MsgBox "��û���㹻��Ȩ��,�������Ա��ϵ��", vbInformation, Me.Caption
            CheckKssJuris = False
            Exit Function
        Else    '���������Ȩ�ޣ����Զ��л������ֻ�г�ʼ���������öδ���
            mblnTmp = True
            optOccasion(IIf(optOccasion(0).Value, 1, 0)).Value = True
            mblnTmp = False
            Call CheckKssJuris
            lngLevel = mlngLevel
        End If
    End If
    mlngLevel = lngLevel
    CheckKssJuris = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '����Ȩ�����ð�ť�ɼ�״̬
    
    Select Case Control.ID
        Case conMenu_Edit_AdviceUnAudit
            If tbcSub.Selected.Tag <> "�����" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_Untread, conMenu_Edit_Save, conMenu_Edit_Audit
            If tbcSub.Selected.Tag = "�����" Then Control.Visible = False: Exit Sub
        Case conMenu_Tool_Archive '���Ӳ�������
            If GetInsidePrivs(p���Ӳ�������) = "" Then
                Control.Visible = False
                Exit Sub
            End If
    End Select
    Control.Visible = True
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
    
        Case conMenu_Edit_Untread, conMenu_Edit_Save   '����,ȡ��
            Control.Enabled = mblnIsUpdate
        Case conMenu_View_Refresh, conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'ˢ��,��ӡ
            Control.Enabled = Not mblnIsUpdate
            If mblnIsUpdate Then
                cboDept.Enabled = False
                txtFind.Enabled = False
                fraDoctor.Enabled = False
                cboDept.BackColor = &H8000000F
                txtFind.BackColor = &H8000000F
                cmdFind.Enabled = True
                cboTime.Enabled = False
                tbcSub.Item(IIf(tbcSub.Selected.Index = 0, 1, 0)).Enabled = False
            Else
                cboDept.Enabled = True
                txtFind.Enabled = True
                fraDoctor.Enabled = True
                cboTime.Enabled = True
                cmdFind.Enabled = True
                cboDept.BackColor = &H80000005
                txtFind.BackColor = &H80000005
                tbcSub.Item(IIf(tbcSub.Selected.Index = 0, 1, 0)).Enabled = True
            End If
        Case conMenu_Edit_AdviceUnAudit 'ȡ�����
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, COL_ҽ��״̬) = "1"
        Case conMenu_Tool_Archive '���Ӳ�������
            Control.Enabled = vsAudit.Row <> 0 And vsAudit.TextMatrix(1, col_ҽ��ID) <> ""
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '������һ��
            Control.Visible = False
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkIsShowAll_Click()
    If mblnTmp Then Exit Sub
    
    Call LoadDoc
End Sub

Private Sub cmdFind_Click()
    Call LoadAdvice(IIf(tbcSub.Selected.Tag = "�����", True, False))
End Sub

Private Sub GetLocalSetting()
'��ȡ���ز���
    cboTime.ListIndex = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", 0)
    mblnTmp = True
    chkIsShowAll.Value = Val(zlDatabase.GetPara("ֻ��ʾ�������ҽ��", glngSys, mlngModul, 0) & "")
    mblnTmp = False
End Sub

Private Sub cmdFindY_Click()
    Call LoadAdvice(IIf(tbcSub.Selected.Tag = "�����", True, False))
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mblnTeam = gblnKSSAuditType
    
    If Not CheckKssJuris Then Unload Me: Exit Sub
    
    mstrPrivs = GetInsidePrivs(p������ҩ���)
    mlngModul = p������ҩ���
    mlngCodeType = zlDatabase.GetPara("���뷽ʽ")
    mblnIsUpdate = False
    mstrChangeRows = ""
    mstrǩ��IDs = "0"
    
    '---cboTime
    cboTime.AddItem "��    ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ  ��]"
    cboTime.ListIndex = 0
    
    '---cboDateY
    cboDateY.AddItem "��    ��"
    cboDateY.AddItem "�������"
    cboDateY.AddItem "�������"
    cboDateY.AddItem "���һ��"
    cboDateY.AddItem "���һ��"
    cboDateY.AddItem "[ָ  ��]"
    cboDateY.ListIndex = 3
    
    'tabControl
    '-----------------------------------------------------
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
        .InsertItem(0, "  �����  ", picUnAudited.hwnd, 0).Tag = "�����"
        .InsertItem(1, "  �����  ", picUnAudited.hwnd, 0).Tag = "�����"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
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
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    Call MainDefCommandBar
    
    'vsFlexGrid
    '-----------------------------------------------------
    strHead = ",450,1; ;���˵��,2560,1;���ʱ��;��������,1000,1;ҽ������,3500,1;��ҩĿ��,1050,1;��ҩ����,2500,1;��Ч,500,1;����;����;Ƶ��,1350,1;��ҩ;��,1350,1;��ʼʱ��,2000,1;ִ����ֹʱ��,2000,1;ִ��ʱ�䷽��,1350,1;ҽ��ID;���ID;�Ա�;����;����;��Ժʱ��;����; ���; ����;����ȼ�;����ID; ��ҳID;�Һŵ�; ��ID;������� ;������Դ;�Һŵ���;ǩ��id;ҽ��״̬"
    strHead = strHead & ";�����;סԺ��;��ǰ����;����ҽ��;����ʱ��;��������id;��Ժ����id;��ǰ����id"
    Call Grid.Init(vsAudit, strHead)
    vsAudit.ExtendLastCol = True
    vsAudit.Editable = flexEDKbdMouse
    vsAudit.Cell(flexcpPicture, 0, col_ѡ��) = img16.ListImages("unCheck").Picture
    vsAudit.Cell(flexcpPictureAlignment, 0, col_ѡ��) = flexPicAlignCenterCenter
    vsAudit.ColDataType(COL_ȡ��ѡ��) = flexDTBoolean
    vsAudit.Cell(flexcpPicture, 0, COL_ȡ��ѡ��) = img16.ListImages("unCheck").Picture
    vsAudit.Cell(flexcpPictureAlignment, 0, COL_ȡ��ѡ��) = flexPicAlignCenterCenter
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mrsDefine = InitAdviceDefine
    Call GetLocalSetting '���ز���
    
    Call LoadDept
End Sub

Private Sub LoadDept()
'���ز���Ա��������
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    strSQL = "Select B.ID,B.����,B.���� " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", ",A.ȱʡ") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", "������Ա A, ") & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", " And a.����id = B.Id And A.��ԱID = [1] ") & vbNewLine & _
            "  And C.�������� = '�ٴ�' And Instr([2],C.������� || '')>0 And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"

    On Error GoTo errH
    cboDept.Clear
    '���в���
    If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 Then
        cboDept.AddItem "���в���"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, IIf(optOccasion(0).Value, "2,3", "1,3"))
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '����ȱʡ
        If InStr(";" & mstrPrivs & ";", ";���в���;") = 0 Then
            If rsTmp!ȱʡ = 1 Then
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long
    Dim strName As String
    
    If mblnTeam Then
        strName = "����С��"
    Else
        strName = "��������"
    End If
    
    With rptDoc
        Set objCol = .Columns.Add(COL_��ԱID, "��ԱID", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 70, True)
        Set objCol = .Columns.Add(COL_רҵְ��, "רҵְ��", 70, True)
        Set objCol = .Columns.Add(COL_������ҩȨ��, "������ҩȨ��", 80, True)
        Set objCol = .Columns.Add(COL_ƴ������, "ƴ������", 0, False)
        Set objCol = .Columns.Add(COL_��ʼ���, "��ʼ���", 0, False)
        Set objCol = .Columns.Add(COL_��������, strName, 0, False)
        Set objCol = .Columns.Add(COL_��������ID, "��������ID", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ��ҽ��..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 Then .GroupsOrder.Add .Columns(COL_��������)
    End With
End Sub


Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����(&U)")
        objControl.BeginGroup = True
        objControl.IconId = 21905
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("������", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����(&U)")
            objControl.BeginGroup = True
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        Set objCustom = .Add(xtpControlCustom, conMenu_View_FindType, "����")
            objCustom.Handle = fraType.hwnd
            objCustom.Flags = xtpFlagRightAlign
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    '�Ҽ�������˲˵�
    Set mobjPopup = cbsMain.Add("�Ҽ��˵�", xtpBarPopup)
    With mobjPopup.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
    End With
    
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsUpdate = True Then
        If MsgBox("��ǰ���������δ���棬�Ƿ�Ҫ�˳���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    If Not mobjESign Is Nothing Then Set mobjESign = Nothing
    mlngFindNum = 0
    Set mclsMipModule = Nothing
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", cboTime.ListIndex
    zlDatabase.SetPara "ֻ��ʾ�������ҽ��", chkIsShowAll.Value & "", glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
End Sub

Private Sub optOccasion_Click(Index As Integer)
    If mblnTmp Then Exit Sub
    If CheckKssJuris = False Then
        '���û��Ȩ�ޣ������л���ȥ
        mblnTmp = True
        optOccasion(IIf(optOccasion(0).Value, 1, 0)).Value = True
        mblnTmp = False
    End If
    Call LoadDept
    vsAudit.Rows = 1
    vsAudit.AddItem ""
End Sub

Private Sub picUnAudited_Resize()
    On Error Resume Next
    picDate.Move 0, 0, picUnAudited.Width
    picDateY.Move 0, 0, picUnAudited.Width
    vsAudit.Move 0, picDate.Top + picDate.Height, picUnAudited.Width, picUnAudited.Height - picDate.Top - picDate.Height
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '����ҽ���б�
    If tbcSub.Selected.Tag = "�����" Then
        If Me.Visible Then Call LoadAdvice
    Else
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Tag = "�����" Then
        picDate.Visible = True
        picDateY.Visible = False
        Call picUnAudited_Resize
        vsAudit.ColWidth(COL_ȡ��ѡ��) = 250
        vsAudit.ColHidden(COL_ȡ��ѡ��) = False
        vsAudit.ColWidth(COL_���ʱ��) = 1800
        vsAudit.ColHidden(COL_���ʱ��) = False
        Set vsAudit.Cell(flexcpPicture, 0, col_ѡ��) = Nothing
        vsAudit.TextMatrix(0, col_ѡ��) = "״̬"
        If Me.Visible Then Call LoadAdvice(True)
    Else
        picDate.Visible = False
        picDateY.Visible = True
        Call picUnAudited_Resize
        vsAudit.Cell(flexcpPicture, 0, col_ѡ��) = img16.ListImages("unCheck").Picture
        vsAudit.TextMatrix(0, col_ѡ��) = ""
        vsAudit.ColWidth(COL_ȡ��ѡ��) = 0
        vsAudit.ColHidden(COL_ȡ��ѡ��) = True
        vsAudit.ColWidth(COL_���ʱ��) = 0
        vsAudit.ColHidden(COL_���ʱ��) = True
        If Me.Visible Then Call LoadAdvice
    End If
End Sub

Private Sub txtFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub LoadAdvice(Optional ByVal blnIsAudited As Boolean)
'���ش���˺�����˵�ҽ��
'�������Ƿ���������ҽ��,Ϊ��Ϊ���ش����ҽ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '���ڶ�λ
    Dim strFormat As String
    Dim strTmp As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    stbThis.Panels(2).Text = ""
    If rptDoc.SelectedRows.Count = 0 Then
        stbThis.Panels(2).Text = "��ѡ��һλ����ҽ����"
        Exit Sub
    End If
    If rptDoc.SelectedRows(0).GroupRow Then
        vsAudit.Rows = 1
        vsAudit.AddItem ""
        stbThis.Panels(2).Text = "��ѡ��һλ����ҽ����"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    strSQL = "Select a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, Nvl(x.���, a.���) As ���, a.�������, Null As ѡ��, Null As ����, A.����, p.��ǰ���� As ����," & vbNewLine & _
            "       Decode(Nvl(a.ҽ����Ч, 0), 0, '����', '����') As ��Ч, To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI') As ��ʼʱ��, To_Char(a.ִ����ֹʱ��, 'YYYY-MM-DD HH24:MI') As ��ֹʱ��, a.ҽ������," & vbNewLine & _
            "       Decode(a.�ܸ�����, Null, Null,Round(a.�ܸ����� / Decode(a.������Դ, 2, d.סԺ��װ, d.�����װ), 5) || Decode(a.������Դ, 2, d.סԺ��λ, d.���ﵥλ)) As ����," & vbNewLine & _
            "       Decode(a.��������, Null, Null, a.�������� || b.���㵥λ) As ����, a.ִ��Ƶ�� As Ƶ��, x.ҽ������ As �÷�, a.ִ��ʱ�䷽�� As ִ��ʱ�䷽��, a.����id," & vbNewLine & _
            "       a.��ҳid, g.ID as �Һŵ�, a.������Ŀid, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, b.���㵥λ As ������λ, h.������, e.����,e.��Ժ����,e.��Ժ����,A.����,A.�Ա�,f.���� as ����ȼ�,a.������Դ,G.NO as �Һŵ���,A.��ҩĿ��,A.��ҩ����" & vbNewLine & _
            IIf(blnIsAudited, ", c.��������, c.����˵��, c.ǩ��id ,a.ҽ��״̬,c.����ʱ�� as ���ʱ��", "") & _
            ",p.�����,p.סԺ��,p.��ǰ����,a.����ҽ��,To_Char(a.����ʱ��,'YYYY-MM-DD HH24:MI') As ����ʱ��,a.��������id,e.��Ժ����id,e.��ǰ����id" & _
            " From ����ҽ����¼ A, ������Ϣ P, ҩƷ��� D, ������ĿĿ¼ B, ����ҽ����¼ X, ҩƷ���� H, ������ҳ E,�շ���ĿĿ¼ F,���˹Һż�¼ G" & vbNewLine & _
            IIf(blnIsAudited, ", (Select ҽ��id,����ʱ��,����˵��,��������,ǩ��ID" & vbNewLine & _
                            "From (Select C.ҽ��id,C.����ʱ��,C.����˵��,C.��������,C.ǩ��ID, Row_Number() Over(Partition By C.ҽ��id Order By C.����ʱ�� Desc) Top" & vbNewLine & _
                            "       From ����ҽ��״̬ C" & vbNewLine & _
                            "       Where c.����ʱ�� Between [3] And [4] " & vbNewLine & _
                            "       and C.�������� in(11,12) And C.������Ա =[2])" & vbNewLine & _
                            "Where Top = 1)  C", "") & _
            " Where a.����id = p.����id And a.������Ŀid = b.Id And a.�շ�ϸĿid = d.ҩƷid(+) And a.���id = x.Id And d.ҩ��id = h.ҩ��id(+) And f.id(+)=e.����ȼ�id And g.no(+)=a.�Һŵ� And" & vbNewLine & _
            "      e.����id(+) = a.����id And e.��ҳid(+) = a.��ҳid " & _
            IIf(blnIsAudited, " And c.ҽ��id = a.Id ", " And a.ҽ��״̬<>4 and (a.ҽ��״̬=1 or a.ҽ��״̬>2 and a.������־=1) And a.���״̬ = 1 ") & vbNewLine & _
            IIf(blnIsAudited, "", "  And A.����ʱ�� between [6] and [7] And Not Exists (Select 1 From ����ҽ��״̬ I Where i.ҽ��id = a.Id And i.��������=11 And i.������Ա = [2]) ") & _
            "  And a.����ҽ��=[1] And A.������Դ=[5] And a.������� In ('5', '6') Order By p.����,To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI'),Nvl(a.���id, a.Id),a.id"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rptDoc.SelectedRows(0).Record(col_����).Value, UserInfo.����, CDate(dtpTime(0).Value), CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60), IIf(optOccasion(0).Value, 2, 1), CDate(dtpTimeY(0).Value), CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60))
    
    With vsAudit
        If Val(.TextMatrix(.Row, col_ҽ��ID)) <> 0 Then lngID = Val(.TextMatrix(.Row, col_ҽ��ID))
        If Not blnIsAudited Then .Cell(flexcpPicture, 0, col_ѡ��) = img16.ListImages("unCheck").Picture
        .Cell(flexcpPicture, 0, COL_ȡ��ѡ��) = img16.ListImages("unCheck").Picture
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(i, COL_��������) = rsTmp!���� & ""
                .TextMatrix(i, col_��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, COL_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_Ƶ��) = rsTmp!Ƶ�� & ""
                .TextMatrix(i, col_��ҩ;��) = rsTmp!�÷� & ""
                .TextMatrix(i, col_��ҩĿ��) = Decode(rsTmp!��ҩĿ�� & "", "1", "Ԥ��", "2", "����", "3", "Ԥ��������", "")
                .TextMatrix(i, col_��ҩ����) = rsTmp!��ҩ���� & ""
                .TextMatrix(i, COL_��ʼʱ��) = rsTmp!��ʼʱ�� & ""
                .TextMatrix(i, COL_��ֹʱ��) = rsTmp!��ֹʱ�� & ""
                .TextMatrix(i, col_ִ��ʱ�䷽��) = rsTmp!ִ��ʱ�䷽�� & ""
                .TextMatrix(i, col_ҽ��ID) = rsTmp!ID & ""
                If Val(rsTmp!ID & "") = lngID And lngID <> 0 Then
                    .Row = i
                End If
                .TextMatrix(i, col_���ID) = rsTmp!���ID & ""
                .TextMatrix(i, col_�Ա�) = rsTmp!�Ա� & ""
                .TextMatrix(i, col_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_��Ժʱ��) = rsTmp!��Ժ���� & ""
                .TextMatrix(i, col_����) = rsTmp!���� & ""
                .TextMatrix(i, col_������) = rsTmp!������ & ""
                .TextMatrix(i, col_����ȼ�) = rsTmp!����ȼ� & ""
                .TextMatrix(i, col_����Id) = rsTmp!����ID & ""
                .TextMatrix(i, col_��ҳID) = rsTmp!��ҳID & ""
                .TextMatrix(i, col_�Һŵ�) = rsTmp!�Һŵ� & ""
                .TextMatrix(i, COL_��ID) = rsTmp!��ID & ""
                .TextMatrix(i, COL_�������) = rsTmp!������� & ""
                .TextMatrix(i, COL_������Դ) = rsTmp!������Դ & ""
                .TextMatrix(i, col_�Һŵ���) = rsTmp!�Һŵ��� & ""
                .TextMatrix(i, COL_����) = rsTmp!��Ժ���� & ""
                '��ʾ���ģʽ�µ�ҽ������
                strFormat = rsTmp!ҽ������
                If .TextMatrix(i, COL_Ƶ��) <> "һ����" Then
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, COL_����)
                        If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                    End If
                    
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, COL_����)
                        If strTmp <> "" Then strFormat = strFormat & ",ÿ��" & strTmp
                    End If
                End If
                .TextMatrix(i, COL_ҽ������) = strFormat
                If blnIsAudited Then
                    .TextMatrix(i, COL_ǩ��id) = rsTmp!ǩ��id & ""
                    .TextMatrix(i, COL_ҽ��״̬) = rsTmp!ҽ��״̬ & ""
                    .Cell(flexcpData, i, col_ѡ��) = Val(rsTmp!�������� & "") - 10
                    .Cell(flexcpPicture, i, col_ѡ��) = imgAdvice.ListImages(Val(.Cell(flexcpData, i, col_ѡ��))).Picture
                    .Cell(flexcpPictureAlignment, i, col_ѡ��) = flexPicAlignCenterCenter
                    .TextMatrix(i, COL_���˵��) = rsTmp!����˵�� & ""
                    .TextMatrix(i, COL_���ʱ��) = Format(rsTmp!���ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                    '���ҽ���������¿�״̬����ı�������ɫ
                    If Val(rsTmp!ҽ��״̬ & "") <> 1 Then
                        .Cell(flexcpForeColor, i, col_ѡ��, i, COL_ǩ��id) = &HC00000
                    End If
                End If
                .TextMatrix(i, COL_�����) = rsTmp!����� & ""
                .TextMatrix(i, col_סԺ��) = rsTmp!סԺ�� & ""
                .TextMatrix(i, COL_��ǰ����) = rsTmp!��ǰ���� & ""
                .TextMatrix(i, COL_����ҽ��) = rsTmp!����ҽ�� & ""
                .TextMatrix(i, COL_����ʱ��) = rsTmp!����ʱ�� & ""
                .TextMatrix(i, COL_��������ID) = rsTmp!��������ID & ""
                .TextMatrix(i, COL_��Ժ����ID) = rsTmp!��Ժ����ID & ""
                .TextMatrix(i, COL_��ǰ����ID) = rsTmp!��ǰ����ID & ""
                rsTmp.MoveNext
                i = i + 1
            Loop
            .Cell(flexcpBackColor, 1, IIf(blnIsAudited, 1, 0), i - 1, COL_���˵��) = &HFAEADA
            If blnIsAudited Then
                For j = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, j, col_ѡ��)) = 1 Or (.TextMatrix(j, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(j, COL_ҽ��״̬) & "" <> "") Then
                        .Cell(flexcpBackColor, j, COL_���˵��) = &H80000005
                        If .TextMatrix(j, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(j, COL_ҽ��״̬) & "" <> "" Then
                            '��У�Ե�ҽ���������޸Ļ����
                            .Cell(flexcpBackColor, j, COL_ȡ��ѡ��) = &H80000005
                        End If
                    End If
                Next
            End If
        Else
            .AddItem ""
            .Cell(flexcpBackColor, 1, IIf(blnIsAudited, 1, 0), 1, COL_���˵��) = &HFAEADA
        End If
        
        strFormat = "������ҽ����" & rptDoc.SelectedRows(0).Record(col_����).Value & "��"
        If blnIsAudited Then
            strTmp = "�ڡ����ʱ�䣺" & Format(dtpTime(0).Value, "YYYY-MM-DD") & " 00:00:00 - " & Format(dtpTime(1).Value, "YYYY-MM-DD") & " 23:59:59���ڣ�"
            If Val(.TextMatrix(1, col_ҽ��ID)) = 0 Then
                strTmp = strTmp & strFormat & "�����ڱ���˹���ҽ����"
            Else
                strTmp = strTmp & strFormat & "����" & (.Rows - 1) & "��ҽ������ˡ�"
            End If
        Else
            Call CheckIsExceed
            strTmp = "�ڡ�����ʱ�䣺" & Format(dtpTimeY(0).Value, "YYYY-MM-DD") & " 00:00:00 - " & Format(dtpTimeY(1).Value, "YYYY-MM-DD") & " 23:59:59���ڣ�"
            If Val(.TextMatrix(1, col_ҽ��ID)) = 0 Then
                strTmp = strTmp & strFormat & "��������Ҫ��˵�ҽ����"
            Else
                strTmp = strTmp & strFormat & "����" & (.Rows - 1) & "��ҽ����Ҫ��ˡ�"
            End If
        End If
        stbThis.Panels(2).Text = strTmp
        
        '�Զ������и�
        .AutoSize COL_ҽ������
        .Redraw = flexRDDirect
        If .Row > 0 Then Call vsAudit_AfterRowColChange(1, 1, .Row, COL_���˵��)
    End With
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CheckIsExceed()
'���ܣ����ͬ��ҩ�����еȼ��ȵ�ǰ����Ա��Ȩ�޸���ʱ����ɾ������
    Dim i As Long, j As Long
    Dim strTmp As String     '��Ҫɾ������
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAudit
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, col_������) & "") > mlngLevel Then
                If Not RowInһ����ҩ(i, lngBegin, lngEnd, vsAudit) Then
                    strTmp = strTmp & IIf(strTmp = "", "", ",") & i
                Else
                    For j = lngEnd To lngBegin Step -1
                        If .TextMatrix(j, COL_��ID) = .TextMatrix(i, COL_��ID) Then
                            strTmp = strTmp & IIf(strTmp = "", "", ",") & j
                        End If
                    Next
                    i = lngBegin
                End If
            End If
        Next
        'ɾ�����Ӻ�ɾ��
        If strTmp = "" Then Exit Sub
        For i = 0 To UBound(Split(strTmp, ","))
            .RemoveItem Val(Split(strTmp, ",")(i) & "")
        Next
        If .Rows = 1 Then .AddItem ""
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If ZLCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(col_����).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(mlngCodeType = 0, COL_ƴ������, COL_��ʼ���)).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(col_����).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "��ǰ����û���ҵ������ҵ�ҽ����", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "�Ѿ������һ��ҽ���ˡ�", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub

Private Sub vsAudit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    With vsAudit
        If Not Visible Then Exit Sub
        If NewCol = COL_���˵�� And tbcSub.Selected.Tag = "�����" Or NewCol = col_ѡ�� Or NewCol = COL_ȡ��ѡ�� Then
            If (Val(vsAudit.Cell(flexcpData, NewRow, col_ѡ��) & "") = "1" And NewCol = COL_���˵��) Or _
                    (vsAudit.TextMatrix(NewRow, COL_ҽ��״̬) & "" <> "1" And vsAudit.TextMatrix(NewRow, COL_ҽ��״̬) & "" <> "" And NewCol = COL_���˵��) _
                    Or (tbcSub.Selected.Tag = "�����" And NewCol = col_ѡ��) Then
                vsAudit.FocusRect = flexFocusNone
            Else
                If .TextMatrix(NewRow, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(NewRow, COL_ҽ��״̬) & "" <> "" Then
                    vsAudit.FocusRect = flexFocusNone
                Else
                    vsAudit.FocusRect = flexFocusHeavy
                End If
            End If
        Else
            vsAudit.FocusRect = flexFocusNone
        End If
        
        '��ɫ
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)

        If vsAudit.TextMatrix(NewRow, col_ҽ��ID) <> "" And NewRow <> 0 Then
            lblInformation(info_��Ժʱ��).Caption = Format(.TextMatrix(NewRow, COL_��Ժʱ��), "yyyy-MM-dd")
            lblInformation(info_�Ա�).Caption = .TextMatrix(NewRow, col_�Ա�)
            lblInformation(info_����).Caption = .TextMatrix(NewRow, col_����)
            lblInformation(info_����).Caption = .TextMatrix(NewRow, COL_����)
            lblInformation(info_����).Caption = .TextMatrix(NewRow, col_����)
            lblInformation(info_����ȼ�).Caption = .TextMatrix(NewRow, col_����ȼ�)
            lblInformation(info_����).Caption = IIf(Val(.TextMatrix(NewRow, COL_����) & "") = 0, "", .TextMatrix(NewRow, COL_����) & "Kg")
            
            '������¼
            Call LoadPatiAllergy(Val(.TextMatrix(NewRow, col_����Id) & ""), cbo����)
            
            '���
            lblInformation(info_���).Caption = GetPatiDiagnose(Val(.TextMatrix(NewRow, col_����Id) & ""), _
            IIf(.TextMatrix(NewRow, COL_������Դ) = "1", Val(.TextMatrix(NewRow, col_�Һŵ�) & ""), Val(.TextMatrix(NewRow, col_��ҳID) & "")), _
            Val(.TextMatrix(NewRow, COL_������Դ)))
            'סԺ��Ϣ��ʾ
            picInShow.Visible = Not .TextMatrix(NewRow, COL_������Դ) = "1"
        Else
            lblInformation(info_��Ժʱ��).Caption = ""
            lblInformation(info_�Ա�).Caption = ""
            lblInformation(info_����).Caption = ""
            lblInformation(info_����).Caption = ""
            lblInformation(info_����).Caption = ""
            lblInformation(info_����ȼ�).Caption = ""
            lblInformation(info_����).Caption = ""
            
            '������¼
            cbo����.Clear
            
            '���
            lblInformation(info_���).Caption = ""
            
            picInShow.Visible = True
        End If
    End With
End Sub

Private Sub vsAudit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = COL_���˵�� And tbcSub.Selected.Tag = "�����") Then
        Cancel = True
    Else
        If Val(vsAudit.Cell(flexcpData, Row, col_ѡ��) & "") = "1" Or vsAudit.TextMatrix(1, col_ҽ��ID) & "" = "" Or _
                (vsAudit.TextMatrix(Row, COL_ҽ��״̬) & "" <> "1" And vsAudit.TextMatrix(Row, COL_ҽ��״̬) & "" <> "") Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAudit_Click()
    Dim i As Long
    
    With vsAudit
        If tbcSub.Selected.Tag = "�����" Then
            If .MouseCol = COL_ȡ��ѡ�� And .MouseRow = .FixedRows - 1 Then
                If .TextMatrix(1, col_ҽ��ID) = "" Then Exit Sub
                If .ColData(COL_ȡ��ѡ��) = "Check" Then
                    .Cell(flexcpPicture, 0, COL_ȡ��ѡ��) = img16.ListImages("unCheck").Picture
                    .ColData(COL_ȡ��ѡ��) = ""
                Else
                    .Cell(flexcpPicture, 0, COL_ȡ��ѡ��) = img16.ListImages("Check").Picture
                    .ColData(COL_ȡ��ѡ��) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_ҽ��ID) = "" Then Exit For
                    If .ColData(COL_ȡ��ѡ��) = "Check" Then
                        If Not (.TextMatrix(i, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(i, COL_ҽ��״̬) & "" <> "") Then
                            .TextMatrix(i, COL_ȡ��ѡ��) = -1
                        End If
                    Else
                        .TextMatrix(i, COL_ȡ��ѡ��) = 0
                    End If
                    
                Next
            ElseIf .MouseCol = COL_ȡ��ѡ�� And .MouseRow > .FixedRows - 1 And .MouseRow < .Rows Then
                 Call vsAudit_KeyPress(vbKeySpace)
            End If
        Else
            If .MouseCol = col_ѡ�� And .MouseRow = .FixedRows - 1 Then
                If .TextMatrix(1, col_ҽ��ID) = "" Then Exit Sub
                For i = 1 To .Rows - 1
                    If .ColData(col_ѡ��) = "" Then
                        If .TextMatrix(i, COL_���˵��) <> "" Then
                            If MsgBox("���Ѿ���д�����˵�����޸�Ϊͨ����ɾ��˵�����Ƿ������", vbQuestion + vbDefaultButton1 + vbYesNo, Me.Caption) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                If .ColData(col_ѡ��) = "Check" Then
                    .Cell(flexcpPicture, 0, col_ѡ��) = img16.ListImages("unCheck").Picture
                    .ColData(col_ѡ��) = ""
                Else
                    .Cell(flexcpPicture, 0, col_ѡ��) = img16.ListImages("Check").Picture
                    .ColData(col_ѡ��) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_ҽ��ID) = "" Then Exit For
                    If .ColData(col_ѡ��) = "Check" Then
                        .Cell(flexcpPicture, i, col_ѡ��) = imgAdvice.ListImages(1).Picture
                        .Cell(flexcpData, i, col_ѡ��) = 1
                        .Cell(flexcpPictureAlignment, i, col_ѡ��) = flexPicAlignCenterCenter
                        vsAudit.Cell(flexcpBackColor, i, COL_���˵��) = &H80000005
                        .TextMatrix(i, COL_���˵��) = ""
                    Else
                        Set .Cell(flexcpPicture, i, col_ѡ��) = Nothing
                        .Cell(flexcpData, i, col_ѡ��) = 0
                        vsAudit.Cell(flexcpBackColor, i, COL_���˵��) = &HFAEADA
                    End If
                    
                Next
                mblnIsUpdate = True
            End If
        End If
    End With
End Sub

Private Sub vsAudit_DblClick()
    With vsAudit
        If .MouseCol = col_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAudit_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsAudit_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAudit
        lngLeft = col_��Ч: lngRight = col_��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Ƶ��: lngRight = col_ִ��ʱ�䷽��
        End If
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_��������: lngRight = COL_��������
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsAudit) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col_���ID)) = Val(.TextMatrix(lngRow, col_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col_���ID)) = Val(.TextMatrix(lngRow, col_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col_���ID)) = Val(.TextMatrix(lngRow, col_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col_���ID)) = Val(.TextMatrix(lngRow, col_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function RowInͬһ����(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�������Ƿ�������ҽ��
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If lngRow = 0 Then Exit Function
        If .TextMatrix(lngRow - 1, COL_��������) = .TextMatrix(lngRow, COL_��������) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, COL_��������) = .TextMatrix(lngRow, COL_��������) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_��������) = .TextMatrix(lngRow, COL_��������) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, COL_��������) = .TextMatrix(lngRow, COL_��������) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInͬһ���� = blnTmp
    End With
End Function

Private Sub vsAudit_KeyPress(KeyAscii As Integer)
    With vsAudit
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call UnAuditEnterNextCell
        ElseIf .Col = COL_���˵�� Then
            .ComboList = "" 'ʹ��ť״̬��������״̬
        ElseIf .Col = col_ѡ�� And KeyAscii = vbKeySpace Then
            Call AuditStateCheck
        ElseIf .Col = COL_ȡ��ѡ�� And KeyAscii = vbKeySpace Then
            Call AuditCancleCheck
        End If
    End With
End Sub

Private Sub AuditCancleCheck()
'���ܣ������ȡ��ѡ���ͬ��ѡ��һ��ҩƷ
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim lngCheck As Long
    Dim blnIsAudit As Boolean   '�ж�ҽ�����¿�״̬
    
    With vsAudit
        If tbcSub.Selected.Tag = "�����" Then Exit Sub
        If .TextMatrix(.Row, col_ҽ��ID) = "" Or (.TextMatrix(.Row, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(.Row, COL_ҽ��״̬) & "" <> "") Then Exit Sub
        '�����������ǩ�������������Ƿ���һ��ǩ���ģ�һ��ѡ��
        If gintCA = 0 Or (.TextMatrix(.Row, COL_ǩ��id) = "" And gintCA > 0) Then
            If Not RowInһ����ҩ(.Row, lngBegin, lngEnd, vsAudit) Then
                lngBegin = .Row: lngEnd = .Row
            End If
        Else
            If Not RowInͬһ����(.Row, lngBegin, lngEnd, vsAudit) Then
                lngBegin = .Row: lngEnd = .Row
            End If
        End If
        lngCheck = Val(.TextMatrix(lngBegin, COL_ȡ��ѡ��))
        For i = lngBegin To lngEnd
            If gintCA = 0 Or (.TextMatrix(.Row, COL_ǩ��id) = "" And gintCA > 0) Then
                .TextMatrix(i, COL_ȡ��ѡ��) = IIf(lngCheck = 0, -1, 0)
            Else
                If .TextMatrix(i, COL_ǩ��id) <> "" And .TextMatrix(i, COL_ǩ��id) = .TextMatrix(.Row, COL_ǩ��id) Then
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                        .TextMatrix(i, COL_ȡ��ѡ��) = IIf(lngCheck = 0, -1, 0)
                    Else
                        blnIsAudit = True
                        Exit For
                    End If
                End If
                If i = lngEnd Then stbThis.Panels(2).Text = "һͬ��ѡ/ȡ����ҽ��Ϊ����ǩ����˵ġ�"
            End If
        Next
        '������в���������ҽ����ȡ��ѡ�񣬲���ʾ
        If blnIsAudit Then
            For i = lngBegin To lngEnd
                If .TextMatrix(i, COL_ǩ��id) <> "" And .TextMatrix(i, COL_ǩ��id) = .TextMatrix(.Row, COL_ǩ��id) Then
                    .TextMatrix(i, COL_ȡ��ѡ��) = 0
                End If
            Next
            MsgBox "�������������ǩ����ҽ���Ѿ�У�ԣ�����ȡ����ˡ�", vbInformation, Me.Caption
        End If
    End With
End Sub

Private Sub AuditStateCheck(Optional ByVal lngState As Long)
'ͬ��ѡ��һ��ҩƷ
'������lngState=0����null Ϊ������һ��״̬��1=�� ��2=����3=�����
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With vsAudit
        If tbcSub.Selected.Tag = "�����" Then Exit Sub
        If .TextMatrix(.Row, col_ҽ��ID) = "" Or (.TextMatrix(.Row, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(.Row, COL_ҽ��״̬) & "" <> "") Then Exit Sub
        If Not RowInһ����ҩ(.Row, lngBegin, lngEnd, vsAudit) Then
            lngBegin = .Row: lngEnd = .Row
        End If
        
        For i = lngBegin To lngEnd
            If lngState = 1 Or Val(.Cell(flexcpData, i, col_ѡ��) & "") = 0 Then
                If .TextMatrix(i, COL_���˵��) <> "" Then
                    If MsgBox("���Ѿ���д�����˵�����޸�Ϊͨ����ɾ��˵�����Ƿ������", vbQuestion + vbDefaultButton1 + vbYesNo, Me.Caption) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        For i = lngBegin To lngEnd
            If lngState = 1 Or Val(.Cell(flexcpData, i, col_ѡ��) & "") = 0 Then
                .TextMatrix(i, COL_���˵��) = ""
            End If
            .Cell(flexcpData, i, col_ѡ��) = IIf(lngState = 0, Val(.Cell(flexcpData, i, col_ѡ��) & "") + IIf(Val(.Cell(flexcpData, i, col_ѡ��) & "") = 2, -2, 1), IIf(lngState = 3, 0, lngState))
            If Val(.Cell(flexcpData, i, col_ѡ��) & "") = 0 Then
                Set .Cell(flexcpPicture, i, col_ѡ��) = Nothing
            Else
                .Cell(flexcpPicture, i, col_ѡ��) = imgAdvice.ListImages(Val(.Cell(flexcpData, i, col_ѡ��) & "")).Picture
            End If
            .Cell(flexcpPictureAlignment, i, col_ѡ��) = flexPicAlignCenterCenter
            vsAudit.Cell(flexcpBackColor, i, COL_���˵��) = IIf(Val(.Cell(flexcpData, i, col_ѡ��) & "") = 1, &H80000005, &HFAEADA)
        Next
        mblnIsUpdate = True
    End With
End Sub


Private Sub vsAudit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = COL_���˵�� Then
        If ZLCommFun.ActualLen(vsAudit.Editable) - ZLCommFun.ActualLen(vsAudit.EditSelText) >= 100 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
            If KeyAscii = vbKeyReturn Then
                Call UnAuditEnterNextCell
                Exit Sub
            End If
            KeyAscii = 0
        ElseIf Chr(KeyAscii) = "'" Then
            KeyAscii = 0
        End If
        
    End If
End Sub

Private Sub UnAuditEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAudit
        If .Col = COL_���˵�� Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call ZLCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAudit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If vsAudit.Rows <= 1 Then Exit Sub
    If vsAudit.TextMatrix(1, col_ҽ��ID) <> "" And (vsAudit.MouseCol = col_ѡ�� Or vsAudit.MouseCol = COL_���˵��) And vsAudit.MouseRow = 0 And tbcSub.Selected.Tag = "�����" Then
        strTip = "ѡ�е�һ�еĵ�Ԫ�񰴿ո��˫���ɸı���˽����" & vbCrLf & "��Ϊ��ͨ������Ϊͨ����"
        ZLCommFun.ShowTipInfo vsAudit.hwnd, strTip, True
    Else
        strTip = ""
        ZLCommFun.ShowTipInfo 0, strTip, True
    End If
End Sub

Private Sub vsAudit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsAudit.MouseRow <= 0 Then Exit Sub
    If Button = 2 And vsAudit.TextMatrix(vsAudit.MouseRow, col_ҽ��ID) <> "" Then
        If tbcSub.Selected.Tag = "�����" Then
            Call vsAudit.Select(vsAudit.MouseRow, vsAudit.FixedCols, vsAudit.MouseRow, vsAudit.Cols - 1)
            mobjPopup.ShowPopup
        Else
            Call vsAudit.Select(vsAudit.MouseRow, vsAudit.FixedCols, vsAudit.MouseRow, vsAudit.Cols - 1)
            mobjPopup.ShowPopup
        End If
    End If
End Sub

Private Sub vsAudit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = COL_���˵�� Then
        vsAudit.EditSelStart = 0
        vsAudit.EditSelLength = Len(vsAudit.EditText)
    End If
End Sub

Private Sub vsAudit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_���˵�� Then
        If vsAudit.EditText <> vsAudit.TextMatrix(Row, Col) Then
            If Val(vsAudit.Cell(flexcpData, Row, col_ѡ��) & "") = "0" And tbcSub.Selected.Tag = "�����" Then
                Call AuditStateCheck(2)
            End If
            mblnIsUpdate = True
            If tbcSub.Selected.Tag = "�����" Then
                mstrChangeRows = mstrChangeRows & IIf(mstrChangeRows = "", "", ",") & Row
            End If
        End If
    End If
End Sub
