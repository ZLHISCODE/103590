VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExamineTransfuse 
   Caption         =   "��Ѫ��˹���"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14805
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   14805
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   12240
      TabIndex        =   38
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "סԺ"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   40
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         Top             =   -10
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "ʹ�ó���"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   60
         Width           =   4000
      End
   End
   Begin VB.Frame fraPati 
      Caption         =   "������Ϣ"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   16
      Top             =   600
      Width           =   11295
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   28
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
         TabIndex        =   17
         Top             =   360
         Width           =   8450
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   18
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   19
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblCaption 
            Caption         =   "���ţ�"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   26
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "����ȼ���"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   25
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblCaption 
            Caption         =   "������"
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   24
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "���أ�"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "��Ժʱ�䣺"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   27
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   29
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblCaption 
         Caption         =   "��ϣ�"
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "����ҩ�"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         Caption         =   "���䣺"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "�Ա�"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
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
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9615
         TabIndex        =   42
         Top             =   120
         Width           =   9615
         Begin VB.CommandButton cmdFindY 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   44
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboDateY 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   30
            Width           =   1365
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
            Format          =   179765251
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
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin VB.Label lblPri 
            Height          =   300
            Index           =   0
            Left            =   7080
            TabIndex        =   49
            Top             =   120
            Width           =   8000
         End
         Begin VB.Label lblDateY 
            Caption         =   "����ʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   75
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2460
            TabIndex        =   47
            Top             =   90
            Width           =   1890
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
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   9375
         Begin VB.CommandButton cmdFind 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   30
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   11
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   12
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   179765251
            CurrentDate     =   40256
         End
         Begin VB.Label lblPri 
            Height          =   300
            Index           =   1
            Left            =   7200
            TabIndex        =   52
            Top             =   120
            Width           =   7995
         End
         Begin VB.Label lblDate 
            Caption         =   "ǩ��ʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2460
            TabIndex        =   13
            Top             =   90
            Width           =   1890
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAudit 
         Height          =   4860
         Left            =   240
         TabIndex        =   15
         Top             =   720
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   41
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.Frame fraDoctor 
      Caption         =   "ҽ��"
      ForeColor       =   &H000040C0&
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   4020
         Left            =   105
         TabIndex        =   1
         Top             =   1500
         Width           =   3330
         _Version        =   589884
         _ExtentX        =   5874
         _ExtentY        =   7091
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picRule 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   3375
         TabIndex        =   50
         Top             =   5520
         Width           =   3375
         Begin VB.Label lbl 
            Caption         =   "ѪҺ��˹涨"
            Height          =   3135
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.CheckBox chkIsShowAll 
         Caption         =   "ֻ��ʾ�������ҽ��"
         Height          =   180
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   3
         Top             =   788
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   315
         TabIndex        =   6
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   315
         TabIndex        =   5
         Top             =   420
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7335
      Left            =   3720
      TabIndex        =   36
      Top             =   1845
      Width           =   11355
      _Version        =   589884
      _ExtentX        =   20029
      _ExtentY        =   12938
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   10575
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21034
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
            Picture         =   "frmExamineTransfuse.frx":0000
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":005E
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":00BC
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":011A
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
            Picture         =   "frmExamineTransfuse.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineTransfuse.frx":0234
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
Attribute VB_Name = "frmExamineTransfuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mobjBar As CommandBar
Private mobjPopup As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean
Private mintAuditPrivs As Integer '��ǰ�û����Ȩ��:0:�������κ�Ȩ�ޣ��������ǩ���κ�����1��800ml���¿����ǩ����2��1600���¾��������ǩ����3������ˡ�ǩ�����С�
Private mstrButPri As String '�ܾ���ˡ���ˡ�ȡ����ˡ�ǩ����ȡ��ǩ�� ��ť�Ƿ���ã�1��ʾ���ã�0��ʾ�����ã��� 11010��
Private mbln������Ѫ������� As Boolean
Private mlngFindNum As Long
Private mstrChangeRows As String   '��¼�޸ĵ���
Private mstrǩ��IDs As String      'ȡ����˵�ʱ���¼һ�����˴������ǩ��ID
'���������ʱ������ǩ�����ܣ������жϼ��� And 1 = 0
Private mblnTmp As Boolean
Private mrsDefine As ADODB.Recordset
Private mint���� As Integer
Private mclsMipModule As zl9ComLib.clsMipModule
Private Enum Enum_Dor
    COL_��ԱID = 0
    col_���� = 1
    COL_רҵ����ְ�� = 2
    COL_����ְ�� = 3
    COL_ƴ������ = 4
    COL_��ʼ��� = 5
    COL_�������� = 6
    COL_��������ID = 7
End Enum

Private Enum Enum_Type  'ָ��Ҫ�������뵽��״̬
    t_����� = 1
    t_��ǩ�� = 2
    t_��ǩ�� = 3
    t_�Ѿܾ� = 4
End Enum

Private Enum Enum_Advice_New
    col_ѡ�� = 0
    COL_���˵�� = 1
    COL_���ʱ�� = 2
    COL_�������� = 3
    COL_ҽ������ = 4
    col_��Ч = 5
'�ü��ģʽ�����������͵���������������ҽ�����ݺϲ�
    COL_���� = 6
    col_��Ѫʱ�� = 7
    COL_��ʼʱ�� = 8
    col_��Ѫ���� = 9
    col_24h��Ѫ�� = 10
    col_���״̬˵�� = 11
'������
    col_ҽ��ID = 12
    col_���ID = 13
    col_�Ա� = 14
    col_���� = 15
    COL_���� = 16
    COL_��Ժʱ�� = 17
    col_���� = 18
    COL_���� = 19
    COL_����ȼ� = 20
    col_����Id = 21
    col_��ҳID = 22
    COL_��ID = 23
    COL_������� = 24
    COL_������Դ = 25
    COL_ǩ��id = 26
    COL_ҽ��״̬ = 27
    col_�Һŵ� = 28
    col_���״̬ = 29
    
    COL_����� = 30
    col_סԺ�� = 31
    COL_��ǰ���� = 32
    COL_����ҽ�� = 33
    COL_����ʱ�� = 34
    COL_��������ID = 35
    COL_��Ժ����ID = 36
    COL_��ǰ����ID = 37
    COL_����ĿID = 38
    
    col_��Ѫҽ�� = 39 '1����Ѫҽ����0����Ѫҽ��
    COL_������� = 40
    COL_�������� = 41
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

Public Function ShowMe(frmParent As Object, ByVal int���� As Integer, Optional ByRef ojbMip As Object) As Boolean
'������mint����=1���2סԺ
    On Error Resume Next
    
    mint���� = int����
    If Not ojbMip Is Nothing Then Set mclsMipModule = ojbMip
    Call frmExamineTransfuse.Show(0, frmParent)
End Function

Private Sub cboDept_Click()
    Call LoadDoc
End Sub
Private Sub cmdFindY_Click()
    Call LoadAdvice(False)
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
    On Error GoTo errH
    Screen.MousePointer = 11
    stbThis.Panels(2).Text = "��ѡ��һλ����ҽ����"
    rptDoc.Records.DeleteAll
    vsAudit.Rows = 1: vsAudit.AddItem ""
    
    '�˴��ų���Ѫҽ��
    If tbcSub.Selected.Tag = "�����" Or tbcSub.Selected.Tag = "��ǩ��" Then
        datBegint = CDate(dtpTimeY(0).Value)
        datEnd = CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct F.����ҽ�� From ������ĿĿ¼ K,����ҽ����¼ H ,����ҽ����¼ F " & _
            " Where (K.��������='8' and nvl(K.ִ�з���,0)=0 or K.��������='9') And K.ID=H.������ĿID And H.���ID=F.id And " & _
             IIf(mbln������Ѫ������� = False, " f.���״̬ in (1,7) ", IIf(tbcSub.Selected.Tag = "�����", " f.���״̬=1 ", " f.���״̬=7 ")) & " and f.ҽ��״̬=1 And F.����ʱ�� Between [4] And [5] And f.������Դ=[3] And f.������� ='K') F"
    Else
        datBegint = CDate(dtpTime(0).Value)
        datEnd = CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60)
        strTmp = ",(Select Distinct f.����ҽ�� From ������ĿĿ¼ K,����ҽ����¼ H ,����ҽ����¼ F,����ҽ��״̬ G " & _
            " Where (K.��������='8' and nvl(K.ִ�з���,0)=0  or K.��������='9') And K.ID=H.������ĿID And H.���ID=F.id And F.id=g.ҽ��id and G.�������� in (11,12,14)" & _
            " And G.����ʱ�� Between [4] And [5] And f.������Դ=[3] And f.������� ='K') F"
    End If
    
    strSQL = "Select DISTINCT a.Id, A.�Ա�" & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", ",b.����ID,e.���� as ��������") & ",a.����,a.רҵ����ְ��,a.����ְ��, Upper(zlSpellCode(a.����)) As ƴ������, Upper(Zlwbcode(a.����)) As ��ʼ���" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� D,���ű� E" & IIf(chkIsShowAll.Value, strTmp, "") & vbNewLine & _
            "Where a.Id = b.��Աid And e.ID=b.����ID And d.��Աid = a.Id  And d.��Ա���� = 'ҽ��' And " & vbNewLine & _
            "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & vbNewLine & _
            "   " & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.����id=[2]") & _
            IIf(chkIsShowAll.Value, " And  f.����ҽ�� = a.���� ", "")
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)), IIf(optOccasion(0).Value, 2, 1), datBegint, datEnd)
    
    With rptDoc
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!�Ա� & "" = "Ů", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!רҵ����ְ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!����ְ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!ƴ������ & "")
            Set objItem = objRecord.AddItem(rsTmp!��ʼ��� & "")
            If Val(cboDept.ItemData(cboDept.ListIndex)) <> -1 Then
                Set objItem = objRecord.AddItem(rsTmp!�������� & "")
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
            End If
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    mlngFindNum = 0
    Screen.MousePointer = 0
    Call vsAudit_KeyPress(vbKeyBack)
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
        strSubhead = rptDoc.SelectedRows(0).Record(col_����).Value & "��Ѫ����嵥"
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
        If Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value) = UserInfo.ID Then
            MsgBox "��������Լ��������Ѫҽ��", vbInformation, Me.Caption
            Exit Sub
        End If
        '�ж��Ƿ�Ϊ�°�ѪҺ���
'        If mbln������Ѫ������� Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_ѡ��) = flexChecked Then
                    Exit For
                End If
                If i = .Rows - 1 Then
                MsgBox "����δѡ����Ҫ�������Ŀ��Ϣ����ѡ��", vbInformation, Me.Caption
                Exit Sub
                End If
            Next
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_ѡ��) = flexChecked And .TextMatrix(i, COL_���˵��) <> "" Then
                    If MsgBox("ȷ�����ɾ��˵����������", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
                    
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
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
        mstrChangeRows = ""
        mblnIsUpdate = False
        End With
End Sub

Private Sub SaveAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, ByVal strDate As String)
'���ܣ����������Ϣ
'�������ӵڼ��п�ʼ�����ڼ��н�����ͬһ�����ˣ�
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String
    Dim strSource As String, strSign As String
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim strSignSQL As String
    Dim int״̬ As Integer
    Dim lngMsgRow As Long, lngBlood As Long
    Dim rsTmp As ADODB.Recordset
    Dim intQuestion As Integer, intAudit As Integer
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .TextMatrix(i, col_ҽ��ID) = "" Then Exit Sub
            If .Cell(flexcpChecked, i, col_ѡ��) = flexChecked Then
                If tbcSub.Selected.Tag = "��ǩ��" Then
                    int״̬ = 3
                ElseIf tbcSub.Selected.Tag = "�����" Then
                    If mbln������Ѫ������� Then
                        Select Case Val(.TextMatrix(i, col_24h��Ѫ��))
                            Case Is >= 1600
                                intAudit = 3
                            Case Is >= 800
                                intAudit = 2
                            Case Else
                                intAudit = 1
                        End Select
                        If mintAuditPrivs >= intAudit And intAudit > 1 Then 'ֻ��800���ϲŻ����˽׶�
                            If intQuestion = 0 Then
                                If MsgBox("������ֱ�����" & .TextMatrix(i, COL_��������) & "�����ǩ����������ǡ�ֱ����ɣ�������񡱽�������ˡ�", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                                    intQuestion = 1
                                Else
                                    intQuestion = 2
                                End If
                            End If
                            If intQuestion = 1 Then
                                int״̬ = 3
                            ElseIf intQuestion = 2 Then
                                int״̬ = 6
                            End If
                        Else
                            If intAudit > 1 Then
                                int״̬ = 6
                            Else
                                int״̬ = 3
                            End If
                        End If
                    Else    '�������������ֱ��ͨ����
                        int״̬ = 3
                    End If
                End If
                If int״̬ = 3 Then lngMsgRow = i  '������ǩ�����ҽ���Ϳ��Է���
                '���δ����Ѫ��ϵͳ��״̬3ΪѪ������գ�����Ϊ���ͨ����״̬1��
                If int״̬ = 3 And Not gblnѪ��ϵͳ Then int״̬ = 1
                strSQL = Val(.TextMatrix(i, col_ҽ��ID)) & "|" & "Zl_ҽ����˹���_Audit(" & Val(.TextMatrix(i, col_ҽ��ID)) & "," & int״̬ & "," & _
                        "'" & UserInfo.���� & "'," & strDate & ",''"
                colsql.Add strSQL, "C" & colsql.Count + 1
            End If
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colsql.Count
        strSQL = Mid(colsql("C" & i), InStr(colsql("C" & i), "|") + 1) & ",2)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    If strSignSQL <> "" Then
        Call zlDatabase.ExecuteProcedure(strSignSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If mint���� = 2 Then
        '����ҽ�����´���Ϣ/��Ѫ��Ѫ������Ϣ
        With vsAudit
            If lngMsgRow <> 0 Then
                strSQL = "select a.�������� from ������ĿĿ¼ a where a.id=[1]"
                If Not mbln������Ѫ������� Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngMsgRow, COL_����ĿID)))
                    Call ZLHIS_CIS_001(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_����Id)), .TextMatrix(lngMsgRow, COL_��������), .TextMatrix(lngMsgRow, col_סԺ��), "", Val(.TextMatrix(lngMsgRow, COL_������Դ)), Val(.TextMatrix(lngMsgRow, col_��ҳID)), _
                        Val(.TextMatrix(lngMsgRow, COL_��ǰ����ID)), "", Val(.TextMatrix(lngMsgRow, COL_��Ժ����ID)), "", "", .TextMatrix(lngMsgRow, COL_��ǰ����), Val(.TextMatrix(lngMsgRow, col_ҽ��ID)), 0, 1, "K", rsTmp!�������� & "", _
                        .TextMatrix(lngMsgRow, COL_����ҽ��), .TextMatrix(lngMsgRow, COL_����ʱ��), .TextMatrix(lngMsgRow, COL_��������ID), "")
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngMsgRow, COL_����ĿID)))
                    Call ZLHIS_CIS_001(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_����Id)), .TextMatrix(lngMsgRow, COL_��������), .TextMatrix(lngMsgRow, col_סԺ��), "", Val(.TextMatrix(lngMsgRow, COL_������Դ)), Val(.TextMatrix(lngMsgRow, col_��ҳID)), _
                        Val(.TextMatrix(lngMsgRow, COL_��ǰ����ID)), "", Val(.TextMatrix(lngMsgRow, COL_��Ժ����ID)), "", "", .TextMatrix(lngMsgRow, COL_��ǰ����), Val(.TextMatrix(lngMsgRow, col_ҽ��ID)), 0, 1, "K", rsTmp!�������� & "", _
                        .TextMatrix(lngMsgRow, COL_����ҽ��), .TextMatrix(lngMsgRow, COL_����ʱ��), .TextMatrix(lngMsgRow, COL_��������ID), "")
                End If
            End If
            If lngBlood <> 0 Then
                lngMsgRow = lngBlood
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        Call ZLHIS_CIS_031(mclsMipModule, Val(.TextMatrix(lngMsgRow, col_����Id)), .TextMatrix(lngMsgRow, COL_��������), .TextMatrix(lngMsgRow, col_סԺ��), "", Val(.TextMatrix(lngMsgRow, COL_������Դ)), Val(.TextMatrix(lngMsgRow, col_��ҳID)), _
                            Val(.TextMatrix(lngMsgRow, COL_��ǰ����ID)), "", Val(.TextMatrix(lngMsgRow, COL_��Ժ����ID)), "", "", .TextMatrix(lngMsgRow, COL_��ǰ����), Val(.TextMatrix(lngMsgRow, col_ҽ��ID)), _
                            .TextMatrix(lngMsgRow, COL_����ҽ��), .TextMatrix(lngMsgRow, COL_����ʱ��), .TextMatrix(lngMsgRow, COL_��������ID), "")
                    End If
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
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 And Me.Visible Then
        If chkIsShowAll.Value = 1 Then
            Call LoadDoc
        Else
            Call LoadAdvice(True)
        End If
    End If
End Sub

Private Sub CancleAudit()
'ȡ�����
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnIsCheck As Boolean
    
    With vsAudit
        If Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value) = UserInfo.ID Then
            MsgBox "����ȡ���Լ��������Ѫҽ��", vbInformation, Me.Caption
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = vbChecked Then Exit For
            If i = .Rows - 1 Then
                Call MsgBox("δ��ѡҪȡ������Ŀ�����֤��", vbInformation, Me.Caption)
                Exit Sub
            End If
        Next
        '�ж��Ƿ��й�ѡ�ģ��й�ѡ���Թ�ѡΪ׼
        For i = i To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = vbChecked Then
                If RowInͬһ����(i, lngBegin, lngEnd, vsAudit) Then
                    Call CancleAuditOnePati(lngBegin, lngEnd)
                    i = lngEnd
                Else
                    Call CancleAuditOnePati(i, i)
                End If
            End If
        Next
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
    Dim strSource As String, strSign As String, strDate As String
    Dim lng֤��ID As Long, lngǩ��ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim rsSQL As New ADODB.Recordset, strExp As String, strTmp As String
    Dim arrIDs(4) As String, arrExp(4) As String
    Dim rsChk As Recordset
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .Cell(flexcpChecked, i, 0) = vbChecked Then
                If tbcSub.Selected.Tag = "��ǩ��" And .TextMatrix(i, COL_��������) = 18 Then '������˹��̣������������
                    arrIDs(t_�����) = arrIDs(t_�����) & IIf(arrIDs(t_�����) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                    arrExp(t_�����) = arrExp(t_�����) & IIf(arrExp(t_�����) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                ElseIf tbcSub.Selected.Tag = "��ǩ��" And Nvl(.TextMatrix(i, COL_��������), 0) = 0 Then '��������˹��̣��ܾ�������
                    If MsgBox("����ȡ�� " & .TextMatrix(i, COL_��������) & "��" & .TextMatrix(i, COL_ҽ������) & "�����룬��ֱ�Ӿܾ���������룬�Ƿ�ȷ����", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        arrIDs(t_�Ѿܾ�) = arrIDs(t_�Ѿܾ�) & IIf(arrIDs(t_�Ѿܾ�) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        arrExp(t_�Ѿܾ�) = arrExp(t_�Ѿܾ�) & IIf(arrExp(t_�Ѿܾ�) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                    End If
                ElseIf tbcSub.Selected.Tag = "�����" Then  '�ܾ��������
                    arrIDs(t_�Ѿܾ�) = arrIDs(t_�Ѿܾ�) & IIf(arrIDs(t_�Ѿܾ�) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                    arrExp(t_�Ѿܾ�) = arrExp(t_�Ѿܾ�) & IIf(arrExp(t_�Ѿܾ�) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                ElseIf tbcSub.Selected.Tag = "�����" Then  '���ݴ�����������˻���������˻���ȡ���ܾ�
                    If .TextMatrix(i, col_���״̬) = 3 Then
                        If MsgBox("������� " & .TextMatrix(i, COL_��������) & "��" & .TextMatrix(i, COL_ҽ������) & ",�Ƿ�ֱ����ˣ�", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                            arrIDs(t_�����) = arrIDs(t_�����) & IIf(arrIDs(t_�����) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                            arrExp(t_�����) = arrExp(t_�����) & IIf(arrExp(t_�����) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                        Else
                            arrIDs(t_��ǩ��) = arrIDs(t_��ǩ��) & IIf(arrIDs(t_��ǩ��) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                            arrExp(t_��ǩ��) = arrExp(t_��ǩ��) & IIf(arrExp(t_��ǩ��) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                        End If
                    ElseIf .TextMatrix(i, col_���״̬) = 4 Or .TextMatrix(i, col_���״̬) = 2 Then
                        arrIDs(t_�����) = arrIDs(t_�����) & IIf(arrIDs(t_�����) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        arrExp(t_�����) = arrExp(t_�����) & IIf(arrExp(t_�����) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                    End If
                ElseIf tbcSub.Selected.Tag = "��ǩ��" And (.TextMatrix(i, col_���״̬) = 4 Or .TextMatrix(i, col_���״̬) = 2) Then
                    strSQL = "select 1 from ����ҽ��״̬ where �������� = 17 and  ҽ��id = [1] and rownum < 2"
                    Set rsChk = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ҽ���Ƿ�ǩ����", .TextMatrix(i, col_ҽ��ID))
                    If rsChk.BOF Then '�����ݣ��޷��жϣ��˻��������
                        arrIDs(t_�����) = arrIDs(t_�����) & IIf(arrIDs(t_�����) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        arrExp(t_�����) = arrExp(t_�����) & IIf(arrExp(t_�����) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                    Else
                        arrIDs(t_��ǩ��) = arrIDs(t_��ǩ��) & IIf(arrIDs(t_��ǩ��) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        arrExp(t_��ǩ��) = arrExp(t_��ǩ��) & IIf(arrExp(t_��ǩ��) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                    End If
                ElseIf tbcSub.Selected.Tag = "��ǩ��" And .TextMatrix(i, col_���״̬) = 3 Then    '��ǩ��ҳ��ȡ���ܾ�
                    If MsgBox("����ǩ�� " & .TextMatrix(i, COL_��������) & "��" & .TextMatrix(i, COL_ҽ������) & ",�Ƿ�ֱ��ǩ����", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        arrIDs(t_��ǩ��) = arrIDs(t_��ǩ��) & IIf(arrIDs(t_��ǩ��) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                        arrExp(t_��ǩ��) = arrExp(t_��ǩ��) & IIf(arrExp(t_��ǩ��) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                    Else
                        strSQL = "select 1 from ����ҽ��״̬ where �������� = 17 and  ҽ��id = [1] and rownum < 2"
                        Set rsChk = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ҽ���Ƿ�ǩ����", .TextMatrix(i, col_ҽ��ID))
                        If rsChk.BOF Then '�����ݣ��޷��жϣ��˻��������
                            arrIDs(t_�����) = arrIDs(t_�����) & IIf(arrIDs(t_�����) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                            arrExp(t_�����) = arrExp(t_�����) & IIf(arrExp(t_�����) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                        Else
                            arrIDs(t_��ǩ��) = arrIDs(t_��ǩ��) & IIf(arrIDs(t_��ǩ��) = "", "", ",") & .TextMatrix(i, col_ҽ��ID)
                            arrExp(t_��ǩ��) = arrExp(t_��ǩ��) & IIf(arrExp(t_��ǩ��) = "", "", ",") & .TextMatrix(i, COL_���˵��)
                        End If
                    End If
                End If
            End If


        Next
    End With
    Call CancleAuditOnePatiChild(arrIDs(t_�����), strDate, arrExp(t_�����), 1)
    Call CancleAuditOnePatiChild(arrIDs(t_��ǩ��), strDate, arrExp(t_��ǩ��), 7)
    Call CancleAuditOnePatiChild(arrIDs(t_��ǩ��), strDate, arrExp(t_��ǩ��), IIf(gblnѪ��ϵͳ, 4, 2))
    Call CancleAuditOnePatiChild(arrIDs(t_�Ѿܾ�), strDate, arrExp(t_�Ѿܾ�), 3)
    mstrǩ��IDs = "0"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CancleAuditOnePatiChild(ByVal strIDs As String, ByVal strDate As String, ByVal strExp As String, ByVal intType As Integer)
    Dim strTmp As String, strSQL As String
    Dim rsSQL As New ADODB.Recordset
    Dim i As Long, blnTrans As Boolean
    
    On Error GoTo errH
    If strIDs <> "" Then
        Call SQLRecord(rsSQL)
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = "Zl_ҽ����˹���_Cancel('" & strIDs & "',2," & intType & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
        For i = 0 To UBound(Split(strIDs, ","))
            If strExp <> "" Then
                If UBound(Split(strExp, ",")) <= i Then
                    strTmp = Split(strExp, ",")(i)
                Else
                    strTmp = ""
                End If
            End If
            strSQL = "Zl_ҽ����˹���_Update('" & Split(strIDs, ",")(i) & "'," & strDate & ",'" & strTmp & "',2,'" & UserInfo.���� & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
        Next
        If Not SQLRecordExecute(rsSQL, Me.Caption) Then blnTrans = False
    End If
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
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext <> conMenu_Edit_Audit Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
     Case conMenu_Edit_Audit '�Ҽ��������
        For i = vsAudit.FixedRows To vsAudit.Rows - 1
            If vsAudit.Cell(flexcpChecked, i, col_ѡ��) = vbChecked Then Exit For
        Next
        If i < vsAudit.Rows And vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_ѡ��) = vbUnchecked Then
            If MsgBox("������˲���ֻ�����ѹ�ѡ��ҽ�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
        ElseIf i = vsAudit.Rows And vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_ѡ��) = vbUnchecked Then
            vsAudit.Cell(flexcpChecked, vsAudit.RowSel, col_ѡ��) = vbChecked
        End If
        Call SaveAudit
    Case conMenu_Edit_AdviceUnAudit 'ȡ�����
        Call CancleAudit
        Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_UnUse, conMenu_Edit_StopAudit ' �ܾ���� ȡ��ǩ��
        Call CancleAudit
        Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_MediAudit, conMenu_Edit_Send '���,ǩ��
            Call SaveAudit
            Call vsAudit_CellChanged(vsAudit.Row, vsAudit.Col)
    Case conMenu_Edit_ApplyView '�鿴��Ѫ���뵥
        If vsAudit.Row <= 0 Then Exit Sub
        If Val(vsAudit.TextMatrix(vsAudit.Row, col_ҽ��ID)) = 0 Then Exit Sub
        Call gobjKernel.ShowBloodApply(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_ҽ��ID)))
    Case conMenu_Tool_Archive '���Ӳ�������
        If vsAudit.Row <= 0 Then Exit Sub
        If Val(vsAudit.TextMatrix(vsAudit.Row, col_ҽ��ID)) = 0 Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_����Id)), Val(vsAudit.TextMatrix(vsAudit.Row, col_��ҳID)))
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
        If tbcSub.Selected.Tag = "�����" Or tbcSub.Selected.Tag = "��ǩ��" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
        If mbln������Ѫ������� Then Call vsAudit_KeyPress(vbKeyBack)
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
        If Not mbln������Ѫ������� Then rptDoc.Height = .Height - 1600
    End With
    If mbln������Ѫ������� Then
        rptDoc.Height = fraDoctor.Height - 1600 - picRule.Height
        picRule.Top = rptDoc.Top + rptDoc.Height
    End If
        picRule.Width = rptDoc.Width
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

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '����Ȩ�����ð�ť�ɼ�״̬
    Select Case Control.ID
        Case conMenu_Edit_AdviceUnAudit 'ȡ�����
            If Not mbln������Ѫ������� Then
                If tbcSub.Selected.Tag <> "�����" Then Control.Visible = False: Exit Sub
            Else
                If tbcSub.Selected.Tag <> "��ǩ��" And mbln������Ѫ������� Then Control.Visible = False: Exit Sub
            End If
        Case conMenu_Edit_Send
            If tbcSub.Selected.Tag <> "��ǩ��" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_MediAudit, conMenu_Edit_UnUse
            If tbcSub.Selected.Tag <> "�����" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_StopAudit
            If tbcSub.Selected.Tag <> "��ǩ��" Then Control.Visible = False: Exit Sub
        Case conMenu_Edit_AdviceUnAudit
            If tbcSub.Selected.Tag <> "�����" Then Control.Visible = False: Exit Sub
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
        Case conMenu_Edit_AdviceUnAudit 'ȡ�����
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 3, 1) = "1"
            'If Not mbln������Ѫ������� Then Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, col_ҽ��״̬) = "1"
        Case conMenu_Edit_UnUse '�ܾ����
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 1, 1) = "1"
        Case conMenu_Edit_MediAudit '���
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 1, 1) = "1"
        Case conMenu_Edit_StopAudit 'ȡ��ǩ��
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And Mid(mstrButPri, 5, 1) = "1" And mbln������Ѫ�������
        Case conMenu_Edit_Send 'ǩ��
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, col_���״̬) = "7" And Mid(mstrButPri, 4, 1) = "1" And mbln������Ѫ�������
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
            Else
                cboDept.Enabled = True
                txtFind.Enabled = True
                fraDoctor.Enabled = True
                cboTime.Enabled = True
                cmdFind.Enabled = True
                cboDept.BackColor = &H80000005
                txtFind.BackColor = &H80000005
            End If
        
        Case conMenu_Edit_ApplyView
            Control.Enabled = Val(vsAudit.TextMatrix(vsAudit.Row, COL_�������)) > 0
        Case conMenu_Tool_Archive '���Ӳ�������
            Control.Enabled = vsAudit.Row > 0
            If Control.Enabled Then
                Control.Enabled = Val(vsAudit.TextMatrix(vsAudit.Row, col_ҽ��ID)) <> 0
            End If
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
    If chkIsShowAll.Value = 1 Then
        Call LoadDoc
    Else
        Call LoadAdvice(True)
    End If
End Sub

Private Sub GetLocalSetting()
'��ȡ���ز���
    cboTime.ListIndex = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", 0)
    mblnTmp = True
    chkIsShowAll.Value = Val(zlDatabase.GetPara("ֻ��ʾ�������ҽ��", glngSys, mlngModul, 0) & "")
    mblnTmp = False
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mbln������Ѫ������� = gbln��Ѫ�����������
    lbl.Visible = mbln������Ѫ�������
    lbl.Enabled = mbln������Ѫ�������
    lbl.Caption = "�����л����񹲺͹���������85����" & vbCrLf & vbCrLf & _
                    "һ��ͬһ����24Сʱ�����뱸Ѫ������800ml�ģ��ɾ����м�����רҵ����ְ����ְ�ʸ��ҽʦ������룬�ϼ�ҽʦ��׼ǩ���󣬷��ɱ�Ѫ��" & vbCrLf & _
                    "����ͬһ����24Сʱ�����뱸Ѫ����800ml������-1600ml֮��ģ��ɾ����м�����רҵ����ְ����ְ�ʸ��ҽʦ������룬�ϼ�ҽʦ��ˣ��������κ�׼ǩ���󣬷��ɱ�Ѫ��" & vbCrLf & _
                    "����ͬһ����24Сʱ�����뱸Ѫ���ﵽ�򳬹�1600ml�ģ��ɾ����м�����רҵ����ְ����ְ�ʸ��ҽʦ������룬����������˺󣬱�ҽ������׼ǩ�������ɱ�Ѫ��" & vbCrLf & _
                    "��������涨�������ڼ�����Ѫ��"
    lbl.ForeColor = vbBlue
    mstrPrivs = GetInsidePrivs(p��Ѫ��˹���)
    If mbln������Ѫ������� Then Call GetPower
    mlngModul = p��Ѫ��˹���
    mblnIsUpdate = False
    mstrChangeRows = ""
    mstrǩ��IDs = "0"
    optOccasion(IIf(mint���� = 2, 0, 1)).Value = True
    
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
        If mbln������Ѫ������� Then
            .InsertItem(0, "  �����  ", picUnAudited.hwnd, 0).Tag = "�����"
            .InsertItem(1, "  ��ǩ��  ", picUnAudited.hwnd, 0).Tag = "��ǩ��"
            .InsertItem(2, "  ��ǩ��  ", picUnAudited.hwnd, 0).Tag = "��ǩ��"
            
            .Item(2).Selected = True
            .Item(1).Selected = True
            .Item(0).Selected = True
        Else
            .InsertItem(0, "  �����  ", picUnAudited.hwnd, 0).Tag = "�����"
            .InsertItem(1, "  �����  ", picUnAudited.hwnd, 0).Tag = "�����"
            lblDate.Caption = "���ʱ��"
            .Item(1).Selected = True
            .Item(0).Selected = True
        End If
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
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    'vsFlexGrid
    '-----------------------------------------------------
        strHead = ",450,1;���˵��,2000,1;���ʱ��;��������,1000,1;ҽ������,3460,1;��Ч,500,1;����;��Ѫʱ��,1550,1;��ʼʱ��,1550,1;����������������,1550,1;24Сʱ��Ѫ��,1000,1;���״̬˵��,800,1;ҽ��ID;���ID ; �Ա�;����;����;��Ժʱ��;����; ���; ����;����ȼ�;����ID; ��ҳID; ��ID;������� ;������Դ;ǩ��id;ҽ��״̬;�Һŵ�;���״̬"
        strHead = strHead & ";�����;סԺ��;��ǰ����;����ҽ��;����ʱ��;��������id;��Ժ����id;��ǰ����id;����ĿID;��Ѫҽ��;�������;����״̬"
        Call Grid.Init(vsAudit, strHead)
        vsAudit.ExtendLastCol = True
        vsAudit.Editable = flexEDKbdMouse
        vsAudit.ColDataType(col_ѡ��) = flexDTBoolean
        vsAudit.Cell(flexcpChecked, 0, col_ѡ��) = flexcpChecked
        vsAudit.Cell(flexcpPictureAlignment, 0, col_ѡ��) = flexPicAlignCenterCenter
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mrsDefine = InitAdviceDefine
    Call GetLocalSetting '���ز���
    Call LoadDept
End Sub

Private Sub GetPower()
    Dim strSQL As String
    Dim rs As Recordset
    
    mintAuditPrivs = 0
    '����ҽ���Ȩ�ޣ�������������
    If InStr(";" & mstrPrivs & ";", ";ҽ���;") > 0 Then
        mintAuditPrivs = 3
    ElseIf InStr(";" & mstrPrivs & ";", ";������;") > 0 Then
        mintAuditPrivs = 2
    Else
        strSQL = "select רҵ����ְ��,����ְ�� from ��Ա�� where id = " & UserInfo.ID
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "��ȡְ����Ϣ")
        If Not rs.BOF Then
            If rs("רҵ����ְ��") = "����ҽʦ" Or rs("רҵ����ְ��") = "����ҽʦ" Or rs("רҵ����ְ��") = "������ҽʦ" Then
                mintAuditPrivs = 1
            End If
        End If
    End If
End Sub

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
        If Me.Visible Then
            If chkIsShowAll.Value = 1 Then
                Call LoadDoc
            Else
                Call LoadAdvice '(True)
            End If
        End If
    End If
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
            "  And C.�������� = '�ٴ�' And Instr([2],C.������� || '')>0   And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"

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
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptDoc
        
        Set objCol = .Columns.Add(COL_��ԱID, "��ԱID", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 70, True)
        Set objCol = .Columns.Add(COL_רҵ����ְ��, "רҵ����ְ��", 80, True)
        Set objCol = .Columns.Add(COL_����ְ��, "����ְ��", 80, True)
        Set objCol = .Columns.Add(COL_ƴ������, "ƴ������", 0, False)
        Set objCol = .Columns.Add(COL_��ʼ���, "��ʼ���", 0, False)
        Set objCol = .Columns.Add(COL_��������, "��������", 0, False)
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
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MediAudit, "���(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "�ܾ����(&R)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����(&U)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ǩ��(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "ȡ��(&Q)")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
        objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Preview
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
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MediAudit, "���(&U)")
            objControl.BeginGroup = True
            objControl.IconId = 21904
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "�ܾ����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ǩ��(&A)")
            objControl.BeginGroup = True
            objControl.IconId = 21904
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����(&U)")
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "ȡ��(&Q)")
            objControl.BeginGroup = True
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_File_Preview
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
    
    '�Ҽ��˵�(������˹���)
    '-----------------------------------------------------
    Set mobjPopup = cbsMain.Add("�Ҽ��˵�", xtpBarPopup)
    With mobjPopup.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "ȡ�����")
        objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������")
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
    mlngFindNum = 0
    Set mclsMipModule = Nothing
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", cboTime.ListIndex
    zlDatabase.SetPara "ֻ��ʾ�������ҽ��", chkIsShowAll.Value & "", glngSys, mlngModul
End Sub


Private Sub optOccasion_Click(Index As Integer)
    If Me.Visible Then
                Call LoadDept
        vsAudit.Rows = 1
        vsAudit.AddItem ""
    End If
End Sub

Private Sub picUnAudited_Resize()
    On Error Resume Next
    picDate.Move 0, 0, picUnAudited.Width
    picDateY.Move 0, 0, picUnAudited.Width
    vsAudit.Move 0, picDate.Top + picDate.Height, picUnAudited.Width, picUnAudited.Height - picDate.Top + picDate.Height
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '����ҽ���б�
    If tbcSub.Selected.Tag = "�����" Then
        If Me.Visible Then
            Call LoadAdvice
            If mbln������Ѫ������� Then Call vsAudit_KeyPress(vbKeyBack)
        End If
    ElseIf tbcSub.Selected.Tag = "��ǩ��" Then
        If Me.Visible Then
            Call LoadAdvice
            Call vsAudit_KeyPress(vbKeyBack)
        End If
    Else
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible And chkIsShowAll.Value = 1 Then
        Call LoadDoc
    End If
    With vsAudit
        If mbln������Ѫ������� Then
            If Item.Tag = "��ǩ��" Then
                picDate.Visible = True
                picDateY.Visible = False
                Call picUnAudited_Resize
                .ColWidth(COL_���ʱ��) = 1800
                .ColHidden(COL_���ʱ��) = False
                .TextMatrix(0, COL_���ʱ��) = "ǩ��ʱ��"
                .Cell(flexcpChecked, 1, 0) = 2
                If Me.Visible Then Call LoadAdvice(True)
                .ColHidden(col_���״̬˵��) = False
                lblPri(1).ForeColor = vbBlue
                lblPri(1).Caption = "����ȡ��ǩ����ѪҺ������Χ��" & IIf(mintAuditPrivs >= 1, "800ml����   ", "") & IIf(mintAuditPrivs >= 2, "800ml-1600ml   ", "") & IIf(mintAuditPrivs = 3, "1600ml������", "")
                If lblPri(1).Caption = "����ǩ����ѪҺ������Χ��" Then
                    lblPri(1).Caption = "��������ȡ��ǩ��ѪҺ��Ȩ�ޣ�"
                    lblPri(1).ForeColor = vbRed
                End If
                lblPri(1).Width = 8000
            ElseIf Item.Tag = "��ǩ��" Then
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Caption = "����ǩ����ѪҺ������Χ��" & IIf(mintAuditPrivs >= 1, "800ml����   ", "") & IIf(mintAuditPrivs >= 2, "800ml-1600ml   ", "") & IIf(mintAuditPrivs = 3, "1600ml������", "")
                lblPri(0).ForeColor = vbBlue
                If lblPri(0).Caption = "����ǩ����ѪҺ������Χ��" Then
                    lblPri(0).Caption = "��������ǩ��ѪҺ��Ȩ�ޣ�"
                    lblPri(0).ForeColor = vbRed
                End If
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_���ʱ��) = 1800
                .ColHidden(COL_���ʱ��) = False
                .TextMatrix(0, COL_���ʱ��) = "���ʱ��"
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_���״̬˵��) = True
            ElseIf Item.Tag = "�����" Then
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Caption = "������˵�ѪҺ������Χ��" & IIf(mintAuditPrivs >= 1, "800ml����   ", "") & IIf(mintAuditPrivs >= 1, "800ml-1600ml   ", "") & IIf(mintAuditPrivs >= 2, "1600ml������", "")
                lblPri(0).ForeColor = vbBlue
                If lblPri(0).Caption = "������˵�ѪҺ������Χ��" Then
                    lblPri(0).Caption = "�����������ѪҺ��Ȩ�ޣ�"
                    lblPri(0).ForeColor = vbRed
                End If
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_���ʱ��) = 0
                .ColHidden(COL_���ʱ��) = True
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_���״̬˵��) = True
            End If
            If Me.Visible Then Call vsAudit_KeyPress(vbKeyBack)
        Else
            .ColHidden(col_24h��Ѫ��) = True
            .ColHidden(col_��Ѫ����) = True
            If Item.Tag = "�����" Then
                picDate.Visible = True
                picDateY.Visible = False
                Call picUnAudited_Resize
                .ColWidth(COL_���ʱ��) = 1800
                .ColHidden(COL_���ʱ��) = False
                .TextMatrix(0, COL_���ʱ��) = "���ʱ��"
                .Cell(flexcpChecked, 1, 0) = 2
                If Me.Visible Then Call LoadAdvice(True)
                .ColHidden(col_���״̬˵��) = False
            Else
                picDate.Visible = False
                picDateY.Visible = True
                lblPri(0).Width = 8000
                Call picUnAudited_Resize
                .ColWidth(COL_���ʱ��) = 0
                .ColHidden(COL_���ʱ��) = True
                If Me.Visible Then Call LoadAdvice
                .ColHidden(col_���״̬˵��) = True
            End If
        End If
    End With
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
    Dim rsTmp As Recordset, rsTemp As Recordset
    Dim strSQL As String, strTemp As String
    Dim strDoorIds As String, strKey As String
    Dim i As Long, j As Long
    Dim lngID As Long       '���ڶ�λ
    Dim strFormat As String
    Dim strTmp As String, strType As String, int���״̬ As Integer
    Dim blnDo As Boolean
    Dim strPatis As String   '������Ϣ�ַ���������ID1:��ҳID1,����ID2:��ҳID2������
    Dim strDate As String '�����ַ���
    Dim intѡ�� As Integer
    Dim dbl��Ѫ���� As Double, dbl24h�� As Double
    
    
    If tbcSub.Selected.Tag = "��ǩ��" Then
        strType = "C.�������� = 18"
        int���״̬ = 7
    ElseIf tbcSub.Selected.Tag = "�����" Then
        strType = "C.�������� = 19"
        int���״̬ = 1
    Else
        strType = "C.�������� in(11,12,14,15)"
    End If
    strSQL = "Select Decode(a.������Դ, 2, a.����id || '_' || a.��ҳid || '_' || Nvl(a.Ӥ��, 0), 1, a.����id || '_' || a.�Һŵ�) Key,a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, a.�������,  Null As ѡ��, Null As ����, " & vbNewLine & _
            " Decode(Nvl(a.Ӥ��, 0), 0, a.����, Nvl(q.Ӥ������, a.���� || '֮Ӥ' || q.���)) As ����,Decode(Nvl(a.Ӥ��, 0), 0, a.�Ա�, q.Ӥ���Ա�) As �Ա�," & vbNewLine & _
            " Decode(Nvl(a.Ӥ��, 0), 0, a.����, (Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��) || '��')) As ����, p.��ǰ���� As ����," & vbNewLine & _
            "       Decode(Nvl(a.ҽ����Ч, 0), 0, '����', '����') As ��Ч, To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI') As ��ʼʱ��, a.ҽ������,a.���״̬," & vbNewLine & _
            "       Decode(a.�ܸ�����, Null, Null, a.�ܸ����� || b.���㵥λ) As ����, NVL(to_char(A.����ʱ��,'YYYY-MM-DD HH24:MI'),a.�걾��λ) As ��Ѫʱ��, a.ִ��ʱ�䷽�� As ִ��ʱ�䷽��, a.����id," & vbNewLine & _
            "       a.��ҳid, a.������Ŀid, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, b.���㵥λ As ������λ, e.����,e.��Ժ����,e.��Ժ����,f.���� as ����ȼ�,a.������Դ,A.�Һŵ�,a.�������" & vbNewLine & _
            ", c.��������, c.����˵��, c.ǩ��id ,a.ҽ��״̬,c.����ʱ�� as ���ʱ��" & _
            ",p.�����,p.סԺ��,p.��ǰ����,a.����ҽ��,To_Char(a.����ʱ��,'YYYY-MM-DD HH24:MI') As ����ʱ��,a.��������id,e.��Ժ����id,e.��ǰ����id,a.������Ŀid,h.ִ�з���" & _
            " From ����ҽ����¼ A, ������Ϣ P, ������ĿĿ¼ B, ������ҳ E,�շ���ĿĿ¼ F" & vbNewLine & _
            ", (Select ҽ��id,����ʱ��,����˵��,��������,ǩ��ID" & vbNewLine & _
                            "From (Select C.ҽ��id,C.����ʱ��,C.����˵��,C.��������,C.ǩ��ID, Row_Number() Over(Partition By C.ҽ��id Order By C.����ʱ�� Desc) Top" & vbNewLine & _
                            "       From ����ҽ��״̬ C" & vbNewLine & _
                            "       Where c.����ʱ�� Between " & IIf(InStr(1, tbcSub.Selected.Tag, "��") > 0, "[3] And [4]", "[6] And [7]") & vbNewLine & _
                            "       and " & strType & " And C.������Ա =[2])" & vbNewLine & _
                            "Where Top = 1)  C" & ",����ҽ����¼ G,������ĿĿ¼ H,������������¼ Q" & _
            " Where a.����id = p.����id And a.������Ŀid = b.Id  And f.id(+)=e.����ȼ�id  And" & vbNewLine & _
            "      e.����id(+) = a.����id And e.��ҳid(+) = a.��ҳid and g.������� = 'E' And a.id=g.���id and g.������Ŀid=h.id And (H.��������='8' and nvl(H.ִ�з���,0)=0  or H.��������='9')  and A.����ID = Q.����ID(+) and A.��ҳID = Q.��ҳID(+) and A.Ӥ�� = Q.���(+) " & _
            IIf(InStr(1, tbcSub.Selected.Tag, "��") > 0, " And c.ҽ��id = a.Id ", _
            " AND a.id=c.ҽ��id(+) And A.����ʱ�� between [6] and [7] And a.ҽ��״̬ = 1 And a.���״̬ = [8] ") & vbNewLine & _
            "    And a.����ҽ��=[1] And A.������Դ=[5] And a.������� ='K'  And a.���ID is null " & _
            " Order By p.����,To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI'),Nvl(a.���id, a.Id),a.id"
            '" & IIf(tbcSub.Selected.Tag = "�����", "And a.ҽ��״̬ = 1 ", "") & "
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rptDoc.SelectedRows(0).Record(col_����).Value, UserInfo.����, CDate(dtpTime(0).Value), CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60), IIf(optOccasion(0).Value, 2, 1), CDate(dtpTimeY(0).Value), CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60), int���״̬)
    If mbln������Ѫ������� Then    '���������������ʱ������24Сʱ��Ѫ���Լ���Ѫ����
        If Not rsTmp.BOF Then
            rsTmp.MoveFirst
            strTemp = ""
            Do While Not rsTmp.EOF
                If optOccasion(0).Value = True Then
                    strKey = rsTmp("����id") & ":" & rsTmp("��ҳid")
                Else
                    strKey = rsTmp("�Һŵ�")
                End If
                If InStr("," & strTemp & ",", "," & strKey & ",") = 0 Then
                        strTemp = strTemp & "," & strKey
                End If
                rsTmp.MoveNext
            Loop
            If Left(strTemp, 1) = "," Then strTemp = Mid(strTemp, 2)
            If optOccasion(0).Value Then
                strSQL = _
                    " Select Key, Id, ����ʱ��, ������, ��Ѫʱ��" & vbNewLine & _
                    " From (With ҽ����¼ As (Select /*+ CARDINALITY(d,10) */" & vbNewLine & _
                    "                     a.����id || '_' || a.��ҳid || '_' || Nvl(a.Ӥ��, 0) Key, a.Id," & vbNewLine & _
                    "                     Decode(Nvl(e.ҽ��id, 0), 0, a.������Ŀid, e.������Ŀid) ������Ŀid," & vbNewLine & _
                    "                     Decode(Nvl(e.ҽ��id, 0), 0, a.�ܸ�����, e.������) ������, a.����ʱ��," & vbNewLine & _
                    "                     Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI'), a.�걾��λ) As ��Ѫʱ��" & vbNewLine & _
                    "                    From ��Ѫ������Ŀ e, ������ĿĿ¼ b, ����ҽ����¼ c, ����ҽ����¼ a, Table(f_Str2list2([1])) d" & vbNewLine & _
                    "                    Where e.ҽ��id(+) = a.Id And b.Id = c.������Ŀid And (b.�������� = '8' And Nvl(b.ִ�з���, 0) = 0 Or b.�������� = '9') And" & vbNewLine & _
                    "                          c.������� = 'E' And c.���id = a.Id And a.����id = d.C1 And a.��ҳid = d.C2 And a.������� = 'K' And" & vbNewLine & _
                    "                          a.ҽ��״̬ Not In (-1, 2, 4))" & vbNewLine & _
                    "       Select b.Key, b.Id, b.����ʱ��, b.������ * Decode(Upper(a.���㵥λ), 'ML', 1, Nvl(a.����ϵ��, 1)) ������, b.��Ѫʱ��" & vbNewLine & _
                    "       From ������ĿĿ¼ a, ҽ����¼ b" & vbNewLine & _
                    "       Where a.Id = b.������Ŀid)"
            Else
                strSQL = _
                    " Select Key, Id, ����ʱ��, ������, ��Ѫʱ��" & vbNewLine & _
                    " From (With ҽ����¼ As (Select /*+ CARDINALITY(d,10) */" & vbNewLine & _
                    "                     a.����id || '_' || a.�Һŵ� Key, a.Id, Decode(Nvl(e.ҽ��id, 0), 0, a.������Ŀid, e.������Ŀid) ������Ŀid," & vbNewLine & _
                    "                     Decode(Nvl(e.ҽ��id, 0), 0, a.�ܸ�����, e.������) ������, a.����ʱ��," & vbNewLine & _
                    "                     Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI'), a.�걾��λ) As ��Ѫʱ��" & vbNewLine & _
                    "                    From ��Ѫ������Ŀ e, ������ĿĿ¼ b, ����ҽ����¼ c, ����ҽ����¼ a, Table(f_Str2list([1])) d" & vbNewLine & _
                    "                    Where e.ҽ��id(+) = a.Id And b.Id = c.������Ŀid And (b.�������� = '8' And Nvl(b.ִ�з���, 0) = 0 Or b.�������� = '9') And" & vbNewLine & _
                    "                          c.������� = 'E' And c.���id = a.Id And a.�Һŵ� = d.Column_Value And a.������� = 'K' And" & vbNewLine & _
                    "                          a.ҽ��״̬ Not In (-1, 2, 4))" & vbNewLine & _
                    "       Select b.Key, b.Id, b.����ʱ��, b.������ * Decode(Upper(a.���㵥λ), 'ML', 1, Nvl(a.����ϵ��, 1)) ������, b.��Ѫʱ��" & vbNewLine & _
                    "       From ������ĿĿ¼ a, ҽ����¼ b" & vbNewLine & _
                    "       Where a.Id = b.������Ŀid)"
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
        End If
    End If
    
    With vsAudit
        If Val(.TextMatrix(.Row, col_ҽ��ID)) <> 0 Then lngID = Val(.TextMatrix(.Row, col_ҽ��ID))
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If mbln������Ѫ������� Then
                    If Not rsTemp.BOF Then
                        rsTemp.MoveFirst
                        dbl��Ѫ���� = 0
                        dbl24h�� = 0
                        rsTemp.Filter = "Key ='" & rsTmp!Key & "'"
                        Do While Not rsTemp.EOF
                            dbl��Ѫ���� = dbl��Ѫ���� + rsTemp("������")
                            If rsTemp!��Ѫʱ�� <> "" And rsTmp("��Ѫʱ��") <> "" Then
                                If CDate(rsTemp!��Ѫʱ��) > CDate(rsTmp("��Ѫʱ��")) - 1 And CDate(rsTemp!��Ѫʱ��) <= CDate(rsTmp("��Ѫʱ��")) Then dbl24h�� = dbl24h�� + rsTemp("������")
                            ElseIf rsTemp!��Ѫʱ�� & "" = "" And rsTmp("��Ѫʱ��") <> "" Then
                                If CDate(rsTemp!����ʱ��) > CDate(rsTmp("��Ѫʱ��")) - 1 And CDate(rsTemp!����ʱ��) <= CDate(rsTmp("��Ѫʱ��")) Then dbl24h�� = dbl24h�� + rsTemp("������")
                            End If
                            rsTemp.MoveNext
                        Loop
                    End If
                End If
                If tbcSub.Selected.Tag = "�����" Then
                    '�������״̬����
                    If rsTmp!���״̬ <> 1 Then GoTo loopNext
                    '�����û�Ȩ�޹���
                    If mbln������Ѫ������� Then
                        If dbl24h�� < 800 And mintAuditPrivs < 1 Then GoTo loopNext
                        If dbl24h�� >= 800 And dbl24h�� < 1600 And mintAuditPrivs < 1 Then GoTo loopNext
                        If dbl24h�� >= 1600 And mintAuditPrivs < 2 Then GoTo loopNext
                    End If
                ElseIf tbcSub.Selected.Tag = "��ǩ��" Then
                    '�������״̬����
                    If rsTmp!���״̬ <> 7 Then GoTo loopNext
                    '�����û�Ȩ�޹���
                    If dbl24h�� < 800 And mintAuditPrivs < 1 Then GoTo loopNext
                    If dbl24h�� >= 800 And dbl24h�� < 1600 And mintAuditPrivs < 2 Then GoTo loopNext
                    If dbl24h�� >= 1600 And mintAuditPrivs < 3 Then GoTo loopNext
                End If
                .AddItem ""
                .TextMatrix(i, COL_��������) = rsTmp!���� & ""
                .TextMatrix(i, col_��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, COL_����) = rsTmp!���� & ""
                .TextMatrix(i, col_��Ѫʱ��) = rsTmp!��Ѫʱ�� & ""
                .TextMatrix(i, COL_��ʼʱ��) = rsTmp!��ʼʱ�� & ""
                .TextMatrix(i, col_ҽ��ID) = rsTmp!ID & ""
                If Val(rsTmp!ID & "") = lngID And lngID <> 0 Then
                    .Row = i
                End If
                If mbln������Ѫ������� Then
                    .TextMatrix(i, col_24h��Ѫ��) = dbl24h�� & "ml"
                    .TextMatrix(i, col_��Ѫ����) = dbl��Ѫ���� & "ml"
                End If
                If tbcSub.Selected.Tag = "��ǩ��" Then
                    .TextMatrix(i, col_���״̬˵��) = Decode(rsTmp!���״̬ & "", "", "�������", "1", "�����", "2", "���ͨ��", "3", "���δͨ��", "4", "Ѫ�������", "5", "Ѫ����Ѫ��", "6", "Ѫ��ֹͣ��Ѫ", "7", "ѪҺ��ǩ��")
                End If
                .TextMatrix(i, col_���ID) = rsTmp!���ID & ""
                .TextMatrix(i, col_�Ա�) = rsTmp!�Ա� & ""
                .TextMatrix(i, col_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_��Ժʱ��) = rsTmp!��Ժ���� & ""
                .TextMatrix(i, col_����) = rsTmp!���� & ""
                .TextMatrix(i, COL_����ȼ�) = rsTmp!����ȼ� & ""
                .TextMatrix(i, col_����Id) = rsTmp!����ID & ""
                .TextMatrix(i, col_��ҳID) = rsTmp!��ҳID & ""
                .TextMatrix(i, col_�Һŵ�) = rsTmp!�Һŵ� & ""
                .TextMatrix(i, col_���״̬) = Val(rsTmp!���״̬ & "")
                .TextMatrix(i, COL_�������) = Val(rsTmp!������� & "")
                .TextMatrix(i, COL_��������) = Val(rsTmp!�������� & "")
                If optOccasion(1).Value Then
                    If InStr(strPatis, rsTmp!����ID & ":" & rsTmp!�Һŵ�) = 0 Then
                        strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!�Һŵ�
                    End If
                Else
                    If InStr(strPatis, rsTmp!����ID & ":" & rsTmp!��ҳID) = 0 Then
                        strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
                    End If
                End If
                If InStr(strDate, Format(rsTmp!��Ѫʱ�� & "", "YYYY-MM-DD")) = 0 Then
                    strDate = strDate & "," & Format(rsTmp!��Ѫʱ�� & "", "YYYY-MM-DD")
                End If
                .TextMatrix(i, COL_��ID) = rsTmp!��ID & ""
                .TextMatrix(i, COL_�������) = rsTmp!������� & ""
                .TextMatrix(i, COL_������Դ) = rsTmp!������Դ & ""
                .TextMatrix(i, COL_����) = rsTmp!��Ժ���� & ""
                '��ʾ���ģʽ�µ�ҽ������
                strFormat = rsTmp!ҽ������
                blnDo = True
                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                If blnDo Then
                    strTmp = .TextMatrix(i, COL_����)
                    If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                End If
                
                .TextMatrix(i, COL_ҽ������) = strFormat
                If blnIsAudited Then
                    .TextMatrix(i, COL_ǩ��id) = rsTmp!ǩ��id & ""
                    .TextMatrix(i, COL_ҽ��״̬) = rsTmp!ҽ��״̬ & ""
                    
                    intѡ�� = Val(rsTmp!�������� & "") - 10
                    If gblnѪ��ϵͳ Then
                        If InStr(",2,4,5,", "," & Val(.TextMatrix(i, col_���״̬)) & ",") > 0 Then
                            intѡ�� = 1
                        End If
                    End If
                    '���ҽ���������¿�״̬����ı�������ɫ
                    If Val(rsTmp!ҽ��״̬ & "") <> 1 Then
                        .Cell(flexcpForeColor, i, col_ѡ��, i, COL_ǩ��id) = &HC00000
                    End If
                    
                    '����Ѫ��ϵͳ����ı�������ɫ������ɫ
                    If gblnѪ��ϵͳ Then
                        If InStr(",2,5,", "," & Val(.TextMatrix(i, col_���״̬)) & ",") > 0 Then
                            .Cell(flexcpForeColor, i, col_ѡ��, i, COL_ǩ��id) = &H8080FF
                        End If
                    End If
                    
                End If
                .TextMatrix(i, COL_���˵��) = rsTmp!����˵�� & ""
                .TextMatrix(i, COL_���ʱ��) = Format(rsTmp!���ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, COL_�����) = rsTmp!����� & ""
                .TextMatrix(i, col_סԺ��) = rsTmp!סԺ�� & ""
                .TextMatrix(i, COL_��ǰ����) = rsTmp!��ǰ���� & ""
                .TextMatrix(i, COL_����ҽ��) = rsTmp!����ҽ�� & ""
                .TextMatrix(i, COL_����ʱ��) = rsTmp!����ʱ�� & ""
                .TextMatrix(i, COL_��������ID) = rsTmp!��������ID & ""
                .TextMatrix(i, COL_��Ժ����ID) = rsTmp!��Ժ����ID & ""
                .TextMatrix(i, COL_��ǰ����ID) = rsTmp!��ǰ����ID & ""
                .TextMatrix(i, COL_����ĿID) = rsTmp!������ĿID & ""
                .TextMatrix(i, col_��Ѫҽ��) = Val(rsTmp!ִ�з��� & "")
                i = i + 1
loopNext:           rsTmp.MoveNext
                
            Loop
            
            strPatis = Mid(strPatis, 2)
            strDate = Mid(strDate, 2)
            If .Rows = 1 Then .AddItem ""
        Else
            .AddItem ""
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
    Call vsAudit_KeyPress(vbKeyBack)
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If zlCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(col_����).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(gint���� = 0, COL_ƴ������, COL_��ʼ���)).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
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
        If Me.Visible = False Then Exit Sub
        If NewCol = COL_���˵�� And tbcSub.Selected.Tag = "�����" Or NewCol = col_ѡ�� Then
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
            lblInformation(info_����ȼ�).Caption = .TextMatrix(NewRow, COL_����ȼ�)
            lblInformation(info_����).Caption = IIf(Val(.TextMatrix(NewRow, COL_����) & "") = 0, "", .TextMatrix(NewRow, COL_����) & "Kg")
            
            '������¼
            Call LoadPatiAllergy(Val(.TextMatrix(NewRow, col_����Id) & ""), cbo����)
            
            '���
            lblInformation(info_���).Caption = GetPatiDiagnose(Val(.TextMatrix(NewRow, col_����Id) & ""), _
            Val(.TextMatrix(NewRow, col_��ҳID) & ""), _
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
    If Not (Col = COL_���˵��) Then
        Cancel = True
    Else
        With vsAudit
            If .TextMatrix(1, col_ҽ��ID) & "" = "" Or Val(.TextMatrix(.Row, col_���״̬)) = 3 Or _
                    (.TextMatrix(Row, COL_ҽ��״̬) & "" <> "1" And .TextMatrix(Row, COL_ҽ��״̬) & "" <> "") Then
                Cancel = True
            End If
            If gblnѪ��ϵͳ And InStr(",2,5,", "," & Val(.TextMatrix(Row, col_���״̬)) & ",") > 0 Then
                Cancel = True
            End If
        End With
    End If
End Sub

Private Sub vsAudit_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsAudit
        mstrButPri = "00000"
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_ѡ��) = vbChecked Then
                Select Case tbcSub.Selected.Tag
                    Case "�����"
                        mstrButPri = "11000"
                    Case "��ǩ��"
                        mstrButPri = "00110"
                    Case "��ǩ��"
                        mstrButPri = "00001"
                    Case "�����"
                        mstrButPri = "00100"
                End Select
                Exit For
            End If
        Next
        If tbcSub.Selected.Tag = "��ǩ��" And .Col = col_ѡ�� Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpBackColor, i, col_ѡ��) = &HFFC0FF And .Cell(flexcpChecked, i, col_ѡ��) = vbChecked Then
                    mstrButPri = "00100"
                    Exit For
                End If
            Next
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_ѡ��) = vbChecked Then
                tbcSub.Item(0).Enabled = False
                tbcSub.Item(1).Enabled = False
                If mbln������Ѫ������� Then tbcSub.Item(2).Enabled = False
                tbcSub(tbcSub.Selected.Index).Enabled = True
                Exit Sub
            End If
        Next
        tbcSub.Item(0).Enabled = True
        tbcSub.Item(1).Enabled = True
        If mbln������Ѫ������� Then tbcSub.Item(2).Enabled = True
    End With
End Sub

Private Sub vsAudit_Click()
    Call vsAudit_KeyPress(vbKeySpace)
End Sub

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
    Dim lngloop As Long

    With vsAudit
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call UnAuditEnterNextCell
        ElseIf KeyAscii = vbKeyBack Then
            .Cell(flexcpChecked, 0, col_ѡ��) = flexUnchecked
        ElseIf .Col = COL_���˵�� And .Cell(flexcpForeColor, .Row, COL_��������) <> &HFFC0FF And Val(.TextMatrix(.Row, col_���״̬)) <> 3 Then
            .ComboList = "" 'ʹ��ť״̬��������״̬
        ElseIf .Col = col_ѡ�� And KeyAscii = vbKeySpace Then
            If .TextMatrix(1, col_ҽ��ID) = "" Then Exit Sub
            If .MouseRow = .FixedRows - 1 Then
                If .Cell(flexcpChecked, 0, col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, 0, col_ѡ��, .Rows - 1, col_ѡ��) = flexUnchecked
                Else
                    .Cell(flexcpChecked, 0, col_ѡ��, .Rows - 1, col_ѡ��) = flexChecked
                End If
            ElseIf .MouseRow < .Rows Then
                If .Cell(flexcpChecked, .Row, col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, .Row, col_ѡ��) = flexUnchecked
                Else
                    .Cell(flexcpChecked, .Row, col_ѡ��) = flexChecked
                End If
            End If
        End If
        If mbln������Ѫ������� Then
            Select Case tbcSub.Selected.Tag
                Case "�����"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, col_24h��Ѫ��)) < 800 Then
                            If mintAuditPrivs < 1 Then .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                        ElseIf Val(.TextMatrix(lngloop, col_24h��Ѫ��)) >= 1600 Then
                            If mintAuditPrivs < 2 Then .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                        Else
                            If mintAuditPrivs < 1 Then .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                        End If
                    Next
                Case "��ǩ��"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, col_24h��Ѫ��)) < 800 Then
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            End If
                        ElseIf Val(.TextMatrix(lngloop, col_24h��Ѫ��)) >= 1600 Then
                            If mintAuditPrivs < 2 Then
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            ElseIf mintAuditPrivs = 2 Then
                                .Cell(flexcpBackColor, lngloop, col_ѡ��) = &HFFC0FF
                            End If
                        Else
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            ElseIf mintAuditPrivs = 1 Then '����ֻ�ܻ��ˣ�����ǩ��
                                .Cell(flexcpBackColor, lngloop, col_ѡ��) = &HFFC0FF
                            End If
                        End If
                        .Cell(flexcpBackColor, lngloop, COL_���˵��) = .Cell(flexcpBackColor, lngloop, col_ѡ��)
                    Next
                Case "��ǩ��"
                    For lngloop = 1 To .Rows - 1
                        If Val(.TextMatrix(lngloop, COL_ҽ��״̬)) <> 1 Or InStr(1, "'2'3'4'", "'" & .TextMatrix(lngloop, col_���״̬) & "'") > 1 Then '(Val(.TextMatrix(lngloop, col_���״̬)) <> 4 And Val(.TextMatrix(lngloop, col_���״̬)) <> 3) Then
                            .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            .Cell(flexcpBackColor, lngloop, col_ѡ��, lngloop, col_24h��Ѫ��) = &H80000016
                        ElseIf Val(.TextMatrix(lngloop, col_24h��Ѫ��)) < 800 Then
                            If mintAuditPrivs < 1 Then
                                .Cell(flexcpBackColor, lngloop, col_ѡ��) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            End If
                        ElseIf Val(.TextMatrix(lngloop, col_24h��Ѫ��)) >= 1600 Then
                            If mintAuditPrivs < 3 Then
                                .Cell(flexcpBackColor, lngloop, col_ѡ��) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            End If
                        Else
                            If mintAuditPrivs < 2 Then
                                .Cell(flexcpBackColor, lngloop, col_ѡ��) = &H80000016
                                .Cell(flexcpChecked, lngloop, col_ѡ��) = flexUnchecked
                            End If
                        End If
                        .Cell(flexcpBackColor, lngloop, COL_���˵��) = .Cell(flexcpBackColor, lngloop, col_ѡ��)
                    Next
            End Select
        End If
    End With
End Sub

Private Sub vsAudit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = COL_���˵�� Then
        If zlCommFun.ActualLen(vsAudit.Editable) - zlCommFun.ActualLen(vsAudit.EditSelText) >= 100 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
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
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAudit_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
        
    With vsAudit
        If tbcSub.Selected.Tag = "��ǩ��" And .Col = col_ѡ�� Then
            mstrButPri = "1"
            For i = 1 To .Rows - 1
                If .Cell(flexcpBackColor, i, col_ѡ��) = &HC0FFFF And .Cell(flexcpChecked, i, col_ѡ��) = vbChecked Then
                    mstrButPri = "0"
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsAudit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then mobjPopup.ShowPopup
End Sub

Private Sub vsAudit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = COL_���˵�� Then
        vsAudit.EditSelStart = 0
        vsAudit.EditSelLength = Len(vsAudit.EditText)
    End If
End Sub


