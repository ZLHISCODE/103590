VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCriticalQuery 
   Caption         =   "Σ��ֵ��ѯ"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18090
   Icon            =   "frmCriticalQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   18090
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPatiC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   4050
      ScaleHeight     =   540
      ScaleWidth      =   1575
      TabIndex        =   34
      Top             =   5550
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   150
      ScaleHeight     =   2535
      ScaleWidth      =   4020
      TabIndex        =   2
      Top             =   2490
      Width           =   4020
      Begin VB.ComboBox cboState 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1275
         Width           =   2040
      End
      Begin VB.PictureBox picLX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1245
         ScaleHeight     =   285
         ScaleWidth      =   2295
         TabIndex        =   30
         Top             =   1560
         Width           =   2295
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   35
            Top             =   45
            Width           =   795
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "סԺ"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   32
            Top             =   45
            Width           =   795
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   31
            Top             =   45
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ˢ��"
         Height          =   300
         Left            =   3075
         TabIndex        =   24
         Top             =   930
         Width           =   615
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   915
         Width           =   2055
      End
      Begin VB.ComboBox cboRegDept 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   510
         Width           =   2055
      End
      Begin VB.ComboBox cboPatiDept 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Width           =   2055
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¼״̬"
         Height          =   180
         Left            =   165
         TabIndex        =   36
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lblLX 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblRegTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   975
         Width           =   720
      End
      Begin VB.Label lblRegDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽǿ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   5
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblPatiDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ�Ͽ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   3
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Timer timeRefreshCard 
      Interval        =   1000
      Left            =   3495
      Top             =   6390
   End
   Begin VB.PictureBox picCItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2190
      Index           =   0
      Left            =   11010
      ScaleHeight     =   2190
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   5445
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label lblAge 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   28
         Top             =   810
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblSex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   27
         Top             =   810
         Width           =   360
      End
      Begin VB.Label lblTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   25
         Top             =   510
         Width           =   450
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   1095
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblSelect 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   435
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCardFra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   11805
      ScaleHeight     =   2730
      ScaleWidth      =   4275
      TabIndex        =   19
      Top             =   4260
      Width           =   4275
      Begin VB.PictureBox picCardCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1470
         ScaleHeight     =   1500
         ScaleWidth      =   2160
         TabIndex        =   29
         Top             =   375
         Width           =   2160
      End
      Begin VB.VScrollBar vscH 
         Height          =   2625
         LargeChange     =   10
         Left            =   3975
         Max             =   100
         SmallChange     =   5
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   250
      End
   End
   Begin VB.Frame fraPati 
      Height          =   1590
      Left            =   4695
      TabIndex        =   11
      Top             =   1215
      Width           =   8265
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   210
         ScaleHeight     =   1245
         ScaleWidth      =   7395
         TabIndex        =   12
         Top             =   150
         Width           =   7395
         Begin VB.Image imgPatient 
            Height          =   705
            Left            =   75
            Picture         =   "frmCriticalQuery.frx":6852
            Stretch         =   -1  'True
            Top             =   210
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1365
            TabIndex        =   17
            Top             =   195
            Width           =   600
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   5445
            TabIndex        =   16
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�Ա�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3375
            TabIndex        =   15
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ʶ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   1485
            TabIndex        =   14
            Top             =   855
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   13
            Top             =   810
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox picCritical 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   6030
      ScaleHeight     =   2310
      ScaleWidth      =   4935
      TabIndex        =   1
      Top             =   2865
      Width           =   4935
      Begin VSFlex8Ctl.VSFlexGrid vsCritical 
         Bindings        =   "frmCriticalQuery.frx":771C
         Height          =   1395
         Left            =   435
         TabIndex        =   10
         Top             =   390
         Width           =   4000
         _cx             =   7056
         _cy             =   2461
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1935
      Left            =   4665
      TabIndex        =   0
      Top             =   2835
      Width           =   1395
      _Version        =   589884
      _ExtentX        =   2461
      _ExtentY        =   3413
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   7785
      Width           =   18090
      _ExtentX        =   31909
      _ExtentY        =   635
      SimpleText      =   $"frmCriticalQuery.frx":7730
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCriticalQuery.frx":7777
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   26829
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
   Begin VB.Image imgWJ 
      Height          =   240
      Index           =   0
      Left            =   9945
      Picture         =   "frmCriticalQuery.frx":800B
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCL 
      Height          =   240
      Index           =   0
      Left            =   9570
      Picture         =   "frmCriticalQuery.frx":E85D
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   0
      Left            =   7050
      Picture         =   "frmCriticalQuery.frx":150AF
      Top             =   5385
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   1
      Left            =   8745
      Picture         =   "frmCriticalQuery.frx":192C8
      Top             =   6045
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgDefual 
      Height          =   705
      Left            =   1770
      Picture         =   "frmCriticalQuery.frx":1CC8F
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLoad 
      Height          =   705
      Left            =   1425
      Picture         =   "frmCriticalQuery.frx":1DB59
      Stretch         =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   405
      Top             =   195
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCriticalQuery.frx":1EA23
      Left            =   405
      Top             =   1125
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCriticalQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PatiCol
    COL_����ID = 0
    COL_��ҳID
    COL_�Һ�ID
    COL_�Һŵ�
    COL_����
    COL_�Ա�
    COL_�����
    COL_סԺ��
    COL_����
    COL_����
    COL_����
End Enum

Private Enum AdviceCol
    colΣ��ֵ����
    col����ʱ��
    col������
    col�������
    colȷ��ʱ��
    colȷ����
    colȷ�Ͽ���
    col���
    
    
    '������
    colID
    col״̬
    colҽ��ID
End Enum

Private Enum e_Ctrl
    e���� = 0
    e�Ա�
    e����
    e��ʶ��
    e����
End Enum

Private Const conMenu_View_AppCritical = 200
Private Const clngX = 100 '���Ͻǵ�һ�ſ�Ƭλ��
Private mobjCISJob As Object
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣ����
Private mlngModul As Long
Private mstrPrivs As String
Private mfrmParent As Object
Private mblnModal As Boolean '��ʾ��ʽ��ģ̬����ģ̬
Private mint��ʽ As Integer '0-�������˲�ѯ��1-�����ѯ
Private mint����  As Integer '0-���1-סԺ��2-�����סԺ��3-����̨������ѯ����
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String
Private mlng����ID As Long
Private mlng����ID As Long '�������ID
Private mlng����ID As Long '����ID
Private mint���� As Integer '0-ҽ��վ,1-ҽ��վ
Private mblnOK As Boolean

Private mrsCard As ADODB.Recordset '������Ϣ
Private mlngCntCard As Long '�ܼ�¼��
Private mblnRefreshCard As Boolean
Private mintCurIndex As Integer '��ǰѡ��Ŀ�Ƭ�±�
Private mlngPreRowCnt As Long 'ǰһ������һ���еĿ�Ƭ����
Private mstrPrePati As String 'ǰһ������
Private mlngPreCardID As Long

Private mintPreTim As Integer
Private mdatB�Ǽ� As Date
Private mdatE�Ǽ� As Date

Private mint��ʾ��ʽ As Integer '0-������ѯ��1-��Ƭѡ����
Private mlng��¼ID As Long

Public Function ShowMe(frmParent As Object, ByVal blnModal As Boolean, ByVal int���� As Integer, ByVal int���� As Integer, ByVal lng����id As Long, ByVal lng����ID As Long, ByRef objMip As Object) As Boolean
'���ܣ���ʾ����
'������frmParent ������ ��blnModal ������ʾģʽ��false-��ģ̬��true-ģ̬
'      int���� �������� 0-���1-סԺ��2-�����סԺ��3-����̨������ѯ����
'      int���� 0-ҽ��վ,1-ҽ��վ
'      lng����ID ҽ��վ����ʱҽ�����ң�ҽ��վ����ʱ�����˿���ID,
'      lng����ID סԺҽ��վ��������ʾʱ������ѡ��Ĳ���ID
'      objMip ���ڷ�����Ϣ�Ķ��� zl9ComLib.clsMipModule
    Set mfrmParent = frmParent
    mint��ʾ��ʽ = 0
    mblnModal = blnModal
    mint���� = int����
    mlng����ID = lng����id
    mint���� = int����
    mlng����ID = lng����ID
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Public Function ShowMeQuery(ByVal lngSys As Long, ByVal lngModul As Long, ByRef frmParent As Object, ByVal strPrivs As String)
'���ܣ�������ѯ����
    mlngModul = lngModul
    mstrPrivs = strPrivs
    mint��ʾ��ʽ = 0
    mint���� = 3
    Set mfrmParent = frmParent
    Me.Show , frmParent
    ShowMeQuery = mblnOK
End Function

Public Function ShowMeSelCard(frmParent As Object, ByVal rsIn As ADODB.Recordset) As Long
'���ܣ���Ƭѡ����ģʽ
'������rsIn Ҫ���صļ�¼��
'���أ�Σ��ֵ��¼ID
    Set frmParent = frmParent
    mint��ʾ��ʽ = 1
    mlng��¼ID = 0
    Set mrsCard = zldatabase.CopyNewRec(rsIn)
    Me.Show 1, frmParent
    ShowMeSelCard = mlng��¼ID
End Function

Private Sub cboRegDept_Click()
'���ܣ��л�����
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Dim objControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIF(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
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
    Case conMenu_View_AppCritical '�鿴����
        Call EditData(2)
    Case conMenu_Edit_Modify '�޸�
        Call EditData(1)
    Case conMenu_Edit_Delete 'ɾ����¼
        Call DeleteData
    Case conMenu_Edit_Send
        Call FunAffirm
    Case conMenu_View_Refresh
        Call LoadPatients
    Case conMenu_Tool_Archive
        If mlng����ID <> 0 Then
            Call mobjCISJob.ShowArchive(Me, mlng����ID, mlng����ID)
        End If
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
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
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_Edit_Modify, conMenu_Edit_Delete 'ɾ����¼
        Control.Enabled = CanEdit()
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(1259) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_Edit_Send 'Σ��ֵȷ�ϣ�ֻ���������˿���
        Control.Enabled = False
        If mintCurIndex > 0 Then
            mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
            If Not mrsCard.EOF Then
                If Val(mrsCard!��ҳID & "") = 0 And mrsCard!�Һŵ� & "" = "" Then
                    Control.Enabled = True
                End If
            End If
        End If
    End Select
End Sub

Private Function CanEdit() As Boolean
'���ܣ���ǰѡ���Σ��ֵ��¼�Ƿ���Ա༭
    Dim strTmp As String
    Dim strTag As String
    Dim blnEdit  As Boolean
    
    strTag = tbcSub.Selected.Tag
    Select Case strTag
    Case "����"
    Case "Σ��ֵ"
        With vsCritical
            If Val(.TextMatrix(.Row, col״̬)) = 1 And Val(.TextMatrix(.Row, colID)) <> 0 Then
                blnEdit = True
            Else
                blnEdit = False
            End If
        End With
    Case "��ϸ��"
        If mintCurIndex > 0 Then
            blnEdit = Not imgCL(mintCurIndex).Visible
        End If
    End Select
    CanEdit = blnEdit
End Function

Private Sub cmdOK_Click()
    Call LoadPatients
End Sub

Private Sub Form_Load()
    Dim intIdx As Integer
    Dim objPane As Pane
        
    If mint��ʾ��ʽ = 0 Then
        Call RestoreWinState(Me, App.ProductName)
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
        If mint���� <> 2 Then
            'DockingPane
            '-----------------------------------------------------
            Me.dkpMain.SetCommandBars Me.cbsMain
            Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
            Me.dkpMain.Options.ThemedFloatingFrames = True
            Me.dkpMain.Options.AlphaDockingContext = True
            Set objPane = Me.dkpMain.CreatePane(1, 350, 400, DockLeftOf, Nothing)
            objPane.Title = "��������"
            objPane.Options = PaneNoCloseable Or PaneNoFloatable
        End If
        
        With Me.tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .ClientFrame = xtpTabFrameSingleLine
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
            
            .InsertItem(intIdx, "�б�", picCritical.Hwnd, 0).Tag = "Σ��ֵ": intIdx = intIdx + 1
            .InsertItem(intIdx, "��Ƭ", picCardFra.Hwnd, 0).Tag = "��ϸ��": intIdx = intIdx + 1
            .Item(1).Selected = True

            If mint���� = 2 Then
                .Item(0).Visible = False
            End If
             
        End With
        
        If mint���� = 3 Then
            Call Init�Ǽǿ���
            Call Initȷ�Ͽ���
            
            Set mobjCISJob = CreateObject("zl9CISJob.clsCISJob")
        End If
        mintPreTim = -1
        With cboSelectTime
            .Clear
            .AddItem "������"
            .ItemData(.NewIndex) = 0
            .AddItem "������"
            .ItemData(.NewIndex) = 1
            .AddItem "ǰ����"
            .ItemData(.NewIndex) = 2
            .AddItem "һ����"
            .ItemData(.NewIndex) = 7
            .AddItem "30����"
            .ItemData(.NewIndex) = 30
            .AddItem "60����"
            .ItemData(.NewIndex) = 60
            .AddItem "[ָ��...]"
            .ItemData(.NewIndex) = -1
        End With
        cboSelectTime.ListIndex = 0
        
        
        With cboState
            .Clear
            .AddItem "ȫ��״̬"
            .AddItem "δȷ��"
            .AddItem "ȷ��Ϊ��Σ��ֵ"
            .AddItem "ȷ��Ϊ��Σ��ֵ"
            .ListIndex = 0
        End With
        
        
        mblnOK = False
        mintCurIndex = -1
        Call SetFaceCtrl
        Call SetFilterInfo
        Call InitTable
        Call MainDefCommandBar
         
        Call LoadPatients
    ElseIf mint��ʾ��ʽ = 1 Then
        Me.BorderStyle = 3
        Me.Caption = "Σ��ֵѡ��(˫��ѡ��)"
        Me.Width = 5800
        Me.Height = 4900
        
        Call ShowAllCard
        
        '���ڹ���������ʾ
        If picCardCon.Height < picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 Then
            vscH.Visible = True
            vscH.value = 0
        Else
            vscH.Visible = False
        End If
        
        picCardCon.Height = picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100
    
    End If
End Sub

Private Sub cboSelectTime_Click()
 
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintPreTim And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdatB�Ǽ�, mdatE�Ǽ�, cboSelectTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call Cbo.SetIndex(cboSelectTime.Hwnd, mintPreTim)
            Exit Sub
        End If
    Else
        mdatE�Ǽ� = datCurr
        mdatB�Ǽ� = mdatE�Ǽ� - intDateCount
    End If
    If mdatB�Ǽ� = CDate(0) Or mdatE�Ǽ� = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "��Χ��" & Format(mdatB�Ǽ�, "yyyy-MM-dd") & " �� " & Format(mdatE�Ǽ�, "yyyy-MM-dd")
    End If
    mintPreTim = cboSelectTime.ListIndex
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picFilter.Hwnd
    ElseIf Item.ID = 2 Then
'        Item.Handle = picPati.Hwnd
    End If
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strFunName As String

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "�鿴Σ��ֵ��(&D)")
            objControl.IconId = 3031
        If mint���� = 2 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "Σ��ֵȷ��")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
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
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "�鿴Σ��ֵ��(&D)")
            objControl.IconId = 3031
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "�鿴Σ��ֵ��") '����
            objControl.IconId = 3031
            objControl.Style = xtpButtonIconAndCaption
            
        If mint���� = 2 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "Σ��ֵȷ��")
                objControl.Style = xtpButtonIconAndCaption
        End If
        
        If mint���� = 3 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
                objControl.BeginGroup = True
                objControl.Style = xtpButtonIconAndCaption
        End If
            
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
            objControl.Style = xtpButtonIconAndCaption
            objControl.IconId = 191
            objControl.BeginGroup = True
    End With
     
    objControl.Style = xtpButtonIconAndCaption
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngH As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    
    If mint���� = 2 Then
        picFilter.Top = lngTop
        picFilter.Height = 400
        picFilter.Width = lngRight - lngLeft
        picFilter.Left = lngLeft
        With Me.fraPati
            .Left = lngLeft: .Top = lngTop - 60 + 400
            .Width = lngRight - lngLeft
        End With
        With Me.tbcSub
            .Left = lngLeft: .Width = lngRight - lngLeft
            .Top = lngTop + fraPati.Height: .Height = lngBottom - lngTop - fraPati.Height - IIF(Me.stbThis.Visible, stbThis.Height, 0)
        End With
    ElseIf mint���� = 3 Then
        lngH = 1150
        picPatiC.Move lngLeft, lngTop, lngRight - lngLeft, lngH
        fraPati.Move 0, -60, picPatiC.Width, lngH + 60
        
        
        With Me.tbcSub
            .Left = lngLeft: .Width = lngRight - lngLeft
            .Top = lngTop + picPatiC.Height: .Height = lngBottom - lngTop - picPatiC.Height - IIF(Me.stbThis.Visible, stbThis.Height, 0)
        End With
        
    End If
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    imgPatient.Top = 10
    imgPatient.Left = 10
    imgPatient.Height = 1000
    
    picInfo.Left = 30
    picInfo.Top = 100
    picInfo.Width = fraPati.Width - 130
    picInfo.Height = imgPatient.Height + 30
    fraPati.Height = 1200
    
    lblInfo(e��ʶ��).Left = lblInfo(e����).Left
    lblInfo(e����).Left = lblInfo(e�Ա�).Left
    lblInfo(e��ʶ��).Top = 800
    lblInfo(e����).Top = 800
    
    If mint��ʾ��ʽ = 1 Then
        picCardFra.Move 0, 0, Me.Width - 80, Me.Height - 430
        picCardFra.ZOrder 0
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mlng��ҳID = 0
    mstr�Һŵ� = ""
    mlng����ID = 0
    Call UnloadControls
    Set mrsCard = Nothing
    mstrPrePati = ""
    Set mobjCISJob = Nothing
    If mint��ʾ��ʽ = 0 Then
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Sub lblText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, lblText(Index).Caption, True
End Sub

Private Sub imgCL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, imgCL(Index).Tag, True
End Sub

Private Sub imgWJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, imgWJ(Index).Tag, True
End Sub

Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblAge_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblSex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblTime_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgCL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgWJ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub


Private Sub lblName_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblAge_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblSex_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblText_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub imgCL_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub imgWJ_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub optInfo_Click(Index As Integer)
    If Index = 0 Then '����
    
    ElseIf Index = 1 Then 'סԺ
    
    ElseIf Index = 2 Then '����
    
    End If
End Sub

Private Sub picCardFra_Resize()
    On Error Resume Next
    picCardCon.Move 0, 0, picCardFra.Width - vscH.Width, picCardFra.Height
    vscH.Left = picCardCon.Width
    vscH.Top = 0
    vscH.Height = picCardFra.Height
 
    '���ý��濨Ƭ��Ӧ
    Call ReSetCardPos
End Sub

Private Sub picCItem_DblClick(Index As Integer)
'���ܣ���ʾ��Ƭ
    Call ShowCardPop
End Sub

Private Sub picCItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
    If mintCurIndex > 0 Then
        '�����һ����ѡ��
        lblSelect(mintCurIndex).Visible = False
    End If
    mintCurIndex = Index
    Call ShowSelect
End Sub

Private Sub picCItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, ""
End Sub

Private Sub picCritical_Resize()
    On Error Resume Next
    vsCritical.Move 0, 0, picCritical.Width, picCritical.Height
End Sub

Private Sub picFilter_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    If mint���� = 2 Then
        'ҽ��վ����
        lblRegDept.Top = 120
        lblRegDept.Left = 60
        
        lblRegTime.Left = 60
        lblRegTime.Top = 450
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegDept, 1200, lblRegTime, 60, cboSelectTime, 80, cmdOK)
    ElseIf mint���� = 3 Then
        lblLX.Top = 120
        lblLX.Left = 60
        
        lngTmp = 200
        
        Call zlControl.SetPubCtrlPos(True, 0, lblLX, lngTmp, lblPatiDept, lngTmp, lblRegDept, lngTmp, lblState, lngTmp, lblRegTime)
 
        Call zlControl.SetPubCtrlPos(False, 0, lblLX, 100, picLX)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblPatiDept, 100, cboPatiDept)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegDept, 100, cboRegDept)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblState, 100, cboState)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegTime, 100, cboSelectTime, 80, cmdOK)
        
    End If
End Sub
 
Private Function LoadPatients() As Boolean
'���ܣ����ز����б�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
    Dim i As Long
    Dim datETmp As Date
    Dim strWhere As String
    Dim lngȷ�Ͽ���ID As Long
    Dim lng�Ǽǿ���ID As Long
    
    On Error GoTo errH
    
    datETmp = Format(mdatE�Ǽ�, "yyyy-MM-dd 23:59:59")
    
    If mint���� = 2 Then
        strSql = "select rownum as ���,a.id,a.����id,a.��ҳid,a.�Һŵ�,a.ҽ��ID,a.״̬,a.�Ƿ�Σ��ֵ,a.����,a.Σ��ֵ����,a.�Ա�,a.����,a.����ʱ�� from ����Σ��ֵ��¼ a" & _
            " where a.�������id = [1] And a.����ʱ�� Between [2] And [3] order by a.����ʱ�� desc "
        Set mrsCard = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mdatB�Ǽ�, datETmp)
        mblnRefreshCard = True
    End If
    
    If mint���� = 3 Then
        If optInfo(0).value Then
            strWhere = " and a.�Һŵ� is not null "
        ElseIf optInfo(1).value Then
            strWhere = " and nvl(a.��ҳid,0)>0 "
        ElseIf optInfo(2).value Then
            strWhere = " and nvl(a.��ҳid,0)=0 and  a.�Һŵ�  is  null"
        End If
        
        If cboPatiDept.ListIndex >= 0 Then
            'ȷ�Ͽ���
            If cboPatiDept.ItemData(cboPatiDept.ListIndex) <> 0 Then
                lngȷ�Ͽ���ID = cboPatiDept.ItemData(cboPatiDept.ListIndex)
                strWhere = strWhere & " and a.ȷ�Ͽ���ID =[1] "
            End If
        End If
        
        If cboRegDept.ListIndex >= 0 Then
            '�Ǽǿ���
            If cboRegDept.ItemData(cboRegDept.ListIndex) <> 0 Then
                lng�Ǽǿ���ID = cboRegDept.ItemData(cboRegDept.ListIndex)
                strWhere = strWhere & " and a.�������ID =[2] "
            End If
        End If
        
        
        If cboState.ListIndex >= 0 Then
            Select Case cboState.ListIndex
            Case 0
            Case 1
                strWhere = strWhere & " and a.״̬=1 "
            Case 2
                strWhere = strWhere & " and a.״̬=2 and nvl(a.�Ƿ�Σ��ֵ,0)=0 "
            Case 3
                strWhere = strWhere & " and a.״̬=2 and nvl(a.�Ƿ�Σ��ֵ,0)=1 "
            End Select
        End If
        
        
        strSql = "select rownum as ���,a.id,a.����id,a.��ҳid,a.�Һŵ�,a.ҽ��ID,a.״̬,a.�Ƿ�Σ��ֵ,a.����,a.Σ��ֵ����,a.�Ա�,a.����,a.����ʱ�� from ����Σ��ֵ��¼ a" & _
            " where a.����ʱ�� Between [3] And [4] " & strWhere & " order by a.����ʱ�� desc "
        Set mrsCard = zldatabase.OpenSQLRecord(strSql, Me.Caption, lngȷ�Ͽ���ID, lng�Ǽǿ���ID, mdatB�Ǽ�, datETmp)
        mblnRefreshCard = True
        
        Call LoadCritical
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadCritical() As Boolean
'���ܣ�����Σ��ֵ�б�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim str��� As String
    Dim lngȷ�Ͽ���ID As Long
    Dim lng�Ǽǿ���ID As Long
    Dim datETmp As Date
    Dim strWhere As String
    
    On Error GoTo errH
    If mint���� <> 3 Then Exit Function
    
    datETmp = Format(mdatE�Ǽ�, "yyyy-MM-dd 23:59:59")
    
    If optInfo(0).value Then
        strWhere = " and a.�Һŵ� is not null "
    ElseIf optInfo(1).value Then
        strWhere = " and nvl(a.��ҳid,0)>0 "
    ElseIf optInfo(2).value Then
        strWhere = " and nvl(a.��ҳid,0)=0 and  a.�Һŵ�  is  null"
    End If
    
    If cboPatiDept.ListIndex >= 0 Then
        'ȷ�Ͽ���
        If cboPatiDept.ItemData(cboPatiDept.ListIndex) <> 0 Then
            lngȷ�Ͽ���ID = cboPatiDept.ItemData(cboPatiDept.ListIndex)
            strWhere = strWhere & " and a.ȷ�Ͽ���ID =[1] "
        End If
    End If
    
    If cboRegDept.ListIndex >= 0 Then
        '�Ǽǿ���
        If cboRegDept.ItemData(cboRegDept.ListIndex) <> 0 Then
            lng�Ǽǿ���ID = cboRegDept.ItemData(cboRegDept.ListIndex)
            strWhere = strWhere & " and a.�������ID =[2] "
        End If
    End If
    
    If cboState.ListIndex >= 0 Then
        Select Case cboState.ListIndex
        Case 0
        Case 1
            strWhere = strWhere & " and a.״̬=1 "
        Case 2
            strWhere = strWhere & " and a.״̬=2 and nvl(a.�Ƿ�Σ��ֵ,0)=0 "
        Case 3
            strWhere = strWhere & " and a.״̬=2 and nvl(a.�Ƿ�Σ��ֵ,0)=1 "
        End Select
    End If
    
    strSql = "select  a.id,a.Σ��ֵ����,a.����ʱ��,a.������,a.�������,a.ȷ��ʱ��,a.ȷ����,a.ȷ�Ͽ���id,b.���� as ȷ�Ͽ���,a.״̬,a.ҽ��id,a.�Ƿ�Σ��ֵ  from ����Σ��ֵ��¼ a,���ű� b" & _
        " where a.ȷ�Ͽ���id=b.id(+) and  a.����ʱ�� Between [3] And [4] " & strWhere & " order by a.����ʱ�� desc "
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lngȷ�Ͽ���ID, lng�Ǽǿ���ID, mdatB�Ǽ�, datETmp)
 
    With vsCritical
        .Redraw = flexRDNone
        .Rows = 1
        .ExplorerBar = 7
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(i, colΣ��ֵ����) = rsTmp!Σ��ֵ���� & ""
                .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col������) = rsTmp!������ & ""
                .TextMatrix(i, col�������) = rsTmp!������� & ""
                
                If Not IsNull(rsTmp!ȷ��ʱ��) Then
                    .TextMatrix(i, colȷ��ʱ��) = Format(rsTmp!ȷ��ʱ��, "yyyy-MM-dd HH:mm")
                End If
                
                .TextMatrix(i, colȷ����) = rsTmp!ȷ���� & ""
                .TextMatrix(i, colȷ�Ͽ���) = rsTmp!ȷ�Ͽ��� & ""
                
                If Val(rsTmp!״̬ & "") = 2 Then
                    If Val(rsTmp!�Ƿ�Σ��ֵ & "") = 1 Then
                        .TextMatrix(i, col���) = "��Σ��ֵ"
                    Else
                        .TextMatrix(i, col���) = "����Σ��ֵ"
                    End If
                End If
                    
                .TextMatrix(i, colID) = Val(rsTmp!ID & "")
                .TextMatrix(i, col״̬) = Val(rsTmp!״̬ & "")
                .TextMatrix(i, colҽ��ID) = Val(rsTmp!ҽ��ID & "")
                i = i + 1
                rsTmp.MoveNext
            Loop
        Else
            .AddItem ""
        End If
        .Redraw = flexRDDirect
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitTable()
    Dim arrHead As Variant, i As Long
    Dim strHead As String
    strHead = "Σ��ֵ����,2500,1;����ʱ��,1800,1;������,700,1;�������,2000,1;ȷ��ʱ��,1800,1;ȷ����,700,1;ȷ�Ͽ���,800,1;���,800,1;ID;״̬;ҽ��ID"
    arrHead = Split(strHead, ";")
    With vsCritical
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub
 
Private Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = sys.Readlob(glngSys, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'���ܣ����������Ϣ
    lblInfo(e����).Caption = "����"
    lblInfo(e�Ա�).Caption = "�Ա�"
    lblInfo(e����).Caption = "����"
    lblInfo(e��ʶ��).Caption = "��ʶ��"
    lblInfo(e����).Caption = "����"
    imgPatient.Picture = imgDefual.Picture
End Sub

Private Sub vscH_Change()
'
    Dim lngMove As Long
    Dim lngY As Long
    If Not vscH.Visible Then Exit Sub
    '���㵥������
    lngMove = CLng((picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 - picCardFra.Height) / 100)

    If lngMove < 0 Then lngMove = 0
    lngY = -1 * vscH.value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 0
    
    picCardCon.Top = lngY
    
End Sub
 
Private Sub vsCritical_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsCritical
        If Val(.TextMatrix(NewRow, colID)) <> 0 Then
            For i = 1 To mlngCntCard
                If Val(.TextMatrix(NewRow, colID)) = Val(lblName(i).Tag) Then
                    
                    If mintCurIndex > 0 Then
                        '�����һ����ѡ��
                        lblSelect(mintCurIndex).Visible = False
                    End If
                    mintCurIndex = i
                    Call ShowSelect
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsCritical_DblClick()
'���ܣ�˫��Ϊ�鿴Σ��ֵ��
    Dim i As Long
    
    With vsCritical
        If Val(.TextMatrix(.Row, colID)) <> 0 Then
            For i = 1 To mlngCntCard
                If Val(.TextMatrix(.Row, colID)) = Val(lblName(i).Tag) Then
                    
                    If mintCurIndex > 0 Then
                        '�����һ����ѡ��
                        lblSelect(mintCurIndex).Visible = False
                    End If
                    mintCurIndex = i
                    Call ShowSelect
                    Exit For
                End If
            Next
        End If
    End With
    
    Call ShowCardPop
    
'    Dim lng��¼ID As Long
'    Dim lngҽ��ID As Long
'    Dim int�������� As Integer
'    Dim lng����ID As Long
'    Dim lng��ҳID As Long
'    Dim str�Һŵ� As String
'    Dim strΣ��ָ�� As String
'    Dim strΣ����� As String
'
'    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '���������
'    With rptPati.SelectedRows(0)
'        If Not .GroupRow Then
'            lng����ID = Val(.Record(COL_����ID).value)
'            lng��ҳID = Val(.Record(COL_��ҳID).value)
'            str�Һŵ� = .Record(COL_�Һŵ�).value
'        End If
'    End With
'
'    If lng����ID = 0 Then
'        MsgBox "��ѡ��һ�����ˡ�", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    With vsCritical
'        lng��¼ID = Val(.TextMatrix(.Row, colID))
'        lngҽ��ID = Val(.TextMatrix(.Row, colҽ��ID))
'    End With
'    If lng��¼ID = 0 Then
'        MsgBox "��ѡ��һ��Σ��ֵ��¼��", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'
'    If str�Һŵ� = "" Then
'        int�������� = 2
'    Else
'        int�������� = 1
'    End If
'
'    Call frmCriticalEdit.ShowMe(Me, True, 2, int��������, lng����ID, lng��ҳID, str�Һŵ�, 0, lng��¼ID, lngҽ��ID)
End Sub

Private Sub DeleteData()
'���ܣ�ɾ��Σ��ֵ��¼
    Dim strSql As String
    Dim lngID As Long
    
    Select Case tbcSub.Selected.Tag
    Case "Σ��ֵ"
    
    lngID = Val(vsCritical.TextMatrix(vsCritical.Row, colID))
    
    strSql = "zl_����Σ��ֵ��¼_delete(" & lngID & ")"
    Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    Call vsCritical.RemoveItem(vsCritical.Row)
    
    Case "��ϸ��"
        lngID = Val(lblName(mintCurIndex).Tag)
        strSql = "zl_����Σ��ֵ��¼_delete(" & lngID & ")"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
        Call LoadPatients
        Call ShowAllCard
    End Select
    mblnOK = True
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub EditData(ByVal intType As Integer)
'���ܣ��޸Ļ��߲鿴��¼
'������intType 1-�޸ģ�2-�鿴
    If tbcSub.Selected.Tag = "��ϸ��" Or tbcSub.Selected.Tag = "Σ��ֵ" Then
        Call ShowCardBybnt(intType)
        Exit Sub
    End If
End Sub

Private Sub ShowAllCard()
'���ܣ���ʾ��Ƭ
    Dim i As Long
    
    mintCurIndex = -1
    mlngCntCard = mrsCard.RecordCount
    
    Call LoadAllCard
    Call LocatePati
  
    stbThis.Panels(2).Text = "һ��" & mlngCntCard & "��Σֵ��Ϣ��"
End Sub

Private Sub UnloadControls()
'���ܣ�ж�ؿؼ�
    Dim i As Long
    Dim lngCnt As Long
    
    lngCnt = picCItem.Count - 1
    
    For i = lngCnt To 1 Step -1
        Unload imgWJ(i)
        Unload imgCL(i)
        
        Unload lblName(i)
        Unload lblAge(i)
        Unload lblSex(i)
        Unload lblTime(i)
        Unload lblSelect(i)
        Unload lblText(i)
        Unload picCItem(i)
    Next
    
End Sub

Private Sub LoadOneCard(ByVal lngIdx As Long, ByVal lngX As Long, lngY As Long)
'���ܣ�����һ�ſ�Ƭ

    Load picCItem(lngIdx)
    Set picCItem(lngIdx).Container = picCardCon
    picCItem(lngIdx).Visible = True
    picCItem(lngIdx).Picture = imgCardBack(1).Picture
    picCItem(lngIdx).Top = lngY
    picCItem(lngIdx).Left = lngX
    
    Load lblText(lngIdx)
    Set lblText(lngIdx).Container = picCItem(lngIdx)
    lblText(lngIdx).Visible = True
    
    Load lblName(lngIdx)
    Set lblName(lngIdx).Container = picCItem(lngIdx)
    lblName(lngIdx).Visible = True
    
    
    Load lblAge(lngIdx)
    Set lblAge(lngIdx).Container = picCItem(lngIdx)
    lblAge(lngIdx).Visible = True
    
    Load lblSex(lngIdx)
    Set lblSex(lngIdx).Container = picCItem(lngIdx)
    lblSex(lngIdx).Visible = True
    
    
    Load lblTime(lngIdx)
    Set lblTime(lngIdx).Container = picCItem(lngIdx)
    lblTime(lngIdx).Visible = True
    
     
    Load lblSelect(lngIdx)
    Set lblSelect(lngIdx).Container = picCItem(lngIdx)
    lblSelect(lngIdx).Visible = False
        
    Load imgWJ(lngIdx)
    Set imgWJ(lngIdx).Container = picCItem(lngIdx)
    imgWJ(lngIdx).Visible = False
    
    Load imgCL(lngIdx)
    Set imgCL(lngIdx).Container = picCItem(lngIdx)
    imgCL(lngIdx).Visible = False
    
End Sub

Private Sub ResiceCard(ByVal lngIdx As Long)
'���ܣ����ÿ�Ƭ�ڲ��ؼ�λ��
    
    lblSelect(lngIdx).Left = 140
    lblSelect(lngIdx).Width = 1510
    lblSelect(lngIdx).Top = 600
    lblName(lngIdx).Top = 660
    lblName(lngIdx).Left = 160
    
    
    lblSex(lngIdx).Left = lblName(lngIdx).Left
    lblSex(lngIdx).Top = lblName(lngIdx).Top + lblName(lngIdx).Height + 120
    
    lblAge(lngIdx).Left = lblSex(lngIdx).Left + lblSex(lngIdx).Width + 300
    lblAge(lngIdx).Top = lblSex(lngIdx).Top
    
    
    lblText(lngIdx).Left = lblName(lngIdx).Left
    lblText(lngIdx).Width = lblSelect(lngIdx).Width
    lblText(lngIdx).Top = lblAge(lngIdx).Top + lblAge(lngIdx).Height + 120
    
    lblTime(lngIdx).Left = 750
    
    lblTime(lngIdx).Top = lblText(lngIdx).Top + lblText(lngIdx).Height + 300
    
    imgCL(lngIdx).Left = lblName(lngIdx).Left
    imgCL(lngIdx).Top = 300
    
    imgWJ(lngIdx).Left = imgCL(lngIdx).Left + imgCL(lngIdx).Width + 10
    imgWJ(lngIdx).Top = imgCL(lngIdx).Top
    
End Sub

Private Sub LoadAllCard()
'���ܣ���ʾ���п�Ƭ
    Dim lngX As Long, lngY As Long
    Dim i As Long
    Dim lngRowCount As Long
    
    lngX = clngX
    lngY = clngX
    
    lngRowCount = (picCardCon.Width) \ (picCItem(0).Width)
    mlngPreRowCnt = lngRowCount
    Call UnloadControls
    mrsCard.Filter = 0
    If mlngCntCard = 0 Then Exit Sub
    mrsCard.MoveFirst
    For i = 1 To mrsCard.RecordCount
        Call LoadOneCard(i, lngX, lngY)
        Call SetCardData(i, mrsCard)
        Call ResiceCard(i)
        '������һ�ſ�Ƭ������
        lngX = lngX + picCItem(i).Width
        
        If i Mod lngRowCount = 0 Then
            lngX = clngX
            lngY = lngY + picCItem(i).Height
        End If
        mrsCard.MoveNext
    Next
    
End Sub

Private Sub ReSetCardPos()
'���ܣ��������п�Ƭ��λ��
    Dim lngX As Long, lngY As Long
    Dim i As Long
    Dim lngRowCount As Long
    
    lngX = clngX
    lngY = clngX
    
    '����޿�Ƭ���˳�
    If mlngCntCard = 0 Then
        Exit Sub
    End If
    
    lngRowCount = (picCardCon.Width) \ (picCItem(0).Width)
    
    If mlngPreRowCnt = lngRowCount Then
        '���ÿһ�еĿ�Ƭ���ݲ������õ�λ��
        Exit Sub
    End If
    
    lngX = clngX
    lngY = clngX
    
    For i = 1 To mlngCntCard
        
        picCItem(i).Top = lngY
        picCItem(i).Left = lngX
    
        '������һ�ſ�Ƭ������
        lngX = lngX + picCItem(0).Width
        If i Mod lngRowCount = 0 Then
            lngX = clngX
            lngY = lngY + picCItem(0).Height
        End If
    Next
    
    '���ڹ���������ʾ
    If picCardCon.Height < picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 Then
        vscH.Visible = True
        vscH.value = 0
    Else
        vscH.Visible = False
    End If
    
    picCardCon.Height = picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100
End Sub

Private Sub timeRefreshCard_Timer()
  
    If Not mblnRefreshCard Then Exit Sub
    mblnRefreshCard = False
    timeRefreshCard.Enabled = False
    Call ShowAllCard
    timeRefreshCard.Enabled = True
End Sub

Private Sub ShowSelect()
'���ܣ�ѡ�п�Ƭ
    lblSelect(mintCurIndex).Visible = True
    
    If mint��ʾ��ʽ = 0 Then
        mlngPreCardID = Val(lblName(mintCurIndex).Tag)
        Call LoadPatiInfobyCard
    End If
End Sub

Private Sub SetCardData(ByVal lngIdx As Long, ByVal rsData As ADODB.Recordset)
'���ܣ����ؿ�Ƭ�ϵ���Ϣ
    lblText(lngIdx).Caption = rsData!Σ��ֵ���� & ""
    lblName(lngIdx).Caption = rsData!���� & ""
    lblName(lngIdx).Tag = rsData!ID & "" '---�ؼ���Ϣ
    lblSex(lngIdx).Caption = rsData!�Ա� & ""
    lblAge(lngIdx).Caption = rsData!���� & ""
    lblTime(lngIdx).Caption = Format(rsData!����ʱ��, "yyyy/MM/dd")
    If Val(rsData!״̬ & "") = 2 Then
        imgCL(lngIdx).Visible = True
        Set imgCL(lngIdx).Picture = imgCL(0).Picture
        imgCL(lngIdx).Tag = "�Ѵ���"
        
        If Val(rsData!�Ƿ�Σ��ֵ & "") = 1 Then
            imgWJ(lngIdx).Visible = True
            Set imgWJ(lngIdx).Picture = imgWJ(0).Picture
            imgWJ(lngIdx).Tag = "ȷ����Σ��ֵ"
        End If
    End If
    '--------���ϲ���Ϊ��Ƭѡ�������ṩ���ֶ� mint��ʾ��ʽ=1
    If mint��ʾ��ʽ = 0 Then
        '������Ϣ��
    End If
End Sub

Private Sub ShowCardPop()
'���ܣ������Ǽǵ�
    
    If mint��ʾ��ʽ = 1 Then
        mlng��¼ID = Val(lblName(mintCurIndex).Tag)
        
        Unload Me
    Else
        If mintCurIndex > 0 Then
            mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
            If Not mrsCard.EOF Then
                Call frmCriticalEdit.ShowMe(Me, True, 2, IIF(IsNull(mrsCard!�Һŵ�), 2, 1), _
                    Val(mrsCard!����ID & ""), Val(mrsCard!��ҳID & ""), mrsCard!�Һŵ� & "", 0, Val(mrsCard!ID & ""), Val(mrsCard!ҽ��ID & ""))
            End If
        End If
    End If
End Sub

Private Sub ShowCardBybnt(ByVal intType As Integer)
'���ܣ��鿴�Ǽǵ�
    Dim blnOK As Boolean
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            blnOK = frmCriticalEdit.ShowMe(Me, True, intType, IIF(IsNull(mrsCard!�Һŵ�), 2, 1), _
                Val(mrsCard!����ID & ""), Val(mrsCard!��ҳID & ""), mrsCard!�Һŵ� & "", 0, Val(mrsCard!ID & ""), Val(mrsCard!ҽ��ID & ""))
            If blnOK Then
                mblnOK = True
                If intType = 1 Then
                    Call LoadPatients
                    mblnRefreshCard = True
                End If
            End If
        End If
    End If
End Sub

Private Sub LoadPatiInfobyCard()
'���ܣ��л���Ƭʱ��ʾ������Ϣ

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo errH
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            If mstrPrePati = Val(mrsCard!����ID & "") & "," & Val(mrsCard!��ҳID & "") & "," & mrsCard!�Һŵ� Then
                Exit Sub
            End If
            mstrPrePati = Val(mrsCard!����ID & "") & "," & Val(mrsCard!��ҳID & "") & "," & mrsCard!�Һŵ�
            If mrsCard!�Һŵ� & "" <> "" Then
                strSql = "select a.id as ����ID, a.����� as ��ʶ��,b.���� as ���� from  ���˹Һż�¼ a,���ű� b where a.ִ�в���id=b.id and a.no=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mrsCard!�Һŵ� & "")
            ElseIf Val(mrsCard!��ҳID & "") <> 0 Then
                strSql = "select a.��ҳid as ����ID, a.סԺ�� as ��ʶ��,b.���� as ����  from ������ҳ a,���ű� b where a.��Ժ����id=b.id and a.����id=[1] and a.��ҳid=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsCard!����ID & ""), Val(mrsCard!��ҳID & ""))
            Else
                '��������
                strSql = "select 0 as ����ID, null as ��ʶ��,b.���� as ���� from ����ҽ����¼ a,���ű� b where a.���˿���ID=b.id and a.id=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsCard!ҽ��ID & ""))
            End If
            
            mlng����ID = Val(rsTmp!����ID & "")
            
            lblInfo(e����).Caption = lblName(mintCurIndex).Caption
            lblInfo(e�Ա�).Caption = lblSex(mintCurIndex).Caption
            lblInfo(e����).Caption = lblAge(mintCurIndex).Caption
            
            If Val(mrsCard!��ҳID & "") = 0 Then
                strTmp = "�����:" & rsTmp!��ʶ��
            Else
                strTmp = "סԺ��:" & rsTmp!��ʶ��
            End If
            lblInfo(e��ʶ��).Caption = strTmp
            lblInfo(e����).Caption = "����:" & rsTmp!����
            
            Call ReadPatPricture(Val(mrsCard!����ID & ""), imgLoad)
            If imgLoad.Picture = 0 Then
                imgPatient.Picture = imgDefual.Picture
            Else
                imgPatient.Picture = imgLoad.Picture
            End If
            mlng����ID = Val(mrsCard!����ID & "")
            
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LocatePati()
'���ܣ�ȱʡ��λ����
    Dim i As Long
    
    If mlngPreCardID = 0 Then
        Exit Sub
    End If
    Call ClearPatiInfo
    For i = 1 To mlngCntCard
        If Val(lblName(i).Tag) = mlngPreCardID Then
            mintCurIndex = i
            mstrPrePati = ""
            Call ShowSelect
            Exit For
        End If
    Next
End Sub

Private Sub SetFaceCtrl()
'���ܣ����ý���Ŀؼ��ɼ���
    If mint���� = 2 Then
        lblPatiDept.Visible = False
        cboPatiDept.Visible = False
        cboRegDept.Visible = False
    ElseIf mint���� = 3 Then
        picPatiC.Visible = True
        Set fraPati.Container = picPatiC
    End If
End Sub

Private Sub SetFilterInfo()
'���ܣ���ʼ����������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    If mint���� = 2 Then
        strTmp = sys.RowValue("���ű�", mlng����ID, "����")
        lblRegDept.Caption = "�Ǽǿ���:" & strTmp
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
 
    lngCur = vscH.value
    lngMin = vscH.Min
    lngMax = vscH.Max

    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 100, lngMin, lngMax) Then
            vscH.value = lngCur + (lngMax - lngMin) / 100
        Else
            vscH.value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 100, lngMin, lngMax) Then
            vscH.value = lngCur - (lngMax - lngMin) / 100
        Else
            vscH.value = lngMin
        End If
    End If
 
End Sub

Private Function Initȷ�Ͽ���() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngPreDept As Long
    
    If cboPatiDept.ListIndex <> -1 Then
        lngPreDept = cboPatiDept.ItemData(cboPatiDept.ListIndex)
    End If
    cboPatiDept.Clear
    cboPatiDept.AddItem "���п���"
    cboPatiDept.ItemData(cboPatiDept.NewIndex) = 0
    On Error GoTo errH
    Set rsTmp = GetDataToDepts
    
    For i = 1 To rsTmp.RecordCount
        cboPatiDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboPatiDept.ItemData(cboPatiDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPreDept Then '����ԭ�ж�λ
            Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
        ElseIf InStr(mstrPrivs, "ȫԺ����") > 0 Then
            If UserInfo.����ID = rsTmp!ID And (lngPreDept = 0 Or cboPatiDept.ListIndex = -1) Then 'ֱ����������
                Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
            End If
        Else
            '����ȱʡ���������Ŀ����ж��
            If rsTmp!ȱʡ = 1 And cboPatiDept.ListIndex = -1 Then
                Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboPatiDept.ListIndex = -1 And cboPatiDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboPatiDept.Hwnd, 0)
    End If
    Initȷ�Ͽ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDataToDepts() As ADODB.Recordset
'���ܣ���ȡ���Ҳ����б����ݼ�¼��
'������strIn ��������
    Dim strSql As String
    Dim strDeptIDs As String
    If optInfo(1).value Then
        '�����Ҷ�ȡ��ʾ
        '�����ż���۲��ҵĲ��˻�û���ϴ�������ֻ�Դ����в��˵Ŀ��ҵ�����
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            strSql = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where B.����ID=A.ID And B.��������='�ٴ�'" & _
                " And ((B.������� IN(2,3) " & _
                ")Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
        Else
            '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
            strSql = _
                " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                " From ���ű� A,��������˵�� B,������Ա C" & _
                " Where B.����ID=A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And B.��������='�ٴ�'"
            strSql = strSql & " Union " & _
                " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) As ȱʡ" & _
                " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                " And Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                " And Not Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
            If InStr(mstrPrivs, "ICU����") > 0 Then
                strSql = strSql & " Union " & _
                    " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                    " From ���ű� A" & _
                    " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                    " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='�ٴ�')" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            End If
            strSql = "Select ID,����,����,Max(ȱʡ) As ȱʡ From (" & strSql & ") Group By ID,����,���� Order by ����"
        End If
    End If
    
    If Not optInfo(1).value Then
        strSql = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
            " From ������Ա A,���ű� B,��������˵�� C" & _
            " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
            " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1]" & _
            " Order by B.����"
    End If
    
    On Error GoTo errH
    
    Set GetDataToDepts = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init�Ǽǿ���() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
    '��������/סԺҽ������
    str��Դ = "3"
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "1,2,3"
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        str��Դ = "1,3"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "2,3"
    End If
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(" & str��Դ & ") And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    Else
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(" & str��Դ & ") And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    End If
    
    cboRegDept.Clear
    cboRegDept.AddItem "���п���"
    cboRegDept.ItemData(cboRegDept.NewIndex) = 0
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    str����IDs = GetUser����IDs
    For i = 1 To rsTmp.RecordCount
        cboRegDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboRegDept.ItemData(cboRegDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.����ID Then
            Call Cbo.SetIndex(cboRegDept.Hwnd, cboRegDept.NewIndex) 'ֱ����������
        End If
        If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And cboRegDept.ListIndex = -1 Then
            Call Cbo.SetIndex(cboRegDept.Hwnd, cboRegDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboRegDept.ListIndex = -1 And cboRegDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboRegDept.Hwnd, 0)
    End If
        
    If cboRegDept.ListIndex <> -1 Then
        Call cboRegDept_Click  'ͬʱ��mstrDeptNode��ֵ
    End If
    Init�Ǽǿ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FunAffirm()
'���ܣ�Σ��ֵȷ��
    Dim lngΣ��ֵID As Long
    Dim lngҽ��ID As Long
    Dim lng����ID As Long
    Dim blnOK As Boolean
    
    On Error GoTo errH
    
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            lngΣ��ֵID = Val(lblName(mintCurIndex).Tag)
            lngҽ��ID = Val(mrsCard!ҽ��ID & "")
            lng����ID = Val(mrsCard!����ID & "")
            blnOK = frmCriticalEdit.ShowMe(Me, True, 3, 3, lng����ID, 0, "", 0, lngΣ��ֵID, lngҽ��ID)
            If blnOK Then
                Call LoadPatients
                Call ShowAllCard
                mblnOK = True
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
