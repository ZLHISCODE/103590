VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmCheckCourseCard 
   Caption         =   "ҩƷ�̵��¼��"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCheckCourseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdClass 
      Caption         =   "�����ࡢ��λ��ȡ(&P)"
      Height          =   350
      Left            =   3120
      TabIndex        =   28
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmbBatch 
      Caption         =   "�������ȡ(&G)"
      Height          =   350
      Left            =   1440
      TabIndex        =   27
      Top             =   5040
      Width           =   1575
   End
   Begin MSMask.MaskEdBox TxtCheckDate 
      Height          =   315
      Left            =   9510
      TabIndex        =   7
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   6600
      TabIndex        =   24
      Top             =   5085
      Width           =   1815
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8880
      TabIndex        =   20
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10200
      TabIndex        =   21
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11715
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   26
         Top             =   950
         Width           =   11235
         _cx             =   19817
         _cy             =   4948
         Appearance      =   0
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCheckCourseCard.frx":014A
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
         TabBehavior     =   1
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
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   11
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   1920
         TabIndex        =   9
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   8640
         TabIndex        =   6
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   17
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   15
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   13
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   3
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   2
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�̵��¼��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ⷿ"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   12
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   16
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   18
         Top             =   4500
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":01BF
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":03D9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":05F3
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":080D
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0A27
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0C41
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0E5B
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1075
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":128F
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":14A9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":16C3
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":18DD
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1AF7
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1D11
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1F2B
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":2145
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCheckCourseCard.frx":235F
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCourseCard.frx":2BF3
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCourseCard.frx":30F5
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����ҩƷ"
      Height          =   180
      Left            =   5760
      TabIndex        =   23
      Top             =   5160
      Width           =   720
   End
   Begin VB.Menu mnuCol 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(���������)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCheckCourseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '�Ƿ��ѡ�ⷿ
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mintDefault As Integer              'ȱʡ��λ
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Dim mstrPrivs As String                     'Ȩ��
Private mblnNoStock As Boolean              '���ز������Ƿ������̵�û�����ô洢�ⷿ��ҩƷ
Private mblnLoadData As Boolean             '���ڼ���Ƿ���װ�����ݣ������Ѵ��ڵ��ݣ�
Private mlngCurrRow As Long
Private mbln���Է������ As Boolean         'Ϊ��ʱ����ҩƷ�ķ������
Private mrsTemp As ADODB.Recordset
Private mbln��ͣ��ҩƷ As Boolean
Private mstr��λ As String                  '����������ѡ��Ļ�λ
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�̵��¼��"

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mlngFindFirst As Long
Private mlngFind As Long                            '���ڲ���
Private mrsFindName As ADODB.Recordset              '���ڲ���

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintCostDigit As Integer           '�ɱ���С��λ��
Private mintNumberDigit0 As Integer         '����С��λ��-��λ
Private mintNumberDigit1 As Integer         '����С��λ��-С��λ
Private mintMoneyDigit As Integer           '���С��λ��

Private mstr��λ As String
Private mbln��ͬ��λ As Boolean             '��С��װ��ͬ������ֻ��ʾһ����װ��λ

Private mblnNotTrigger As Boolean
Private mblnBatch As Boolean

Private Type Type_ҩƷid
    strҩƷid As String
    int�˳� As Integer
End Type

Private SQLCondition As Type_ҩƷid

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
Private Const mconIntCol��� As Integer = 7
Private Const mconIntCol���� As Integer = 8
Private Const mconIntCol�������� As Integer = 9
Private Const mconIntCol����ϵ�� As Integer = 10
Private Const mconIntcol�ӳ��� As Integer = 11
Private Const mconIntColʵ�ʲ�� As Integer = 12
Private Const mconIntColʵ�ʽ�� As Integer = 13
Private Const mconIntCol���� As Integer = 14
Private Const mconIntCol�ⷿ��λ As Integer = 15
Private Const mconIntCol��λ As Integer = 16
Private Const mconIntCol���� As Integer = 17
Private Const mconIntColЧ�� As Integer = 18
Private Const mconIntCol��׼�ĺ� As Integer = 19
Private Const mconintCol�ɱ��� As Integer = 20
Private Const mconIntCol�ۼ� As Integer = 21
Private Const mconintCol�������� As Integer = 22
Private Const mconIntCol��λ���� As Integer = 23
Private Const mconintCol��λ As Integer = 24
Private Const mconIntColС��λ���� As Integer = 25
Private Const mconintColС��λ As Integer = 26
Private Const mconintCol����_�ϼ� As Integer = 27
Private Const mconintCol��λ_�ϼ� As Integer = 28
Private Const mconintCol��־ As Integer = 29
Private Const mconintCol������ As Integer = 30
Private Const mconintCol���� As Integer = 31
Private Const mconintCol��۲� As Integer = 32
Private Const mconintCol�̵��� As Integer = 33
Private Const mconIntColҩƷ��������� As Integer = 34
Private Const mconIntColҩƷ���� As Integer = 35
Private Const mconIntColҩƷ���� As Integer = 36
Private Const mconIntColS  As Integer = 37             '������
'=========================================================================================

Private Sub GetBatchRec()
    '��ȡ������м�¼
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim lngRows As Long
    Dim i As Integer
    Dim strTemp As Variant
    Dim rsProperty As ADODB.Recordset           'ҩƷ���
    Dim rs��λ As ADODB.Recordset       '��λ
    Dim arrDrugID As Variant
    Dim j As Integer
    Dim lngҩƷID As Long
    Dim x As Integer
    Dim strҩƷid As String
    Dim strArry As Variant  '�����λ������
    Dim str��λid As String
    Dim str��λ As String
    Dim str��λsql As String
    
    On Error GoTo ErrHandle
    Set rsProperty = New ADODB.Recordset
    With rsProperty
        If .State = 1 Then .Close
        .Fields.Append "ҩƷ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷid", adDouble, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��׼�ĺ�", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rs��λ = New ADODB.Recordset
    
    With rs��λ
        If .State = 1 Then .Close
        .Fields.Append "ҩƷid", adDouble, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    x = 1
    strArry = Array()
    str��λid = ""
    For j = 0 To UBound(Split(mstr��λ, ",")) - 1
        str��λ = Mid(mstr��λ, x, InStr(x, mstr��λ, ",") - x)
        x = InStr(x, mstr��λ, ",") + 1
        If Len(IIf(str��λid = "", "", str��λid & ",") & str��λ) > 4000 Then
            ReDim Preserve strArry(UBound(strArry) + 1)
            strArry(UBound(strArry)) = str��λid
            str��λid = str��λ
        Else
            str��λid = IIf(str��λid = "", "", str��λid & ",") & str��λ
        End If
    Next
    
    If str��λid <> "" Then
'        SQLCondition.strҩƷID = ""
        ReDim Preserve strArry(UBound(strArry) + 1)
        strArry(UBound(strArry)) = str��λid
        
        gstrSQL = " Select distinct a.ҩƷid" & _
                    " From ҩƷ�����޶� A," & _
                         "�շ���ĿĿ¼ C,(select * from Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) B" & _
                    " Where a.�ⷿid = [1] and a.ҩƷid=c.id And (Instr(',' || a.�ⷿ��λ || ',', ',' || b.Column_Value || ',') > 0) "
        
        If mbln���Է������ = False Then
            gstrSQL = gstrSQL & _
                " and (Decode(c.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3)) " & _
                    " or Decode(c.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                    " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]))"
        End If
        
        For i = 0 To UBound(strArry)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ݻ�λ��ѯҩƷ", Val(txtStock.Tag), CStr(strArry(i)))
            
            If Not rsData.EOF Then
                Do While Not rsData.EOF
                    With rs��λ
                        .AddNew
                        !ҩƷid = rsData!ҩƷid
                        
                        .Update
                    End With
                    rsData.MoveNext
                Loop
            End If
        Next
    End If
    
'    If rs��λ.RecordCount > 0 Then
'        rsData.MoveFirst
'        For i = 0 To rsData.RecordCount - 1 '���ѡ���˻�λ�����ջ�λ����ȡҩƷ��Ȼ��������ȡ����ҩƷ�ڴӿ�����ȡ����
'            SQLCondition.strҩƷID = rsData!ҩƷID & "," & SQLCondition.strҩƷID
'            rsData.MoveNext
'        Next
'    End If
    
'    If SQLCondition.strҩƷID = "" Then
'        MsgBox "δ��ѯ�����ݣ�", vbInformation, gstrSysName
'        Exit Sub
'    Else
        If SQLCondition.strҩƷid <> "" And str��λid <> "" Then
            strTemp = Split(SQLCondition.strҩƷid, ",")
            SQLCondition.strҩƷid = ""
            
            For i = 0 To UBound(strTemp) - 1
                rs��λ.MoveFirst
                For j = 0 To rs��λ.RecordCount - 1
                    If rs��λ.EOF Then Exit For
                    If Val(strTemp(i)) = Val(rs��λ!ҩƷid) Then
                        SQLCondition.strҩƷid = strTemp(i) & "," & SQLCondition.strҩƷid
                    End If
                    If j <> rs��λ.RecordCount - 1 Then
                        rs��λ.MoveNext
                    End If
                Next
            Next
        ElseIf SQLCondition.strҩƷid = "" And str��λid <> "" Then
            If rs��λ.RecordCount > 0 Then
                rs��λ.MoveFirst
            End If
            
            Do While Not rs��λ.EOF
                SQLCondition.strҩƷid = rs��λ!ҩƷid & "," & SQLCondition.strҩƷid
                rs��λ.MoveNext
            Loop
        ElseIf SQLCondition.strҩƷid = "" And str��λid = "" Then
            Exit Sub
        End If
        
        x = 1
        arrDrugID = Array()
        strҩƷid = ""
        For j = 0 To UBound(Split(SQLCondition.strҩƷid, ",")) - 1
            lngҩƷID = Mid(SQLCondition.strҩƷid, x, InStr(x, SQLCondition.strҩƷid, ",") - x)
            x = InStr(x, SQLCondition.strҩƷid, ",") + 1
            If Len(IIf(strҩƷid = "", "", strҩƷid & ",") & lngҩƷID) > 4000 Then
                ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
                arrDrugID(UBound(arrDrugID)) = strҩƷid
                strҩƷid = lngҩƷID
            Else
                strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & lngҩƷID
            End If
        Next
        
        If strҩƷid = "" And UBound(arrDrugID) < 0 Then
            Exit Sub
        ElseIf strҩƷid <> "" Then
            ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
            arrDrugID(UBound(arrDrugID)) = strҩƷid
        End If
        
        gstrSQL = "Select b.���� As ҩƷ����, a.ҩƷid, Nvl(a.����, 0) As ����, a.��׼�ĺ�" & _
                   " From ҩƷ��� A, �շ���ĿĿ¼ B" & _
                   " Where A.���� = 1 And A.ҩƷid = b.Id And A.�ⷿid = [1] And " & _
                   " b.Id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
                   " And (Nvl(A.ʵ������,0)<>0 Or Nvl(A.ʵ�ʽ��,0)<>0 Or Nvl(A.ʵ�ʲ��,0)<>0 )"
        
        If mbln���Է������ = False Then
            gstrSQL = gstrSQL & _
                " and (Decode(b.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3)) " & _
                    " or Decode(b.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                    " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]))"
        End If
        
        gstrSQL = gstrSQL & " Order By b.����"
        
        For i = 0 To UBound(arrDrugID)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetBatchRec", Val(txtStock.Tag), CStr(arrDrugID(i)))
            
            If Not rsData.EOF Then
                Do While Not rsData.EOF
                    With rsProperty
                        .AddNew
                        !ҩƷ���� = rsData!ҩƷ����
                        !ҩƷid = rsData!ҩƷid
                        !���� = rsData!����
                        !��׼�ĺ� = rsData!��׼�ĺ�
                        
                        .Update
                    End With
                    rsData.MoveNext
                Loop
            End If
        Next
'    End If
    
    If rsProperty.RecordCount = 0 Then
        Exit Sub
    End If
    rsProperty.MoveFirst
    With rsProperty
        If .RecordCount = 0 Then Exit Sub
        
        mblnBatch = True
        
        lngRows = .RecordCount
        
        vsfBill.rows = lngRows + 1
        
        For lngRow = 1 To lngRows
            vsfBill.Row = lngRow
            Call SetPhiscRows(!ҩƷid, !����, Nvl(!��׼�ĺ�, ""), True)
            
            DoEvents
            Call zlControl.StaShowPercent(lngRow / lngRows, staThis.Panels(2), frmCheckCourseCard)
            DoEvents
            
            .MoveNext
        Next
    End With
    
    staThis.Panels(2).Text = ""
    
    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    
    mblnBatch = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With vsfBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntColҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Sub SetSortRecord()
    Dim n As Integer
    
    If vsfBill.rows < 2 Then Exit Sub
    If vsfBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfBill.rows - 1
            If vsfBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(vsfBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(vsfBill.TextMatrix(n, mconIntCol���)))
                !ҩƷid = Val(vsfBill.TextMatrix(n, 0))
                !���� = Val(vsfBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub
'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 14  and b.ϵ��=1 "
    Set rsDepend = zlDatabase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�̵��¼��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mblnFirst = True
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1307)

    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        CmdSave.Visible = False
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cmbBatch_Click()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim lngRows As Long
    
    If MsgBox("��ȡ��ǰ����¼�������������ݽ�������Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    vsfBill.rows = 1
    
    gstrSQL = "Select B.���� As ҩƷ����, A.ҩƷid, Nvl(A.����, 0) As ����, A.��׼�ĺ� " & _
            " From ҩƷ��� A, �շ���ĿĿ¼ B " & _
            " Where A.���� = 1 And A.ҩƷid = B.Id And A.�ⷿid = [1] " & _
                " And (Nvl(A.ʵ������,0)<>0 Or Nvl(A.ʵ�ʽ��,0)<>0 Or Nvl(A.ʵ�ʲ��,0)<>0 )"
            
    If mbln���Է������ = False Then
        gstrSQL = gstrSQL & _
            " and (Decode(B.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3)) " & _
                " or Decode(B.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]))"
    End If
            
    gstrSQL = gstrSQL & " Order By B.���� "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetBatchRec", Val(txtStock.Tag))
    
    If rsData.RecordCount = 0 Then
        Exit Sub
    End If
    rsData.MoveFirst
    With rsData
        If .RecordCount = 0 Then Exit Sub
        
        mblnBatch = True
        
        lngRows = .RecordCount
        
        vsfBill.rows = lngRows + 1
        
        For lngRow = 1 To lngRows
            vsfBill.Row = lngRow
            Call SetPhiscRows(!ҩƷid, !����, Nvl(!��׼�ĺ�, ""), True)
            
            DoEvents

            Call zlControl.StaShowPercent(lngRow / lngRows, staThis.Panels(2), frmCheckCourseCard)
            DoEvents
            
            .MoveNext
        Next
    End With
    
    staThis.Panels(2).Text = ""
    
    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    
    mblnBatch = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Public Sub getҩƷid(ByRef strҩƷid As String, ByRef intExit As Integer)
'    SQLCondition.strҩƷid = strҩƷid
'    SQLCondition.int�˳� = intExit
'End Sub

Private Sub cmdClass_Click()
    Dim lngValue As Long
    Dim intCol As Integer
    
'    lngValue = MsgBox("��ȡ�����¼�������������ݽ�������Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
'    If lngValue = vbYes Then
        frmCheckClass.ShowME Me, txtStock.Tag, mstr��λ, SQLCondition.strҩƷid, SQLCondition.int�˳�
        If SQLCondition.int�˳� = 1 Then    '1-ѡ����������0-û��ѡ������ �˳���ִ��ˢ�²���
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
            Call GetBatchRec
        End If
'    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If Trim(txtCode.Text) = "" Then
            txtCode.SetFocus
        Else
            Call FindGridRow(txtCode.Text)
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long, lngRows As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim str�������� As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strҩ�� As String
    
    '����ҩƷ
    On Error GoTo ErrHandle
    If strInput = txtCode.Tag Then
        '��ʾ������һ����¼
        If mlngFind >= vsfBill.rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '��ʾ�µĲ���
        mlngFindFirst = 0
        lngStart = 0
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    mlngFind = 0
    lngStart = lngStart + 1
    lngRows = vsfBill.rows - 1
    
    mrsFindName.MoveFirst
    For n = 1 To mrsFindName.RecordCount
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If
    
        lngFindRow = vsfBill.FindRow(strҩ��, lngStart, mconIntColҩƷ���������, True, True)
        If lngFindRow > 0 Then
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
            mlngFind = lngFindRow
            
            '��¼�ҵ��ĵ�1����¼
            If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
            
            Exit For
        End If
        mrsFindName.MoveNext
        
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF And lngFindRow = -1 And mlngFindFirst <> 0 Then
            vsfBill.Select mlngFindFirst, 1, mlngFindFirst, vsfBill.Cols - 1
            vsfBill.TopRow = mlngFindFirst
            mlngFind = mlngFindFirst
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    
    '�����������ݼ�
    Call SetSortRecord
    
    Me.txtNo.Tag = ""
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
        If Val(zlDatabase.GetPara("���̴�ӡ", glngSys, ģ���.ҩƷ�̵�)) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    vsfBill.rows = 2
    vsfBill.Cell(flexcpText, 1, 0, 1, vsfBill.Cols - 1) = ""
    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
            
    mblnFirst = False
    mbln��ͣ��ҩƷ = IIf(Val(zlDatabase.GetPara("����ͣ�õ�ҩƷ", glngSys, glngModul, 0)) = 0, False, True)
    If mint�༭״̬ = 1 Then
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    Else
        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    vsfBill.SetFocus
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) = "" Then
        vsfBill.Col = mconIntColҩ��
    Else
        vsfBill.Col = mconIntCol��λ����
        vsfBill.EditCell
    End If
End Sub

Private Sub Form_Load()
    mintMoneyDigit = GetDigit(0, 1, 4)
    mblnNoStock = (Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, ģ���.ҩƷ�̵�)) = 1)
    mbln���Է������ = (Val(zlDatabase.GetPara("����ҩƷ�������", glngSys, ģ���.ҩƷ�̵�)) = 1)
    mintBatchNoLen = GetBatchNoLen()
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnLoadData = False
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�̵����", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call initCard
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    ElseIf strCompare = "3" Then
        strSqlOrder = "�ⷿ��λ"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ����,���"
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
            TxtCheckDate.Text = Txt��������.Caption
            txtStock = mfrmMain.cboStock.Text
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Call ��ȡ��λ
            initGrid
        Case 2, 3, 4
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Call ��ȡ��λ
            initGrid
            If mint�༭״̬ <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "select distinct b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id " _
                        & "  and A.���� = 14 and a.no=[1]"
                Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!����
                txtStock.Tag = rsInitCard!id
                
                rsInitCard.Close
            End If
            
            strUnitQuantity = "A.���� ʵ������,A.��д���� ��������,A.ʵ������ ������,B.סԺ��λ AS סԺ��λ,B.סԺ��װ as סԺϵ��,a.���ۼ�*B.סԺ��װ as סԺ�ۼ�,"
            strUnitQuantity = strUnitQuantity & "B.���ﵥλ AS ���ﵥλ,B.�����װ as ����ϵ��,a.���ۼ�*B.�����װ as �����ۼ�,"
            strUnitQuantity = strUnitQuantity & "B.ҩ�ⵥλ AS ҩ�ⵥλ,B.ҩ���װ as ҩ��ϵ��,a.���ۼ�*B.ҩ���װ as ҩ���ۼ�,"
            strUnitQuantity = strUnitQuantity & "D.���㵥λ AS �ۼ۵�λ,'1' as �ۼ�ϵ��,a.���ۼ� as �ۼ��ۼ�,"

            gstrSQL = "SELECT * " & _
                " FROM " & _
                " (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                " NVL(B.���Ч��,0) ���Ч��,B.ҩƷ��Դ,B.����ҩ��,D.���,A.����,Nvl(A.�ⷿ��λ,C.�ⷿ��λ) As �ⷿ��λ, A.����,A.Ч��,A.����," & strUnitQuantity & _
                " A.���۽�� AS ����,A.��� AS ��۲�,A.���ۼ�,A.���� As �ɱ���, " & _
                " A.ժҪ,������,��������,�����,�������,A.Ƶ�� AS �̵�ʱ��,A.�ɱ��� AS �����,A.�ɱ���� AS �����,B.�ӳ���,D.�Ƿ���,B.ҩ������ AS ҩ����������,A.��׼�ĺ� " & _
                " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ D,ҩƷ�����޶� C " & _
                " WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=D.ID " & _
                " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                " AND A.ҩƷID=C.ҩƷID(+) AND A.�ⷿID=C.�ⷿID(+) AND A.��¼״̬ =[2] " & _
                " AND A.���� =14 AND A.NO = [1]) " & _
                " ORDER BY " & strSqlOrder
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û�����
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd HH:mm:ss")
            
            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd HH:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            TxtCheckDate.Text = rsInitCard!�̵�ʱ��
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            intRow = 0
            With vsfBill
                Do While Not rsInitCard.EOF
                    
                    intRow = intRow + 1
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsInitCard!ͨ����
                    Else
                        strҩ�� = IIf(IsNull(rsInitCard!��Ʒ��), rsInitCard!ͨ����, rsInitCard!��Ʒ��)
                    End If
                    
                    .TextMatrix(intRow, mconIntColҩƷ���������) = rsInitCard!ҩƷ���� & strҩ��
                    .TextMatrix(intRow, mconIntColҩƷ����) = rsInitCard!ҩƷ����
                    .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsInitCard!��Ʒ��), "", rsInitCard!��Ʒ��)
                    
                    .TextMatrix(intRow, mconIntCol��Դ) = Nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = Nvl(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsInitCard!�ⷿ��λ), "", rsInitCard!�ⷿ��λ)
                    .TextMatrix(intRow, mconIntCol��λ) = IIf(IsNull(rsInitCard.Fields(Split(mstr��λ, "|")(1)).Value), "", rsInitCard.Fields(Split(mstr��λ, "|")(1)).Value)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx(rsInitCard!�ӳ��� / 100, 2, , True) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = ��ȡ����ϵ��(rsInitCard)

                    If mbln��ͬ��λ = True Then
                        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Nvl(rsInitCard!�ɱ���, 0) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0)), mintPriceDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Nvl(rsInitCard!���ۼ�, 0) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0)), mintPriceDigit, , True)
                    
                        .TextMatrix(intRow, mconIntCol��λ����) = zlStr.FormatEx(rsInitCard.Fields("ʵ������").Value / Split(��ȡ����ϵ��(rsInitCard), "|")(0), mintNumberDigit0, , True)
                        .TextMatrix(intRow, mconintCol��λ) = IIf(IsNull(rsInitCard.Fields(Split(mstr��λ, "|")(0)).Value), "", rsInitCard.Fields(Split(mstr��λ, "|")(0)).Value)
                    Else
                        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Nvl(rsInitCard!�ɱ���, 0) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1)), mintPriceDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Nvl(rsInitCard!���ۼ�, 0) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1)), mintPriceDigit, , True)
                    
                        .TextMatrix(intRow, mconIntCol��λ����) = zlStr.FormatEx(Int(rsInitCard.Fields("ʵ������").Value / Split(��ȡ����ϵ��(rsInitCard), "|")(0)), mintNumberDigit0, , True)
                        .TextMatrix(intRow, mconintCol��λ) = IIf(IsNull(rsInitCard.Fields(Split(mstr��λ, "|")(0)).Value), "", rsInitCard.Fields(Split(mstr��λ, "|")(0)).Value)
                        
                        .TextMatrix(intRow, mconIntColС��λ����) = zlStr.FormatEx((rsInitCard.Fields("ʵ������").Value / Split(��ȡ����ϵ��(rsInitCard), "|")(0) - Val(.TextMatrix(intRow, mconIntCol��λ����))) * Split(��ȡ����ϵ��(rsInitCard), "|")(0) / Val(Split(��ȡ����ϵ��(rsInitCard), "|")(1)), mintNumberDigit1, , True)
                        .TextMatrix(intRow, mconintColС��λ) = IIf(IsNull(rsInitCard.Fields(Split(mstr��λ, "|")(1)).Value), "", rsInitCard.Fields(Split(mstr��λ, "|")(1)).Value)
                        
                        .TextMatrix(intRow, mconintCol����_�ϼ�) = zlStr.FormatEx(rsInitCard.Fields("ʵ������").Value, mintNumberDigit1, , True)
                        .TextMatrix(intRow, mconintCol��λ_�ϼ�) = IIf(IsNull(rsInitCard.Fields("�ۼ۵�λ")), "", rsInitCard.Fields("�ۼ۵�λ"))
                    End If
                    
                    .RowData(intRow) = Val(IIf(IsNull(rsInitCard!���Ч��), 0, rsInitCard!���Ч��))
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    mblnLoadData = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



'��ʼ���༭�ؼ�
Private Sub initGrid()
    Dim i As Integer
    
    With vsfBill
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntColS
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .RowHeightMax = 315
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconintCol��������) = "��������"
        .TextMatrix(0, mconIntCol��λ����) = IIf(mbln��ͬ��λ, "����", "���װ")
        .TextMatrix(0, mconintCol��λ) = "��λ"
        .TextMatrix(0, mconIntColС��λ����) = "С��װ"
        .TextMatrix(0, mconintColС��λ) = "��λ"
        .TextMatrix(0, mconintCol����_�ϼ�) = "�ϼ�"
        .TextMatrix(0, mconintCol��λ_�ϼ�) = "��λ"
        .TextMatrix(0, mconintCol��־) = "��־"
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconintCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconintCol����) = "����"
        .TextMatrix(0, mconintCol��۲�) = "��۲�"
        .TextMatrix(0, mconintCol�̵���) = "�̵���"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol�ⷿ��λ) = 2000
        .ColWidth(mconIntCol��λ) = 0
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconintCol��������) = 0
        .ColWidth(mconIntCol��λ����) = 1000
        .ColWidth(mconintCol��λ) = 500
        .ColWidth(mconIntColС��λ����) = IIf(mbln��ͬ��λ, 0, 1000)
        .ColWidth(mconintColС��λ) = IIf(mbln��ͬ��λ, 0, 500)
        .ColWidth(mconintCol����_�ϼ�) = IIf(mbln��ͬ��λ, 0, 1000)
        .ColWidth(mconintCol��λ_�ϼ�) = IIf(mbln��ͬ��λ, 0, 500)
        .ColWidth(mconintCol��־) = 0
        .ColWidth(mconintCol������) = 0
        .ColWidth(mconintCol�ɱ���) = IIf(mblnViewCost, 1000, 0)
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconintCol����) = 0
        .ColWidth(mconintCol��۲�) = 0
        .ColWidth(mconintCol�̵���) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
                
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
        End If
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconintCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconintColС��λ) = flexAlignCenterCenter
        .ColAlignment(mconintCol��λ_�ϼ�) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconintCol��������) = flexAlignRightCenter
        .ColAlignment(mconintCol��־) = flexAlignCenterCenter
        .ColAlignment(mconintCol������) = flexAlignRightCenter
        .ColAlignment(mconintCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconintCol����) = flexAlignRightCenter
        .ColAlignment(mconintCol��۲�) = flexAlignRightCenter
        .ColAlignment(mconintCol�̵���) = flexAlignRightCenter
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .Cell(flexcpFontBold, 1, mconIntCol��λ����, 1, mconIntCol��λ����) = True
        .Cell(flexcpFontBold, 1, mconIntColС��λ����, 1, mconIntColС��λ����) = True
        
        .Redraw = flexRDDirect
    End With
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, MStrCaption
    '�ָ����Ի��������ú�Ȩ�޿��Ƶ�����Ҫ��һ��������ʾ
    vsfBill.ColWidth(mconintCol�ɱ���) = IIf(mblnViewCost, 1000, 0)
    
    vsfBill.ColWidth(mconIntColС��λ����) = IIf(mbln��ͬ��λ, 0, 1000)
    vsfBill.ColWidth(mconintColС��λ) = IIf(mbln��ͬ��λ, 0, 500)
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        vsfBill.ColWidth(mconIntCol��Ʒ��) = IIf(vsfBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, vsfBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        vsfBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    
    With vsfBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = vsfBill.Left + vsfBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    TxtCheckDate.Left = vsfBill.Left + vsfBill.Width - TxtCheckDate.Width
    lblCheckDate.Left = TxtCheckDate.Left - lblCheckDate.Width - 100
    
    LblStock.Left = vsfBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = vsfBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = vsfBill.Left + vsfBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = vsfBill.Left + vsfBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = vsfBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = Pic����.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
        
    End With
    
    With vsfBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + vsfBill.Left + vsfBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + vsfBill.Left
        .Top = CmdCancel.Top
    End With
    
    With cmbBatch
        .Top = cmdHelp.Top
    End With
    With cmdClass
        .Top = cmbBatch.Top
        .Left = cmbBatch.Left + cmbBatch.Width + 100
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�̵����", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
End Sub

Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
    End If
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntColҩ��
                .ColComboList(mconIntColҩ��) = "..."
            Case mconIntCol����
                .ColComboList(mconIntCol����) = "..."
        End Select
    End With
End Sub

Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    Dim rsProvider As ADODB.Recordset
    
    intOldRow = vsfBill.Row
    With vsfBill
        Select Case Col
            Case mconIntColҩ��
'                If mblnNotTrigger = True Then
'                    mblnNotTrigger = False
'                    Exit Sub
'                End If
                
                If mblnNotTrigger <> True Then
                    mblnNotTrigger = True
'                    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , False, True, False, True, zlStr.IsHavePrivs(mstrPrivs, "�鿴�̵㵥���"), 0, mblnNoStock, 0, False, mbln���Է������)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , txtStock.Tag, txtStock.Tag, , 0, False, True, zlStr.IsHavePrivs(mstrPrivs, "�鿴�̵㵥���"), IIf(mbln��ͣ��ҩƷ, 1, 0), , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                    mblnNotTrigger = False
                Else
                    Exit Sub
                End If
                
                If RecReturn.RecordCount > 0 Then
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        intCurRow = .Row
                        Call SetPhiscRows(RecReturn!ҩƷid, IIf(IsNull(RecReturn!����), 0, RecReturn!����), IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�))
'                        .EditCell
                        
                        If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                            .rows = .rows + 1
                        End If
                        .Row = .rows - 1
                        RecReturn.MoveNext
                    Next
                    .Row = intOldRow
                    If Val(.TextMatrix(Row, mconIntCol����)) = -1 And .TextMatrix(Row, mconIntCol����) = "" Then
                        .Col = mconIntCol����
                    Else
                        .Col = mconIntCol��λ����
                    End If
                End If
            Case mconIntCol����
                vRect = zlControl.GetControlRect(vsfBill.hWnd)
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top + vsfBill.CellTop
                
                gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
                Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                True, dblLeft, dblTop, 300, blnCancel, False, True, gstrNodeNo)
                
                If rsProvider Is Nothing Then
                    Exit Sub
                End If
                If Not rsProvider.EOF Then
                    .TextMatrix(.Row, mconIntCol����) = rsProvider!����
                End If
        End Select
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim rs����ʱ�� As ADODB.Recordset
    Dim str��� As String
    Dim strSQL As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim strNotNum As String  '�޿��ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strNotҩ�� As String    '������¼��ЩҩƷ��ʱ�۵��޿��
    Dim str�̵�ʱ���ҩƷ As String       '��¼���̵�ʱ�������ҩƷ
    Dim strSql�̵� As String   '�����̵�ʱ�������ҩƷ
    
    rsTemp.MoveFirst
    str�̵�ʱ���ҩƷ = ""
    strSql�̵� = ""
    str���� = ""
    strTemp = ""
    Do While Not rsTemp.EOF
    
        str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
        
        gstrSQL = "Select a.����ʱ�� From �շ���ĿĿ¼ A Where a.Id =[1]"
        Set rs����ʱ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����ʱ��", rsTemp!ҩƷid)
        If Format(rs����ʱ��!����ʱ��, "yyyy-MM-dd HH:mm:ss") > Format(TxtCheckDate.Text, "yyyy-MM-dd HH:mm:ss") Then
            str�̵�ʱ���ҩƷ = str�̵�ʱ���ҩƷ & ";" & "[" & rsTemp!ҩƷ���� & "]" & rsTemp!ͨ����
            strSql�̵� = strSql�̵� & "ҩƷid<>" & rsTemp!ҩƷid & " and "
        End If
        
        rsTemp.MoveNext
    Loop
        
    If strSql�̵� <> "" Then
        MsgBox Mid(str�̵�ʱ���ҩƷ, 2) & vbCrLf & "����ҩƷΪ�̵�ʱ����������Բ��ᱻ��ӣ�", vbInformation, gstrSysName
        rsTemp.Filter = Mid(strSql�̵�, 1, Len(strSql�̵�) - 4)
    End If
    
    With vsfBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            '35.60�汾֧��ͬ������ҩƷ¼�������Σ����������=-1(��������)������
            If Val(.TextMatrix(i, mconIntCol����)) >= 0 Then
                If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                    strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
                End If
            End If
        Next
        
        If strInfo <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        '�ж���ʲô��ʽƴ��sql
        If str�ظ�ҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
End Function

Private Sub vsfBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
    If vsfBill.rows <= 2 Then
        TxtCheckDate.Enabled = True
    Else
        TxtCheckDate.Enabled = False
    End If
End Sub

Private Sub vsfBill_EnterCell()
    Dim lng���� As Long
    
    If mblnBatch = True Then Exit Sub
    
    With vsfBill
        .Editable = flexEDNone
        If mint�༭״̬ = 4 Then Exit Sub
        
        lng���� = Val(.TextMatrix(.Row, mconIntCol����))
        
        Select Case .Col
            Case mconIntColҩ��
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntColҩ��) = "..."
                End If
                
            Case mconIntCol����
                .EditMaxLength = mintBatchNoLen
                
                If lng���� = -1 Then
                    .Editable = flexEDKbdMouse
                End If
            Case mconIntCol����
                If lng���� = -1 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2) Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol����) = "..."
                End If
            Case mconIntColЧ��
'                .TextMask = "1234567890-"
                .EditMaxLength = 10
                
                If lng���� = -1 Then
                    .Editable = flexEDKbdMouse
                End If
                
                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                        If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                            strxq = TranNumToDate(strxq)
                            If strxq = "" Then Exit Sub
                            
                            .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                                '����Ϊ��Ч��
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mconIntCol��λ����, mconIntColС��λ����
                .EditMaxLength = 16
'                .TextMask = ".1234567890"
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                    .Editable = flexEDKbdMouse
                End If
            Case mconintCol�ɱ���
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                    If Val(.TextMatrix(.Row, mconIntCol����)) = -1 Then
                       .Editable = flexEDKbdMouse
                    End If
                End If
        End Select
        
        If mlngCurrRow <> .Row Then
            mlngCurrRow = .Row
            Call ��ʾ�ϼƽ��
            Call ��ʾ�����
        End If
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBill
        If KeyCode = vbKeyDelete Then
            If .rows = 2 Then Exit Sub
            If .TextMatrix(.Row, mconIntCol�к�) = "" Then Exit Sub
            If InStr(1, "34", mint�༭״̬) <> 0 Then Exit Sub
            
            If MsgBox("�Ƿ�ɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshRowNO(vsfBill, mconIntCol�к�, .Row)
            End If
        End If
        
        If txtCode.Visible And KeyCode = vbKeyF3 Then
            Call txtCode_KeyPress(13)
        End If
        
        Select Case .Col
            Case mconIntColҩ��
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntColҩ��) = ""
                ElseIf .EditText = "" Then
'                    mblnNotTrigger = True
                    If .TextMatrix(.Row, mconIntColҩ��) = "" Then
                        txtժҪ.SetFocus
                    Else
                        If Val(.TextMatrix(.Row, mconIntCol����)) = -1 And .TextMatrix(.Row, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol��λ����
                        End If
                        .EditCell
                    End If
                End If
        End Select
    End With
End Sub
Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    Dim rsProvider As ADODB.Recordset
    
    intOldRow = vsfBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfBill
        .EditText = Trim(.EditText)
        strKey = Trim(.EditText)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mconIntColҩ��
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + vsfBill.Left + vsfBill.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + vsfBill.Top + vsfBill.CellTop + vsfBill.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - vsfBill.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , strkey, sngLeft, sngTop, False, True, False, True, zlStr.IsHavePrivs(mstrPrivs, "�鿴�̵㵥���"), 0, mblnNoStock, 0, False, mbln���Է������)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strKey, sngLeft, sngTop, txtStock.Tag, txtStock.Tag, , 0, False, True, zlStr.IsHavePrivs(mstrPrivs, "�鿴�̵㵥���"), IIf(mbln��ͣ��ҩƷ, 1, 0), , mstrPrivs)
                    
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            Call SetPhiscRows(RecReturn!ҩƷid, IIf(IsNull(RecReturn!����), 0, RecReturn!����), IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�))
                            
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                        If Val(.TextMatrix(Row, mconIntCol����)) = -1 And .TextMatrix(Row, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol��λ����
                        End If
                    End If
                    Call ��ʾ�����
                End If
            Case mconIntCol����
                '�޴���
                .TextMatrix(.Row, mconIntCol����) = strKey
                
                If .TextMatrix(.Row, mconIntColЧ��) = "" Then
                    .Col = mconIntColЧ��
                Else
                    .Col = mconIntCol��λ����
                End If
                .EditCell
            Case mconIntCol����
                vRect = zlControl.GetControlRect(vsfBill.hWnd)
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top + vsfBill.CellTop
                
                gstrSQL = "Select ���� as id,����,���� From ҩƷ������ " _
                            & "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) Order By ����"
                
                Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", strKey & "%", gstrNodeNo)
                
                If rsProvider Is Nothing Then
                    .EditText = ""
                    .TextMatrix(.Row, .Col) = ""
                    Exit Sub
                End If
                If Not rsProvider.EOF Then
                    .TextMatrix(.Row, mconIntCol����) = rsProvider!����
                    .EditText = rsProvider!����
                End If
            Case mconIntColЧ��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            .EditText = ""
                            MsgBox "�Բ���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Exit Sub
                        End If
                        .EditText = strKey
                    ElseIf Not IsDate(strKey) Then
                        .EditText = ""
                        MsgBox "�Բ���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    End If
                End If
                
                .TextMatrix(.Row, mconIntColЧ��) = strKey
                .Col = mconIntCol��λ����
                .EditCell
            Case mconIntCol��λ����, mconIntColС��λ����
                If strKey <> "" Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ���ʵ����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    End If
                Else
                    .EditText = IIf(.TextMatrix(.Row, .Col) = "", " ", .TextMatrix(.Row, .Col))
                    .TextMatrix(.Row, .Col) = .EditText
                End If
                
                If strKey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    If .Col = mconIntCol��λ���� Then
                        strKey = zlStr.FormatEx(strKey, mintNumberDigit0, , True)
                    Else
                        strKey = zlStr.FormatEx(strKey, mintNumberDigit1, , True)
                    End If
                    .EditText = strKey
                End If
                
                '��ʾ�ϼ�����
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If .Col = mconIntCol��λ���� Then
                    strKey = Val(.TextMatrix(.Row, mconIntColС��λ����)) + Val(strKey) * Val(Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(0)) / Val(Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(1))
                Else
                    strKey = Val(strKey) + Val(.TextMatrix(.Row, mconIntCol��λ����)) * Val(Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(0)) / Val(Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(1))
                End If
                .TextMatrix(.Row, mconintCol����_�ϼ�) = zlStr.FormatEx(strKey * Val(Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(1)), mintNumberDigit1, , True)
                
                Call ��ʾ�ϼƽ��
                
                If Col = mconIntCol��λ���� Then
                    If .ColWidth(mconIntColС��λ����) > 0 Then
                        .Col = mconIntColС��λ����
                    Else
                        '�����һ��Ϊ�ջ���ҩ����Ϊ���򷵻ص�ҩ���У����򷵻ص�ʵ��������
                        If .Row < .rows - 1 Then
                            .Row = .Row + 1
                            If .TextMatrix(.Row, mconIntColҩ��) <> "" Then
                                .Col = mconIntCol��λ����
                            Else
                                .Col = mconIntColҩ��
                            End If
                        Else
                            .rows = .rows + 1
                            .Row = .rows - 1
                            .TextMatrix(.Row, mconIntCol�к�) = .Row
                            .Col = mconIntColҩ��
                            
                            .Cell(flexcpFontBold, .rows - 1, mconIntCol��λ����, .rows - 1, mconIntCol��λ����) = True
                            .Cell(flexcpFontBold, .rows - 1, mconIntColС��λ����, .rows - 1, mconIntColС��λ����) = True
                        End If
                    End If
                Else
                    '�����һ��Ϊ�ջ���ҩ����Ϊ���򷵻ص�ҩ���У����򷵻ص�ʵ��������
                    If .Row < .rows - 1 Then
                        .Row = .Row + 1
                        If .TextMatrix(.Row, mconIntColҩ��) <> "" Then
                            .Col = mconIntCol��λ����
                        Else
                            .Col = mconIntColҩ��
                        End If
                    Else
                        .rows = .rows + 1
                        .Row = .rows - 1
                        .TextMatrix(.Row, mconIntCol�к�) = .Row
                        .Col = mconIntColҩ��
                        
                        .Cell(flexcpFontBold, .rows - 1, mconIntCol��λ����, .rows - 1, mconIntCol��λ����) = True
                        .Cell(flexcpFontBold, .rows - 1, mconIntColС��λ����, .rows - 1, mconIntColС��λ����) = True
                    End If
                End If
            End Select
    End With
End Sub


Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBill
        Select Case Col
            Case mconIntCol��λ����, mconIntColС��λ����
                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If InStr(.EditText, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    
                    Select Case .Col
                        Case mconIntCol��λ����
                            intDigit = mintNumberDigit0
                        Case mconIntColС��λ����
                            intDigit = mintNumberDigit1
                    End Select
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Case mconIntColЧ��
                If InStr("1234567890-" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case mconintCol�ɱ���
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= mintCostDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
        End Select
    End With
End Sub


Private Sub vsfBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With vsfBill
            If .Col = mconIntColҩ�� Then
                If .Row < 1 Then Exit Sub
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub TxtCheckDate_GotFocus()
    With TxtCheckDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheckDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub TxtCheckDate_LostFocus()
    If Not IsDate(TxtCheckDate.Text) Then
        MsgBox "��������ȷ�����ڸ�ʽ��"
        TxtCheckDate.SetFocus
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim lngҩƷID As Long
    Dim str���� As String, str���� As String, dbl�ɱ��� As Double
    Dim intRow As Integer
    
    With vsfBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    '����ҩƷ����¼����غ�����
                    If Val(.TextMatrix(intLop, mconIntCol����)) = -1 And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntCol����) = "") Then
                        MsgBox "��" & intLop & "�е�ҩƷ���������η���ҩƷ,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol����
                        End If
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "��ҩƷ�����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntCol����
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol��λ����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�Ĵ��װ�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntCol��λ����
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntColС��λ����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ��С��װ�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntColС��λ����
                        .EditCell
                        Exit Function
                    End If
                End If
            Next
            
            '������ҩƷ�������εĲ��أ������Ƿ��ظ�
            For intLop = 1 To .rows - 1
                If Val(.TextMatrix(intLop, mconIntCol����)) = -1 Then
                    lngҩƷID = Val(.TextMatrix(intLop, 0))
                    str���� = .TextMatrix(intLop, mconIntCol����)
                    str���� = .TextMatrix(intLop, mconIntCol����)
                    dbl�ɱ��� = Val(.TextMatrix(intLop, mconintCol�ɱ���))
                    
                    For intRow = 1 To .rows - 1
                        If intLop <> intRow And _
                            lngҩƷID = Val(.TextMatrix(intRow, 0)) And _
                            str���� = .TextMatrix(intRow, mconIntCol����) And _
                            str���� = .TextMatrix(intRow, mconIntCol����) And _
                            dbl�ɱ��� = Val(.TextMatrix(intRow, mconintCol�ɱ���)) Then
                                                        
                            MsgBox "��" & intLop & "�е�ҩƷ(" & Trim(.TextMatrix(intLop, mconIntColҩ��)) & ")�������εĲ��أ����ţ��ɱ��ۺ͵�" & intRow & "���ظ��ˣ�" & vbCrLf & "������¼����غ�������Ϣ��", vbInformation, gstrSysName
                            
                            vsfBill.SetFocus
                            .Row = intLop
                            .TopRow = intLop
                            .Col = mconIntCol����
                            .EditCell
                            Exit Function
                        End If
                    Next
                End If
                
            Next
            
            '���ۼ��
            For intLop = 1 To .rows - 1
                If vsfBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                    If Val(.TextMatrix(intLop, mconIntCol����)) = -1 Then
                    '��������ʱ
                         If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                            '��������۹����������ۼۺͳɱ��۹�ϵ
                            If Val(vsfBill.TextMatrix(intLop, mconintCol�ɱ���)) <> Val(vsfBill.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                                MsgBox "��" & intLop & "��ҩƷ���������۹������̵������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                vsfBill.SetFocus
                                vsfBill.Row = intLop
                                vsfBill.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    Else
                        '������������ʱ
                        If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                            If CheckPriceAdjust(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol����))) = False Then
                                MsgBox "��" & intLop & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                vsfBill.SetFocus
                                vsfBill.Row = intLop
                                vsfBill.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function


Private Function SaveCard() As Boolean
    Dim lng������id As Long
    Dim int���ϵ�� As Integer
    Dim lng������ID As Integer
    Dim lng�������ID As Integer
    
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim str���� As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dbl�������� As Double
    Dim dblʵ������ As Double
    Dim dbl������ As Double
    Dim dbl�ۼ� As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim dat�������� As String
    Dim str�̵�ʱ�� As String
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim rs������ As New Recordset
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim dbl�ɱ��� As Double
    Dim n As Integer
    Dim str�ⷿ��λ As String
    Dim arrSql As Variant
    Dim i As Integer
    
    SaveCard = False
    arrSql = Array()
    On Error GoTo ErrHandle
    '����������������ID����Ҫ������ҩƷ��Ҫ����
    gstrSQL = "SELECT b.ϵ��,b.id AS ���id " _
            & "FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID AND a.���� = 14 "
    Set rs������ = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption)
    
    If rs������.EOF Then
        MsgBox "�Բ���û������ҩƷ�̵���������������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng������ID = 0
    lng�������ID = 0
    
    If rs������!ϵ�� = 1 Then lng������ID = rs������!���id
    rs������.Close
    
    If lng������ID = 0 Then
        MsgBox "�Բ���û������ҩƷ�̵��¼��������������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    With vsfBill
        chrNo = Trim(txtNo)
        lng�ⷿID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(62, lng�ⷿID)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        dat�������� = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
        str�̵�ʱ�� = TxtCheckDate.Text
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_ҩƷ�̵��¼��_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngҩƷID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                str���� = Trim(.TextMatrix(intRow, mconIntCol����))
                lng����ID = IIf(.TextMatrix(intRow, mconIntCol����) = "", 0, .TextMatrix(intRow, mconIntCol����))
                datЧ�� = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                dbl�������� = Val(.TextMatrix(intRow, mconintCol��������)) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1))
                
                dblʵ������ = Val(.TextMatrix(intRow, mconIntCol��λ����)) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0))
                
                If mbln��ͬ��λ = False Then
                    dblʵ������ = dblʵ������ + Val(.TextMatrix(intRow, mconIntColС��λ����)) * Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1))
                End If
                
                dbl������ = 0
                
'                If mbln��ͬ��λ = False Then
'                    dbl�ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1)), gtype_UserDrugDigits.Digit_���ۼ�)
'                    dbl�ۼ� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1)), gtype_UserDrugDigits.Digit_���ۼ�)
'                Else
'                    dbl�ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0)), gtype_UserDrugDigits.Digit_���ۼ�)
'                    dbl�ۼ� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0)), gtype_UserDrugDigits.Digit_���ۼ�)
'                End If

                dbl�ۼ� = Get�̵�ʱ���ۼ�(Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(1) = 1, lngҩƷID, lng�ⷿID, lng����ID, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                
                '�����۸�ʱȥ����۸񣬲�����������ʱȡԭʼ�۸�
                If lng����ID = -1 Then
                    If mbln��ͬ��λ = False Then
                        dbl�ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(1)), gtype_UserDrugDigits.Digit_���ۼ�)
                    Else
                        dbl�ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) / Val(Split(.TextMatrix(intRow, mconIntCol����ϵ��), "|")(0)), gtype_UserDrugDigits.Digit_���ۼ�)
                    End If
                Else
                    dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����ID, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                dbl���� = Val(.TextMatrix(intRow, mconintCol����))
                dbl��۲� = Val(.TextMatrix(intRow, mconintCol��۲�))
                dbl����� = Val(.TextMatrix(intRow, mconIntColʵ�ʽ��))
                dbl����� = Val(.TextMatrix(intRow, mconIntColʵ�ʲ��))
                str�ⷿ��λ = IIf(Trim(.TextMatrix(intRow, mconIntCol�ⷿ��λ)) = "", "", .TextMatrix(intRow, mconIntCol�ⷿ��λ))
                
                If dbl�������� <= dblʵ������ Then
                    lng������id = lng������ID
                    int���ϵ�� = 1
                Else
                    lng������id = lng�������ID
                    int���ϵ�� = -1
                End If
                 
                lng��� = intRow
                
                'zl_ҩƷ�̵��¼��_INSERT( /*NO_IN*/, /*���_IN*/, /*�ⷿID_IN*/, /*����_IN*/,
                    '/*������ID_IN*/, /*���ϵ��_IN*/, /*ҩƷID_IN*/, /*��������_IN*/,
                    '/*ʵ������_IN*/, /*������_IN*/, /*�ۼ�_IN*/, /*����_IN*/, /*��۲�_IN*/,
                    '/*������_IN*/, /*��������_IN*/, /*ժҪ_IN*/, /*����_IN*/, /*����_IN*/,
                    '/*Ч��_IN*/, /*�̵�ʱ��_IN*/ );
                
                gstrSQL = "zl_ҩƷ�̵��¼��_INSERT('" & chrNo & "'," & lng��� & "," & lng�ⷿID & "," & lng����ID & "," _
                    & lng������id & "," & int���ϵ�� & "," & lngҩƷID & "," & dbl�������� & "," _
                    & dblʵ������ & "," & dbl������ & "," & dbl�ۼ� & "," & dbl���� & "," & dbl��۲� & ",'" _
                    & str������ & "',to_date('" & dat�������� & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & strժҪ & "','" & str���� & "','" & str���� & "'," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & str�̵�ʱ�� & "'," & dbl����� & "," & dbl����� & ",'" & str��׼�ĺ� & "'," & dbl�ɱ��� & ",'" & str�ⷿ��λ & "')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
End Sub

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    Dim dblʵ������ As Double
    
    On Error GoTo ErrHandle
    If Not zlStr.IsHavePrivs(mstrPrivs, "�鿴�̵㵥���") Then Exit Sub
    
    With vsfBill
        If .TextMatrix(.Row, mconIntColҩ��) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(vsfBill.Row, 0) = "" Then Exit Sub
        gstrSQL = "select ��������/" & Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(1) & " as  ��������, " & _
            " ʵ������/" & Split(.TextMatrix(.Row, mconIntCol����ϵ��), "|")(1) & " as  ʵ������ " & _
            " from ҩƷ��� where �ⷿid=[1] " _
            & " and ҩƷid=[2] " _
            & " and ����=1 and " _
            & " nvl(����,0)=[3]"
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", txtStock.Tag, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
        Else
            .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            dblʵ������ = IIf(IsNull(rsUseCount!ʵ������), 0, rsUseCount!ʵ������)
        End If
        rsUseCount.Close
        
        staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(dblʵ������, mintNumberDigit1, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    OS.OpenIme True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    OS.OpenIme
End Sub

Private Function SetPhiscRows(ByVal lngID As Long, ByVal lng���� As Long, ByVal str��׼�ĺ� As String, Optional ByVal blnBatch As Boolean = False) As Boolean
'���ܣ�����ҩƷID���̴������ʾ�������ҩƷ�ĳ�ʼ�̴���Ϣ
'˵����
'   1.����ǷǷ�������ҩ,���Ѿ�������,����ʾ���˳���
'   2.����Ƿ�������ҩ����ֱ����ҩ��δ����ĸ����ο���С�
    Dim i As Integer, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsPrice As New Recordset
    Dim strҩ�� As String
    Dim str�̵�ʱ�� As String
     
    On Error GoTo errH
    
    str�̵�ʱ�� = TxtCheckDate.Text
    
    SetPhiscRows = False
    Set rsData = GetPhysicDetail(txtStock.Tag, lngID, Not blnBatch)
    intRecordCount = rsData.RecordCount
    If intRecordCount = 0 Then Exit Function
    '��������ҩƷ
    If lng���� <> -1 Then
        rsData.MoveFirst
        rsData.Find "����=" & lng����
        If rsData.EOF Then Exit Function
    End If
    
    With vsfBill
        intRow = .Row
        intCurrentRow = .Row
        
        vsfBill.Redraw = flexRDNone
        
        .TextMatrix(intRow, 0) = rsData!ҩƷid
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = rsData!ͨ����
        Else
            strҩ�� = IIf(IsNull(rsData!��Ʒ��), rsData!ͨ����, rsData!��Ʒ��)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = rsData!ҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = rsData!ҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsData!��Ʒ��), "", rsData!��Ʒ��)
        
        If .Col = mconIntColҩ�� Then
            .EditText = .TextMatrix(intRow, mconIntColҩ��)
        End If

        .TextMatrix(intRow, mconIntCol��Դ) = Nvl(rsData!ҩƷ��Դ)
        .TextMatrix(intRow, mconIntCol����ҩ��) = Nvl(rsData!����ҩ��)
        .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsData!���), "", rsData!���)
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
        
        'ȡ��ҩƷ�Ĳ���
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
        If .TextMatrix(intRow, mconIntCol����) = "" Then .TextMatrix(intRow, mconIntCol����) = Nvl(rsData!ȱʡ����)
        
        .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
        .TextMatrix(intRow, mconIntCol��λ) = IIf(IsNull(rsData.Fields(Split(mstr��λ, "|")(1)).Value), "", rsData.Fields(Split(mstr��λ, "|")(1)).Value)
        
        If lng���� = -1 Then
            .TextMatrix(intRow, mconIntCol����) = lng����
            .TextMatrix(intRow, mconIntCol����) = ""
            .TextMatrix(intRow, mconIntColЧ��) = ""
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        Else
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "0", rsData!����)
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                '����Ϊ��Ч��
                .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
            End If
                
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsData!��׼�ĺ�), "", rsData!��׼�ĺ�)
        End If
        
        .TextMatrix(intRow, mconIntCol��λ����) = ""
        .TextMatrix(intRow, mconIntColС��λ����) = ""
        .TextMatrix(intRow, mconintCol��λ) = IIf(IsNull(rsData.Fields(Split(mstr��λ, "|")(0)).Value), "", rsData.Fields(Split(mstr��λ, "|")(0)).Value)
        .TextMatrix(intRow, mconintColС��λ) = IIf(IsNull(rsData.Fields(Split(mstr��λ, "|")(1)).Value), "", rsData.Fields(Split(mstr��λ, "|")(1)).Value)
        .TextMatrix(intRow, mconintCol����_�ϼ�) = ""
        .TextMatrix(intRow, mconintCol��λ_�ϼ�) = IIf(IsNull(rsData!�ۼ۵�λ), "", rsData!�ۼ۵�λ)
        .TextMatrix(intRow, mconIntCol����ϵ��) = ��ȡ����ϵ��(rsData)
        .TextMatrix(intRow, mconIntcol�ӳ���) = rsData!�ӳ��� / 100 & "||" & rsData!�Ƿ��� & "||" & rsData!ҩ����������
        
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Nvl(rsData!�ۼ�, 0) * rsData.Fields(Replace(Split(mstr��λ, "|")(1), "��λ", "ϵ��")).Value, mintPriceDigit, , True)
        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Nvl(rsData!�ɱ���, 0) * rsData.Fields(Replace(Split(mstr��λ, "|")(1), "��λ", "ϵ��")).Value, mintPriceDigit, , True)
        
        If rsData!�Ƿ��� = 1 Then
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Get�̵�ʱ�����ۼ�(CLng(rsData!ҩƷid), txtStock.Tag, CLng(IIf(IsNull(rsData!����), "0", rsData!����)), rsData.Fields(Replace(Split(mstr��λ, "|")(1), "��λ", "ϵ��")).Value, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss"))), mintPriceDigit, , True)
        End If

        .RowData(intRow) = Val(IIf(IsNull(rsData!���Ч��), 0, rsData!���Ч��))
        rsData.MoveNext
        
        If blnBatch = False Then
            Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        End If
        
        .Col = IIf(lng���� = -1, mconIntCol����, mconIntCol��λ����)
        .EditCell
        
        vsfBill.Redraw = flexRDDirect
    End With
    
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
'    Dim strUnit As String
'    Dim int��λϵ�� As Integer
'    Dim StrNo As String
'
'    strUnit = GetDrugUnit(txtStock.Tag)
'    Select Case strUnit
'        Case "סԺ��λ"
'            int��λϵ�� = 1
'        Case "���ﵥλ"
'            int��λϵ�� = 2
'        Case "ҩ�ⵥλ"
'            int��λϵ�� = 3
'        Case "�ۼ۵�λ"             '�ۼ۵�λ����Ҫ���Ƽ���
'            int��λϵ�� = 4
'    End Select
'    StrNo = txtNo
'    Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), mint��¼״̬, int��λϵ��, 1307, "ҩƷ�̵㵥", StrNo)
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Set rsBatchNolen = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "-ȡ���ų���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ��ȡ��λ() As String
    Dim intUnit As Integer, strUnit As String, strDefault As String
    Dim strCompare As String
    Dim int���� As Integer
    
    Const conInt���㾫�� As Integer = 0
    
    Const conIntҩƷ As Integer = 1
    
    Const conint�ۼ۵�λ As Integer = 1
    Const conint���ﵥλ As Integer = 2
    Const conintסԺ��λ As Integer = 3
    Const conintҩ�ⵥλ As Integer = 4
    
    Const conInt�ɱ��� As Integer = 1
    Const conInt�ۼ� As Integer = 2
    Const conInt���� As Integer = 3
    Const conInt��� As Integer = 4
    
    int���� = conInt���㾫��
    
    strCompare = "ҩ�ⵥλ;���ﵥλ;סԺ��λ;�ۼ۵�λ"
    'ȡ��ȱʡ��λ
    strDefault = GetDrugUnit(Val(txtStock.Tag), "ҩƷ�̵����")
    
    'ȡ�̵㵥��ָ����λ
    intUnit = Val(zlDatabase.GetPara("С��װ��λ", glngSys, ģ���.ҩƷ�̵�))
    
    If intUnit = 0 Then
        strUnit = strDefault
    Else
        strUnit = Split(strCompare, ";")(intUnit - 1)
    End If
    
    '��ָ����λ��ȱʡ��λ����λ��С��λ��˳������
    mintDefault = 1
    If strUnit <> strDefault Then
        If InStr(1, strCompare, strUnit) < InStr(1, strCompare, strDefault) Then
            ��ȡ��λ = strUnit & "|" & strDefault
        Else
            mintDefault = 0
            ��ȡ��λ = strDefault & "|" & strUnit
        End If
    Else
        ��ȡ��λ = strUnit & "|" & strDefault
    End If
    
    mstr��λ = ��ȡ��λ
    
    'ȡ��λ�ľ��ȣ��ۼۡ���������
    Select Case Split(mstr��λ, "|")(0)
        Case "�ۼ۵�λ"
            intUnit = conint�ۼ۵�λ
        Case "���ﵥλ"
            intUnit = conint���ﵥλ
        Case "סԺ��λ"
            intUnit = conintסԺ��λ
        Case "ҩ�ⵥλ"
            intUnit = conintҩ�ⵥλ
    End Select
    
    mintCostDigit = GetDigit(int����, conIntҩƷ, conInt�ɱ���, intUnit)
    mintPriceDigit = GetDigit(int����, conIntҩƷ, conInt�ۼ�, intUnit)
    mintNumberDigit0 = GetDigit(int����, conIntҩƷ, conInt����, intUnit)
    mintMoneyDigit = GetDigit(int����, conIntҩƷ, conInt���)
    
    'ȡС��λ�ľ��ȣ�������
    Select Case Split(mstr��λ, "|")(1)
        Case "�ۼ۵�λ"
            intUnit = conint�ۼ۵�λ
        Case "���ﵥλ"
            intUnit = conint���ﵥλ
        Case "סԺ��λ"
            intUnit = conintסԺ��λ
        Case "ҩ�ⵥλ"
            intUnit = conintҩ�ⵥλ
    End Select
    mintNumberDigit1 = GetDigit(int����, conIntҩƷ, conInt����, intUnit)
    
    mbln��ͬ��λ = False
    If Split(mstr��λ, "|")(0) = Split(mstr��λ, "|")(1) Then
        mbln��ͬ��λ = True
    End If
End Function
Private Function ��ȡ����ϵ��(ByVal rsData As ADODB.Recordset) As String
    ��ȡ����ϵ�� = Replace(mstr��λ, "��λ", "ϵ��")
    ��ȡ����ϵ�� = rsData.Fields(Split(��ȡ����ϵ��, "|")(0)).Value & "|" & rsData.Fields(Split(��ȡ����ϵ��, "|")(1)).Value
End Function

Private Function GetPhysicDetail(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long, _
    Optional ByVal bln���޿��ҩƷ As Boolean = True, Optional ByVal bln�����̵㵥 As Boolean = False) As ADODB.Recordset
    'bln���޿��ҩƷ=�Ƿ��޿��ҩƷҲ��ȡ����
    'bln�����̵㵥=�Ƿ���Ҫ����ָ���̵�ʱ����̵㵥�γ��̵��
    '��ȡ��ҩƷ��ǰ�ⷿ����������ϸ��¼
    Dim str��λ As String, str�̵�ʱ�� As String, str�����̵㵥 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    str�̵�ʱ�� = TxtCheckDate.Text
    str��λ = ",A.סԺ��λ,A.סԺ��װ AS סԺϵ��"
    str��λ = str��λ & ",A.���ﵥλ,A.�����װ AS ����ϵ��"
    str��λ = str��λ & ",A.ҩ�ⵥλ,A.ҩ���װ AS ҩ��ϵ��"
    str��λ = str��λ & ",E.���㵥λ AS �ۼ۵�λ,1 As �ۼ�ϵ��"
    
    '�����̵㵥��SQL
    If bln�����̵㵥 Then
        str�����̵㵥 = "" & _
            " UNION ALL" & _
            " SELECT A.�ⷿID,A.ҩƷID,NVL(A.����, 0) AS ����,0 AS ʵ������,SUM(A.����) �̵�����," & _
                    " 0 AS ʵ�ʽ��,0 AS ʵ�ʲ��,0 AS ��������,A.����,A.����,A.Ч��,A.��׼�ĺ�" & _
            " FROM ҩƷ�շ���¼ A" & _
            " Where A.����=14 AND A.�ⷿID=[1] AND A.Ƶ��=[3] " & _
            " GROUP BY A.�ⷿID,A.ҩƷID,A.����,A.����,A.����,A.Ч��,A.��׼�ĺ�"
    End If
    
    'ȡҩƷ��ǰ��漰�̵�ʱ���Ժ�ľ�������
    gstrSQL = "" & _
        " SELECT DISTINCT A.ҩƷID,A.�ɱ��� As ƽ���ɱ���,E.���� ȱʡ����,'[' || E.���� || ']' As ҩƷ����, E.���� As ͨ����, C.���� As ��Ʒ��," & _
        "   A.ҩƷ��Դ,A.����ҩ��,A.ҩ����� AS ��������,A.ҩ������ AS ҩ����������,E.�Ƿ���,A.�ӳ���," & _
        "   NVL(B.ʵ�ʽ��,0) ʵ�ʽ��,NVL(B.ʵ�ʲ��,0) ʵ�ʲ��,D.�ּ� �ۼ�,NVL(B.����,0) ����,B.����,B.Ч��,F.�ⷿ��λ,E.���, decode(b.����,null,decode(a.�ϴβ���,null,e.����,a.�ϴβ���),b.����) as ����,A.���Ч��," & _
        "   B.��׼�ĺ�,B.��������,B.�̵�����,B.��������" & str��λ & ",Decode(sign(NVL(b.��������,0)), 1,Decode(x.�ּ�,Null,Decode(k.�ɱ���, Null, a.�ɱ���, k.�ɱ���),x.�ּ�), Decode(x.�ּ�,Null,a.�ɱ���,x.�ּ�)) �ɱ��� " & _
        " FROM ҩƷ��� A,�շ���ĿĿ¼ E,�շ���Ŀ���� C,�շѼ�Ŀ D,ҩƷ�����޶� F," & _
        "     (SELECT �ⷿID, ҩƷID, ����, SUM (ʵ������) AS ��������,SUM (�̵�����) AS �̵�����,SUM (ʵ�ʽ��) AS ʵ�ʽ��," & _
        "         SUM (ʵ�ʲ��) AS ʵ�ʲ��, SUM(��������) AS ��������,MAX(����) AS ����,MAX(����) AS ���� ,MAX(Ч��) AS Ч��,��׼�ĺ�" & _
        "         From" & _
        "             ( SELECT A.�ⷿID,A.ҩƷID,NVL(����,0) AS ����,Nvl(A.ʵ������,0) ʵ������,0 �̵�����,Nvl(A.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(A.ʵ�ʲ��,0) ʵ�ʲ��,Nvl(A.��������,0) ��������,A.�ϴ����� AS ����,A.�ϴβ��� AS ����,A.Ч��,A.��׼�ĺ�" & _
        "             FROM ҩƷ��� A" & _
        "             Where A.���� = 1 And A.�ⷿID=[1] And A.ҩƷID=[2] " & _
        "             Union All" & _
        "             SELECT A.�ⷿID,A.ҩƷID,NVL(A.����,0) AS ����,SUM(-1*A.���ϵ��*A.ʵ������*A.����) AS ʵ������,0 �̵�����," & _
        "             SUM (-1*A.���ϵ��*A.���۽��) AS ʵ�ʽ��, SUM(-1*A.���ϵ��*A.���) AS ʵ�ʲ��,0 AS ��������,A.����,A.����,A.Ч��,A.��׼�ĺ�" & _
        "             FROM ҩƷ�շ���¼ A" & _
        "             Where A.�ⷿID+0=[1] And A.ҩƷID+0=[2] " & _
        "             AND A.������� >[4] " & _
        "             GROUP BY A.�ⷿID, A.ҩƷID, A.����,A.����,A.����,A.Ч��,A.��׼�ĺ� " & IIf(Not bln�����̵㵥, "", str�����̵㵥) & _
        "     ) GROUP BY �ⷿID, ҩƷID, ����,��׼�ĺ�) B,(Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and [4] between x.ִ������ and x.��ֹ����) X," & _
        "      (Select ҩƷid,����,ƽ���ɱ��� �ɱ��� From ҩƷ��� Where ���� = 1 And �ⷿid =[1]) K " & _
        " Where A.ҩƷID+0=[2] And A.ҩƷID=E.ID And A.ҩƷID=B.ҩƷID" & IIf(bln���޿��ҩƷ, "(+)", "") & _
        " AND A.ҩƷID=F.ҩƷID(+) AND F.�ⷿID(+)=[1] And B.ҩƷid=K.ҩƷid(+) And Nvl(B.����, 0)=nvl(K.����(+),0)" & _
        " AND A.ҩƷID=C.�շ�ϸĿID(+) AND C.����(+)=3 And b.ҩƷid = x.ҩƷid(+) And b.�ⷿid = x.�ⷿid(+) And Nvl(b.����, 0) = Nvl(x.����(+), 0) " & GetPriceClassString("D") & _
        " AND A.ҩƷID=D.�շ�ϸĿID(+) AND D.ִ������(+)<=SYSDATE AND NVL(D.��ֹ����(+),SYSDATE)>=SYSDATE"
        
        gstrSQL = gstrSQL & " and e.����ʱ�� <= [4] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ��ǰ�ⷿ����������ϸ��¼]", lng�ⷿID, lngҩƷID, str�̵�ʱ��, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
    
    Set GetPhysicDetail = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Double
    Dim strKey As String
    
    With vsfBill
        Select Case Col
            Case mconIntCol��λ����, mconIntColС��λ����
                '��ʾ�ϼ�����
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                If .EditText <> "" Then .TextMatrix(Row, Col) = Val(.EditText)
                If .Col = mconIntCol��λ���� Then
                    lngSum = Val(.TextMatrix(Row, mconIntColС��λ����)) + Val(.TextMatrix(Row, mconIntCol��λ����)) * Val(Split(.TextMatrix(Row, mconIntCol����ϵ��), "|")(0)) / Val(Split(.TextMatrix(Row, mconIntCol����ϵ��), "|")(1))
                Else
                    lngSum = Val(.TextMatrix(Row, mconIntColС��λ����)) + Val(.TextMatrix(Row, mconIntCol��λ����)) * Val(Split(.TextMatrix(Row, mconIntCol����ϵ��), "|")(0)) / Val(Split(.TextMatrix(Row, mconIntCol����ϵ��), "|")(1))
                End If
                .TextMatrix(Row, mconintCol����_�ϼ�) = zlStr.FormatEx(lngSum * Val(Split(.TextMatrix(Row, mconIntCol����ϵ��), "|")(1)), mintNumberDigit1, , True)
            Case mconintCol�ɱ���
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                .EditText = zlStr.FormatEx(Val(.EditText), mintCostDigit, , True)
                strKey = Trim(.EditText)

                If strKey <> "" And .TextMatrix(Row, 0) <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .EditText = strKey

                    If Split(.TextMatrix(Row, mconIntcol�ӳ���), "||")(1) = 1 Then
                        'ʱ��ҩƷʱ
                        If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                            '���۹����ۼ۵��ڳɱ���
                            .TextMatrix(Row, mconIntCol�ۼ�) = strKey
                        End If
                    Else
                        '����ҩƷ
                        If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                            '���۹���Ҫ�жϳɱ����Ƿ�����ۼ�
                            If Val(strKey) <> Val(.TextMatrix(Row, mconIntCol�ۼ�)) Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & .TextMatrix(Row, mconIntCol�ۼ�) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(Row, mconIntCol�ۼ�)
                                .TextMatrix(.Row, mconintCol�ɱ���) = zlStr.FormatEx(strKey, mintCostDigit, , True)
                                .EditText = strKey
                            End If
                        End If
                    End If
                End If
                
                .TextMatrix(Row, Col) = .EditText
        End Select
    End With
End Sub


Private Function Get�̵�ʱ�����ۼ�(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double, ByVal date�̵�ʱ�� As Date) As Double
    '���ܣ���ȡָ��ʱ��ʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����,�̵�ʱ��
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle
    '1���ж�ҩƷ�۸��¼�Ƿ�������
    gstrSQL = "select �ּ� as ���ۼ� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 1 and [4] between ִ������ and ��ֹ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
    
    If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
    
        gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
            '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get�̵�ʱ�����ۼ� = 0
            dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get�̵�ʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            If rsData!���ۼ� = 0 Then
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�̵�ʱ�����ۼ� = 0
                dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�̵�ʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
            Else
                Get�̵�ʱ�����ۼ� = rsData!���ۼ� * dbl����ϵ��
            End If
        End If
    Else '�ж�ӦҩƷ�۸��¼
        Get�̵�ʱ�����ۼ� = rsData!���ۼ� * dbl����ϵ��
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�̵�ʱ���ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal date�̵�ʱ�� As Date) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get�̵�ʱ���ۼ�-ȡ����ҩƷ�ۼ�", lngҩƷID)
        
        If Not rsData.EOF Then
            Get�̵�ʱ���ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        '1���ж�ҩƷ�۸��¼�Ƿ�������
        gstrSQL = "select �ּ� as ���ۼ� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 1 and [4] between ִ������ and ��ֹ����"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
        
        If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
        
            gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
                " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷID, lng�ⷿID, lng����)
            
            If rsData.EOF Then
                'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
                '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�̵�ʱ���ۼ� = 0
                dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�̵�ʱ���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
            Else
                If rsData!���ۼ� = 0 Then
                    gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                    dblָ�����ۼ� = rsData!ָ�����ۼ�
                    dbl��������� = rsData!���������
                    
                    Get�̵�ʱ���ۼ� = 0
                    dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                    dbl�ӳ��� = rsData!�ӳ��� / 100
                    dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                    dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                    Get�̵�ʱ���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
                Else
                    Get�̵�ʱ���ۼ� = rsData!���ۼ�
                End If
            End If
        Else
            Get�̵�ʱ���ۼ� = rsData!���ۼ�
        End If
        
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�̵�ʱ�̳ɱ���(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal date�̵�ʱ�� As Date) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo ErrHandle
    
    '1���ж�ҩƷ�۸��¼�Ƿ�������
    gstrSQL = "select �ּ� as �ɱ��� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 2 and [4] between ִ������ and ��ֹ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
    
    If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
    
        gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ����=1 and ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            blnNullPrice = True
        ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
            blnNullPrice = True
        ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
            blnNullPrice = True
        End If
        
        If Not blnNullPrice Then
            Get�̵�ʱ�̳ɱ��� = rsData!ƽ���ɱ���
        Else
            '����޷��ӿ����ȡ�ɱ��ۣ����ҩƷ�����ȡ
            gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID)
            If Not rsData.EOF Then
                If Val(Nvl(rsData!�ɱ���, 0)) > 0 Then
                    Get�̵�ʱ�̳ɱ��� = rsData!�ɱ���
                End If
            End If
        End If
    Else
        Get�̵�ʱ�̳ɱ��� = rsData!�ɱ���
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



