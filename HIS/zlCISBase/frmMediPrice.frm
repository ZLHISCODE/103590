VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediPrice 
   Caption         =   "ҩƷ���۵�"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmMediPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14700
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picItem 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   5040
      ScaleHeight     =   2415
      ScaleWidth      =   5175
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   3720
         Picture         =   "frmMediPrice.frx":058A
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���(&A)"
         Height          =   350
         Left            =   2520
         Picture         =   "frmMediPrice.frx":06D4
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "��"
         Height          =   300
         Left            =   2450
         TabIndex        =   28
         Top             =   55
         Width           =   255
      End
      Begin VB.CheckBox ChkSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3000
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   25
         Top             =   60
         Width           =   1485
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSpec 
         Height          =   1200
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   4800
         _cx             =   8467
         _cy             =   2117
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediPrice.frx":081E
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
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "��ҩƷ��"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "ѡ��Ʒ��(&I)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0A3D
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPstor 
      Caption         =   "��ӡ���䶯��(&S)��"
      Height          =   350
      Left            =   8400
      Picture         =   "frmMediPrice.frx":0B87
      TabIndex        =   5
      Top             =   4200
      Width           =   1965
   End
   Begin TabDlg.SSTab sstabDetail 
      Height          =   4095
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "���䶯��(&S)"
      TabPicture(0)   =   "frmMediPrice.frx":0CD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTitle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "BillStore"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ӧ����䶯��(&P)"
      TabPicture(1)   =   "frmMediPrice.frx":0CED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BillPay"
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit BillStore 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   14737632
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit BillPay 
         Height          =   3555
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6271
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   14737632
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䶯��"
         Height          =   180
         Left            =   3240
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2505
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin ZL9BillEdit.BillEdit BillPrice 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4577
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)��"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0D09
      TabIndex        =   3
      Top             =   1161
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0E53
      TabIndex        =   2
      Top             =   663
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -435
      TabIndex        =   7
      Top             =   4060
      Width           =   16815
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0F9D
      TabIndex        =   4
      Top             =   1659
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":10E7
      TabIndex        =   1
      Top             =   165
      Width           =   1215
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Width           =   14535
      Begin VB.ComboBox cbo�ۼۼ��㷽ʽ 
         Height          =   300
         Left            =   11880
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   900
         Width           =   2415
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "ָ������ִ��"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   36
         Top             =   503
         Width           =   1695
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "����ִ��"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   35
         Top             =   503
         Width           =   1215
      End
      Begin VB.CheckBox chk��ҩ�������� 
         Caption         =   "ͬƷ��ҩƷ�۸�һ��(����������ʱ)"
         Height          =   210
         Left            =   10440
         TabIndex        =   31
         Top             =   525
         Width           =   3210
      End
      Begin VB.CheckBox chk�Զ����ɱ��� 
         Caption         =   "���ۼ�ʱ�Զ����ӳ��ʵ����ɱ���"
         Height          =   210
         Left            =   4680
         TabIndex        =   21
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chk�Զ�����Ӧ����䶯 
         Caption         =   "�Զ�����Ӧ����䶯"
         Height          =   210
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Chk���� 
         Caption         =   "ʱ��ҩƷ��Ϊ����"
         Enabled         =   0   'False
         Height          =   210
         Left            =   8520
         TabIndex        =   16
         Top             =   525
         Width           =   1770
      End
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Top             =   60
         Width           =   6765
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   8805
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   60
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtpRunDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   5880
         TabIndex        =   33
         Top             =   480
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   184745987
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lbl���۷�ʽ 
         AutoSize        =   -1  'True
         Caption         =   "�ۼۼ��㷽ʽ"
         Height          =   180
         Left            =   10680
         TabIndex        =   38
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lblִ��ʱ�� 
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "�޵���Ȩ�޲��ܵ��ۣ�"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   32
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   8175
         TabIndex        =   19
         Top             =   120
         Width           =   540
      End
   End
   Begin VB.Label lblHelp 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   12600
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMediPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngBillId As Long                '��������:0-���۴���;����-��ʾlngBillIdȷ������ʷ���۵�
Public lngMediId As Long                '��������:0-δָ������ҩƷ;����-����ʱֱ����ʾlngMediId��ԭ�۸����
Public lngItemID As Long                '��������:>0ʱ��Ʒ����ȡ���й��

Private blnModify As Boolean
Private blnFirst As Boolean
Private intDrugType As Integer          '1-��ҩ������ҩ���г�ҩ��;2-�в�ҩ
Private mstrPrivs As String
Private mstrAdjMsg As String            '����δִ�е��ۼ�¼��ҩƷ����ʾ��Ϣ
Private mblnAllUnAdj As Boolean         'Ʒ�ֶ�Ӧ�Ĺ�񶼴���δִ�м۸�
Private Const mlngColUpdate As Long = &H8000000F '���ܱ��޸ĵı�����ɫ
Private mstr���м�¼ As String          '��¼���������е����ݣ��������Ƿ�������޸�
Private mrs�ֶμӳ� As ADODB.Recordset    '��¼��������Щ�ӳ��ʶ�
Private mdbl�ֶμӳ��� As Double
Private mdbl�ɱ��� As Double            '��¼�޸�֮ǰ�ĳɱ���

'--------���۵��У��ۼ۵��ۣ�--------------
Private Enum �ۼ��б�
    ҩƷid = 0
    Ʒ�� = 1
    ��� = 2
    ���� = 3
    ��λ = 4
    ���� = 5
    �ϴ����� = 6
    ԭ�ɱ��� = 7
    �ֳɱ��� = 8
    ԭ�� = 9
    �ּ� = 10
    ������ID = 11
    ԭ����ID = 12
    �������� = 13
    ԭ�ɹ��޼� = 14
    �ֲɹ��޼� = 15
    ԭָ���ۼ� = 16
    ��ָ���ۼ� = 17
    �Ƿ��п�� = 18
    ����ϵ�� = 19
    ҩ��ID = 20
    ��װϵ�� = 21
    ��������� = 22
    �ӳ��� = 23
    ���� = 24
End Enum

'--------���䶯�У�ʱ��ҩƷ�����ε��ۼۡ��ɱ��۵��ۣ�--------------
Private Enum ����б�
    �ⷿ = 0
    ��Ӧ�� = 1
    ҩƷ = 2
    ��� = 3
    ��λ = 4
    ���� = 5
    Ч�� = 6
    ���� = 7
    ���� = 8
    ԭ�� = 9
    �ּ� = 10
    ������� = 11
    �ӳ��� = 12
    ԭ�ɱ��� = 13
    �ֳɱ��� = 14
    ��۲� = 15
    ���� = 16
    ��� = 17
    ҩƷid = 18
    �ⷿid = 19
    ��Ӧ��ID = 20

    ���� = 21
End Enum

'--------Ӧ�����У��ɱ��۵�����Ҫ����Ӧ����¼ʱ��--------------
Private Enum Ӧ������
    ҩƷid = 0
    Ʒ�� = 1
    ��Ʊ�� = 2
    ��Ʊ���� = 3
    ��Ʊ��� = 4
    
    ���� = 5
End Enum

Dim rsTemp As New ADODB.Recordset
Dim intCount As Integer
Dim objItem As ListItem
Dim objNode As Node
Dim dtToday As Date
Dim intҩ�ⵥλ As Integer      '�Ƿ���ҩ�ⵥλ��ʾ
Dim mstrNo As String            '���۵�No

Private mblnʱ��ҩƷ���� As Boolean         'ʱ��ҩƷ�����Ƿ�����ִ��
Private mbln�޼���ʾ As Boolean             '���ۼ۳����޼�ʱ�Ƿ���ʾ
Private mstrҩƷ As String
Private mlng���� As Long
Private mlngҩƷID As Long
Private mintCurRow As Integer
Private mintCurCol As Integer

'�Ӳ�������ȡҩƷ�۸�С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��
Private mstrMoneyFormat As String

Private mintSalePriceDigit As Integer

'���۵����д���Ĳ���
Private mint���� As Integer             '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���;3-������������Ŀ
Private mlng��Ӧ��ID As Long
Private mdbl�ӳ��� As Double
Private mblnӦ����¼ As Boolean         'False-������Ӧ����¼;True-����Ӧ����¼
Private Sub BatchAdjustPriceByItem(ByVal lngRow As Long, ByVal dblPrice As Double)
    '��Ʒ�ֵ��ۣ���ͬƷ�ֵĹ����ۼ۱���һ�£�������ϵ�����㣩����ʱ��֧���в�ҩ
    Dim lngҩ��id As Long
    Dim dbl����ϵ�� As Double
    Dim dbl��װϵ�� As Double
    Dim n As Long
    Dim dbl���� As Double
    Dim dbl�ּ� As Double

    If chk��ҩ��������.Visible = False Then Exit Sub
    If chk��ҩ��������.Value <> 1 Then Exit Sub
    
    With BillPrice
        lngҩ��id = Val(.TextMatrix(lngRow, �ۼ��б�.ҩ��ID))
        dbl����ϵ�� = Val(.TextMatrix(lngRow, �ۼ��б�.����ϵ��))
        dbl��װϵ�� = Val(.TextMatrix(lngRow, �ۼ��б�.��װϵ��))
        dbl���� = dblPrice / dbl��װϵ�� / dbl����ϵ��
        
        For n = 1 To .Rows - 1
            If Val(.TextMatrix(n, �ۼ��б�.ҩƷid)) > 0 Then
                If Val(.TextMatrix(n, �ۼ��б�.ҩ��ID)) = lngҩ��id And n <> lngRow Then
                    dbl�ּ� = dbl���� * Val(.TextMatrix(n, �ۼ��б�.��װϵ��)) * Val(.TextMatrix(n, �ۼ��б�.����ϵ��))
                    
                    '�ּ۴���ָ���ۼ�ʱ����ʾ�Ƿ����
                    If mbln�޼���ʾ = True Then
                        If .TextMatrix(n, �ۼ��б�.����) = "����" And dbl�ּ� > Val(BillPrice.TextMatrix(n, �ۼ��б�.��ָ���ۼ�)) Then
                           MsgBox .TextMatrix(n, �ۼ��б�.Ʒ��) & "�ּ۸���ָ�����ۼ�" & Val(BillPrice.TextMatrix(n, �ۼ��б�.��ָ���ۼ�)) & "���ɹ��޼۽��Ͳɹ���һ�£�", vbInformation, gstrSysName
                        End If
                    End If
            
                    .TextMatrix(n, �ۼ��б�.�ּ�) = dbl�ּ�
                    If dbl�ּ� > Val(BillPrice.TextMatrix(n, �ۼ��б�.��ָ���ۼ�)) Then
                        .TextMatrix(.Row, �ۼ��б�.��ָ���ۼ�) = FormatEx(dbl�ּ�, mintPriceDigit)
                    End If
                    
                    Call ChangeDrugStore(n, Val(.TextMatrix(n, �ۼ��б�.ҩƷid)), dbl�ּ�)
                End If
            End If
        Next
    End With
End Sub

Private Sub BatchAdjustCostByItem(ByVal lngRow As Long, ByVal dblCost As Double)
    '��Ʒ�ֵ��ۣ���ͬƷ�ֵĹ��ĳɱ��۱���һ�£�������ϵ�����㣩����ʱ��֧���в�ҩ
    Dim lngҩ��id As Long
    Dim dbl����ϵ�� As Double
    Dim dbl��װϵ�� As Double
    Dim n As Long
    Dim dbl���� As Double
    Dim dbl�ּ� As Double

    If chk��ҩ��������.Visible = False Then Exit Sub
    If chk��ҩ��������.Value <> 1 Then Exit Sub
    
    With BillPrice
        lngҩ��id = Val(.TextMatrix(lngRow, �ۼ��б�.ҩ��ID))
        dbl����ϵ�� = Val(.TextMatrix(lngRow, �ۼ��б�.����ϵ��))
        dbl��װϵ�� = Val(.TextMatrix(lngRow, �ۼ��б�.��װϵ��))
        dbl���� = dblCost / dbl��װϵ�� / dbl����ϵ��
        
        For n = 1 To .Rows - 1
            If Val(.TextMatrix(n, �ۼ��б�.ҩƷid)) > 0 Then
                If Val(.TextMatrix(n, �ۼ��б�.ҩ��ID)) = lngҩ��id And n <> lngRow Then
                    dbl�ּ� = dbl���� * Val(.TextMatrix(n, �ۼ��б�.��װϵ��)) * Val(.TextMatrix(n, �ۼ��б�.����ϵ��))
                    
                    '�ּ۴���ָ���ۼ�ʱ����ʾ�Ƿ����
                    If mbln�޼���ʾ = True Then
                        If dbl�ּ� > Val(BillPrice.TextMatrix(n, �ۼ��б�.�ֲɹ��޼�)) Then
                            MsgBox .TextMatrix(n, �ۼ��б�.Ʒ��) & "�ֳɱ��۸���ָ���ɹ��޼�" & Val(BillPrice.TextMatrix(n, �ۼ��б�.�ֲɹ��޼�)) & "���ɹ��޼۽��Ͳɹ���һ�£�", vbInformation, gstrSysName
                        End If
                    End If
            
                    .TextMatrix(n, �ۼ��б�.�ֳɱ���) = dbl�ּ�
                    
                    If dbl�ּ� > Val(BillPrice.TextMatrix(n, �ۼ��б�.�ֲɹ��޼�)) Then
                        .TextMatrix(.Row, �ۼ��б�.�ֲɹ��޼�) = FormatEx(dbl�ּ�, mintPriceDigit)
                    End If
                    
                    Call CaculateCost(Val(.TextMatrix(n, �ۼ��б�.ҩƷid)), dbl�ּ�)
                End If
            End If
        Next
    End With
End Sub

Private Sub CaculateCost(ByVal lngҩƷID As Long, ByVal dbl�ֳɱ��� As Double)
    Dim n As Integer
    Dim dbl��Ʊ��� As Double
    
    With BillStore
        For n = 1 To .Rows - 1
            If .TextMatrix(n, ����б�.ҩƷid) <> "" Then
                If Val(.TextMatrix(n, ����б�.ҩƷid)) = lngҩƷID Then
                    .TextMatrix(n, ����б�.�ֳɱ���) = FormatEx(dbl�ֳɱ���, mintCostDigit)
                    If dbl�ֳɱ��� <> 0 Then
                        .TextMatrix(n, ����б�.�ӳ���) = FormatEx((Val(.TextMatrix(n, ����б�.�ּ�)) / dbl�ֳɱ��� - 1) * 100, 5)
                    End If
                    If cbo�ۼۼ��㷽ʽ = "�ۼ۰��ֶμӳɼ���" Then
                        .TextMatrix(n, ����б�.�ӳ���) = FormatEx(mdbl�ֶμӳ��� * 100, 5)
                    End If
                    
                    .TextMatrix(n, ����б�.��۲�) = Format((dbl�ֳɱ��� - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����)), mstrMoneyFormat)
                        
                    dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֳɱ��� - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����))
                     
                    If (cbo�ۼۼ��㷽ʽ = "�ۼ۰��ֶμӳɼ���" Or cbo�ۼۼ��㷽ʽ = "�ۼ۰��̶���������") And BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.����) = "ʱ��" And mint���� = 2 Then
                        .TextMatrix(n, ����б�.�ּ�) = BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�)
                    End If
                End If
            End If
        Next
    End With
    
    If chk�Զ�����Ӧ����䶯.Value = 1 Then
        For n = 1 To BillPay.Rows - 1
            If BillPay.TextMatrix(1, 0) <> "" Then
                If Val(BillPay.TextMatrix(n, Ӧ������.ҩƷid)) = lngҩƷID Then
                    BillPay.TextMatrix(n, Ӧ������.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub

Private Sub CaluateAverCost(ByVal lngҩƷID As Long)
    '����ƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double
    
    With BillStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, ����б�.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, ����б�.ҩƷid)) = lngҩƷID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, ����б�.�ֳɱ���)) * Val(.TextMatrix(i, ����б�.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, ����б�.����))
                End If
            End If
        Next
    End With
    
    With BillPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, �ۼ��б�.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, �ۼ��б�.ҩƷid)) = lngҩƷID Then
                        .TextMatrix(i, �ۼ��б�.�ֳɱ���) = FormatEx(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugId As Long, ByVal dblNewPrice As Double)
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl��װ As Double
    Dim n As Integer
    Dim dbl��Ʊ��� As Double
    
    If intRow = 0 Or mint���� = 1 Then Exit Sub
    
    dblOldPrice = Val(BillPrice.TextMatrix(intRow, �ۼ��б�.ԭ��))
    dbl��װ = GetModulus(lngDrugId)
    
    With BillStore
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, ����б�.ҩƷid)) = lngDrugId Then
                    dblNum = Val(.TextMatrix(n, ����б�.����))
                    
                    .TextMatrix(n, ����б�.�ּ�) = FormatEx(dblNewPrice, mintPriceDigit)
                    .TextMatrix(n, ����б�.�������) = Format(Val(.TextMatrix(n, ����б�.����)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    
                    If mint���� = 2 And chk�Զ����ɱ���.Value = 1 Then
                        dblOldCost = .TextMatrix(n, ����б�.ԭ�ɱ���)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, ����б�.�ӳ���)) / 100, 7))
                        .TextMatrix(n, ����б�.�ֳɱ���) = FormatEx(dblNewCost, mintCostDigit)
                        .TextMatrix(n, ����б�.��۲�) = Format((dblNewCost - dblOldCost) * dblNum, mstrMoneyFormat)
                        dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * dblNum
                    End If
                End If
            End If
        Next
    End With
    
    If chk�Զ�����Ӧ����䶯.Value = 1 Then
        With BillPay
            For n = 1 To .Rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, Ӧ������.ҩƷid)) = lngDrugId Then
                        .TextMatrix(n, Ӧ������.��Ʊ���) = FormatEx(dbl��Ʊ���, 2)
                    End If
                End If
            Next
        End With
    End If
    
    CaluateAverCost lngDrugId
End Sub

Private Function CheckUnVerify(ByVal lngҩƷID As Long) As Boolean
    '���ҩƷ�Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select 1 From ҩƷ�շ���¼ Where ҩƷid = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "���ҩƷ�Ƿ����δ��˵���", lngҩƷID)
    
    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetBatchData(ByVal BlnAll As Boolean)
    Dim lngRow As Long
    Dim n As Long
    Dim blnRepeat As Boolean
    
    For lngRow = 1 To vsfSpec.Rows - 1
        blnRepeat = False
        
        If Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ҩƷid"))) > 0 Then
            If Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ѡ��"))) <> 0 Or BlnAll = True Then
                For n = 1 To BillPrice.Rows - 1
                    If BillPrice.TextMatrix(n, �ۼ��б�.ҩƷid) <> "" Then
                        If Val(BillPrice.TextMatrix(n, �ۼ��б�.ҩƷid)) = Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ҩƷid"))) Then
                            blnRepeat = True
                            Exit For
                        End If
                    End If
                Next
                
                '���ظ�������
                If blnRepeat = False Then


                    With BillPrice
                        If .TextMatrix(.Rows - 1, �ۼ��б�.ҩƷid) <> "" Then
                            .Rows = .Rows + 1
                        End If
                        .TextMatrix(.Rows - 1, �ۼ��б�.ҩƷid) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ҩƷid"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.Ʒ��) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ҩƷ"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.���) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("���"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.����) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("����"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.��λ) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("��λ"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.����) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("����"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.ԭ�ɱ���) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("�ɱ���"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.�ֳɱ���) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("�ɱ���"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.ԭ�ɹ��޼�) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("�ɹ��޼�"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.�ֲɹ��޼�) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("�ɹ��޼�"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.ԭָ���ۼ�) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ָ���ۼ�"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.��ָ���ۼ�) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ָ���ۼ�"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.����ϵ��) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("����ϵ��"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.ҩ��ID) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("ҩ��ID"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.��װϵ��) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("��װϵ��"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.���������) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("���������"))
                        .TextMatrix(.Rows - 1, �ۼ��б�.�ӳ���) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("�ӳ���"))
                        
                        Call zlGetPrice(.Rows - 1, .TextMatrix(.Rows - 1, �ۼ��б�.ҩƷid), IIf(.TextMatrix(.Rows - 1, �ۼ��б�.����) = "ʱ��", True, False))
                        
                        DoEvents
                        
                        Call GetDrugStore(.Rows - 1, Val(.TextMatrix(.Rows - 1, �ۼ��б�.ҩƷid)))
                    End With
                    
                    DoEvents
                End If
            End If
        End If
    Next
End Sub

Private Sub GetDrugStore(ByVal intRow As Integer, ByVal lngAddDrugId As Long, Optional ByVal lngDelDrugId As Long = 0)
    Dim n As Integer
    Dim intRows As Integer
    Dim dbl��װ As Double
    Dim dblOldPrice As Double
    Dim dblNewPrice As Double
    Dim strSql��Ӧ��ID As String
    Dim dbl�ӳ��� As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim strҩƷ���� As String
    Dim dbl��Ʊ��� As Double
    
    On Error GoTo errHandle
    If lngDelDrugId > 0 Then
        With BillStore
            For n = .Rows - 1 To 1 Step -1
                If Val(.TextMatrix(n, ����б�.ҩƷid)) = lngDelDrugId Then
                    .MsfObj.RemoveItem n
                End If
            Next
        End With
        
        If mint���� = 1 Or mint���� = 2 Then
            With BillPay
                For n = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(n, Ӧ������.ҩƷid)) = lngDelDrugId Then
                       .MsfObj.RemoveItem n
                    End If
                Next
            End With
        End If
    End If
    
    If lngAddDrugId = 0 Then Exit Sub
    
    With BillStore
        .Active = True
        dbl��װ = GetModulus(lngAddDrugId)
        dblOldPrice = Val(BillPrice.TextMatrix(intRow, �ۼ��б�.ԭ��))
        dblNewPrice = Val(BillPrice.TextMatrix(intRow, �ۼ��б�.�ּ�))
        
        If mint���� = 1 Or mint���� = 2 Then
            strSql��Ӧ��ID = IIf(mlng��Ӧ��ID = 0, "", " And S.�ϴι�Ӧ��ID=[2] ")
        End If
            
        gstrSql = "select S.�ⷿID,D.���� as �ⷿ,'['||M.����||']'||M.���� as ҩƷ,M.���,M.����,M.���㵥λ �ۼ۵�λ,p.ҩ�ⵥλ,S.����,S.����,S.����, Nvl(M.�Ƿ���, 0) ���, M.ID, S.ʱ���ۼ�,P.ָ������� As �����,S.�ɱ���,S.�ϴι�Ӧ��ID, N.���� As ��Ӧ��,S.Ч��,S.���� " & _
            " from (select S.�ⷿID,S.ҩƷID,S.�ϴι�Ӧ��ID,S.�ϴ����� ����,S.Ч��,S.�ϴβ��� As ����,S.ʵ������ as ����,S.����, Decode(Nvl(S.����,0),0,Nvl(S.ʵ�ʽ��,0) / S.ʵ������,Nvl(S.���ۼ�,Nvl(S.ʵ�ʽ��,0) / S.ʵ������)) ʱ���ۼ�, s.ƽ���ɱ��� As �ɱ���" & _
            "       from ҩƷ��� S" & _
            "       where S.����=1 and S.ʵ������<>0 and S.ҩƷid=[1] ) S, " & _
            "      ���ű� D,�շ���ĿĿ¼ M,ҩƷ��� P, ��Ӧ�� N " & _
            " where D.id=S.�ⷿid and S.ҩƷID=M.ID And M.ID=P.ҩƷID And Nvl(S.�ϴι�Ӧ��id, 0) = N.ID(+) " & _
            " order by �ⷿ,S.����"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAddDrugId, mlng��Ӧ��ID)
            
'        If rsTemp.RecordCount = 0 And mint���� = 1 Then
'            MsgBox "��ҩƷ�޿�棬���ܵ����ɱ���!", vbInformation, gstrSysName
'            If BillPrice.Rows = 2 Then
'                BillPrice.Rows = BillPrice.Rows + 1
'            End If
'            BillPrice.MsfObj.RemoveItem intRow
'            Exit Sub
'        End If
        
        intRows = .Rows - 1
        
        BillPrice.TextMatrix(intRow, �ۼ��б�.�Ƿ��п��) = IIf(rsTemp.EOF, 0, 1)
        
        If mlng��Ӧ��ID > 0 Then
            rsTemp.Filter = "�ϴι�Ӧ��ID=" & mlng��Ӧ��ID
        End If
        
        .Rows = .Rows + rsTemp.RecordCount
        
        Do While Not rsTemp.EOF
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�ⷿ) = rsTemp!�ⷿ
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.��Ӧ��) = NVL(rsTemp!��Ӧ��)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.ҩƷ) = rsTemp!ҩƷ
            strҩƷ���� = rsTemp!ҩƷ
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            If intҩ�ⵥλ = 0 Then
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.��λ) = IIf(IsNull(rsTemp!�ۼ۵�λ), "", rsTemp!�ۼ۵�λ)
            Else
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.��λ) = IIf(IsNull(rsTemp!ҩ�ⵥλ), "", rsTemp!ҩ�ⵥλ)
            End If
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����) = FormatEx(rsTemp!���� / dbl��װ, mintNumberDigit)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.Ч��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�ּ�) = FormatEx(dblNewPrice, mintPriceDigit)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����) = NVL(rsTemp!����, 0)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.���) = rsTemp!���
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.ҩƷid) = rsTemp!ID
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�ⷿid) = rsTemp!�ⷿid
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.��Ӧ��ID) = IIf(mlng��Ӧ��ID > 0, mlng��Ӧ��ID, NVL(rsTemp!�ϴι�Ӧ��ID))
            If mint���� = 1 Or mint���� = 2 Then
                dblOldCost = FormatEx(rsTemp!�ɱ��� * dbl��װ, mintCostDigit)
               
                If mdbl�ӳ��� > 0 Then
                    dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl�ӳ��� = Round(dblOldPrice / dblOldCost - 1, 7)
                Else
                    dbl�ӳ��� = Round(1 / (1 - rsTemp!����� / 100) - 1, 7)
                End If
                
'                If dblOldPrice = dblNewPrice Then
'                    dblNewCost = dblOldCost
'                Else
                    dblNewCost = dblNewPrice / (1 + dbl�ӳ���)
'                End If
                
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.ԭ��) = FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ, dblOldPrice), mintPriceDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�������) = Format(rsTemp!���� / dbl��װ * (dblNewPrice - IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ, dblOldPrice)), mstrMoneyFormat)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�ӳ���) = dbl�ӳ��� * 100
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.ԭ�ɱ���) = FormatEx(dblOldCost, mintCostDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�ֳɱ���) = FormatEx(dblNewCost, mintCostDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.��۲�) = Format((dblNewCost - dblOldCost) * Val(.TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����)), mstrMoneyFormat)
                dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * Val(.TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.����))
            Else
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.ԭ��) = FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ, dblOldPrice), mintPriceDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, ����б�.�������) = Format(rsTemp!���� / dbl��װ * (dblNewPrice - IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ, dblOldPrice)), mstrMoneyFormat)
            End If
            
            rsTemp.MoveNext
        Loop
    
    End With
    
    If mint���� = 1 Or mint���� = 2 Then
        With BillPay
            .Active = True
            .TextMatrix(.Rows - 1, Ӧ������.ҩƷid) = lngAddDrugId
            .TextMatrix(.Rows - 1, Ӧ������.Ʒ��) = strҩƷ����
            .TextMatrix(.Rows - 1, Ӧ������.��Ʊ���) = FormatEx(dbl��Ʊ���, 2)
            .Rows = .Rows + 1
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem(ByVal strKey As String)
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(txtItem.hWnd)
    sngX = picItem.Left + vRect.Left - 100
    sngY = picItem.Top + vRect.Top + txtItem.Height + 175
    sngH = picItem.Height - vsfSpec.Top
    
    If strKey = "" Then
        gstrSql = "Select Distinct I.ID, '[' || I.���� || ']' || I.���� As ҩƷ, I.���㵥λ " & _
            " From ������ĿĿ¼ I, ������Ŀ���� N " & _
            " Where I.ID = N.������Ŀid And I.��� = '7' And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By '[' || I.���� || ']' || I.����"
        Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "ҩƷѡ����", False, "", "ѡ��ҩƷ", False, False, True, sngX, sngY, sngH, blnCancel, False, False)
    Else
        gstrSql = "Select Distinct I.ID, '[' || I.���� || ']' || I.���� As ҩƷ, I.���㵥λ " & _
            " From ������ĿĿ¼ I, ������Ŀ���� N " & _
            " Where I.ID = N.������Ŀid And I.��� = '7' And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And (I.���� Like [1] Or N.���� Like [2] Or N.���� Like [2]) " & _
            " Order By '[' || I.���� || ']' || I.����"
         Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "ҩƷѡ����", False, "", "ѡ��ҩƷ", False, False, True, sngX, sngY, sngH, blnCancel, False, False, UCase(strKey) & "%", "%" & UCase(strKey) & "%")
    End If
    
    If blnCancel = True Then Exit Sub

    If Not rsTemp Is Nothing Then
        Call GetSpec(Val(rsTemp!ID), 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSpec(ByVal lngItem As Long, ByVal intType As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim dbl��װ As Double
    
    On Error GoTo errHandle
    gstrSql = "Select Distinct I.ID, I.����, I.����, I.���, I.����, I.���㵥λ, P.ҩ�ⵥλ, Decode(I.�Ƿ���, 1, 'ʱ��', '����') ����, Nvl(P.�ɱ���, 0) �ɱ���," & _
        " P.ָ�������� , P.ָ�����ۼ�, Z.���� As Ʒ��, P.����ϵ��, P.ҩ��ID,p.���������,1/(1-p.ָ�������/100)-1  �ӳ���" & _
        " From �շ���ĿĿ¼ I, �շ���Ŀ���� N, ҩƷ��� P, ������ĿĿ¼ Z " & _
        " Where I.ID = N.�շ�ϸĿid And I.��� In (" & IIf(intType = 1, "'5','6'", "'7'") & ") And I.ID = P.ҩƷid And P.ҩ��id = Z.ID And " & _
        " (I.����ʱ�� Is Null Or I.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And P.ҩ��id = [1] " & _
        " Order By I.���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "ȡҩƷ���", lngItem)
    
    With vsfSpec
        .Redraw = flexRDNone
'        .Rows = 1
'        .Rows = 2
        lngRow = .Rows - 1
        
        mblnAllUnAdj = True
        
        If rsTemp.RecordCount > 0 Then
            Do While Not rsTemp.EOF
                If Check����δִ�м۸�(Val(rsTemp!ID)) = False Then
                    mblnAllUnAdj = False
                    
                    dbl��װ = GetModulus(Val(rsTemp!ID))
                    
                    .TextMatrix(lngRow, .ColIndex("ҩƷid")) = rsTemp!ID
                    .TextMatrix(lngRow, .ColIndex("Ʒ��")) = rsTemp!Ʒ��
                    .TextMatrix(lngRow, .ColIndex("ҩƷ")) = "[" & rsTemp!���� & "]" & rsTemp!����
                    .TextMatrix(lngRow, .ColIndex("���")) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    
                    If intҩ�ⵥλ = 0 Then
                        .TextMatrix(lngRow, .ColIndex("��λ")) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
                    Else
                        .TextMatrix(lngRow, .ColIndex("��λ")) = IIf(IsNull(rsTemp!ҩ�ⵥλ), "", rsTemp!ҩ�ⵥλ)
                    End If
                    
                    .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("�ɱ���")) = FormatEx(Val(IIf(IsNull(rsTemp!�ɱ���), "", rsTemp!�ɱ���)) * dbl��װ, mintCostDigit)
                    .TextMatrix(lngRow, .ColIndex("�ɹ��޼�")) = FormatEx(Val(IIf(IsNull(rsTemp!ָ��������), "", rsTemp!ָ��������)) * dbl��װ, mintCostDigit)
                    .TextMatrix(lngRow, .ColIndex("ָ���ۼ�")) = FormatEx(Val(IIf(IsNull(rsTemp!ָ�����ۼ�), "", rsTemp!ָ�����ۼ�)) * dbl��װ, mintPriceDigit)
                    .TextMatrix(lngRow, .ColIndex("����ϵ��")) = rsTemp!����ϵ��
                    .TextMatrix(lngRow, .ColIndex("ҩ��ID")) = rsTemp!ҩ��ID
                    .TextMatrix(lngRow, .ColIndex("��װϵ��")) = dbl��װ
                    .TextMatrix(lngRow, .ColIndex("���������")) = IIf(IsNull(rsTemp!���������), 0, rsTemp!���������)
                    .TextMatrix(lngRow, .ColIndex("�ӳ���")) = rsTemp!�ӳ���
                                                        
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                Else
                    mstrAdjMsg = IIf(mstrAdjMsg = "", "", mstrAdjMsg & vbCrLf) & "[" & rsTemp!���� & "]" & rsTemp!����
                End If
                
                rsTemp.MoveNext
            Loop
        End If
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub IniBatchData()
    Dim strToday As String
    
    '������۱༭״̬
    Me.BillPrice.Active = True
    
    strToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    
    Me.lblTitle.Caption = "���䶯��(���ڵ���δ���棬��ӳ�Ŀ����ܲ�׼ȷ)"
    Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(strToday))
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(strToday))
    Me.txtValuer.Text = gstrUserName
    
    Call GetSpec(lngItemID, intDrugType)
    
    If mstrAdjMsg <> "" Then
        MsgBox "����ҩƷ����δִ�м۸񣬲����ٽ��е��۲�����" & vbCrLf & mstrAdjMsg, vbInformation, gstrSysName
    End If
    
    If mblnAllUnAdj = True Then
        '������й�񶼴���δִ�м۸����˳�
        Unload Me
    Else
        Call GetBatchData(True)
    End If
End Sub

Private Sub IniData()
    Dim strToday As String
    Dim dbl��װ As Double
    
    On Error GoTo errHandle
    strToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    If lngBillId = 0 Then
        '������۱༭״̬
        Me.BillPrice.Active = True
        
        Me.lblTitle.Caption = "���䶯��(���ڵ���δ���棬��ӳ�Ŀ����ܲ�׼ȷ)"
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(strToday))
        Me.dtpRunDate.Value = DateAdd("d", 1, CDate(strToday))
        Me.txtValuer.Text = gstrUserName
        
        If lngMediId = 0 Then Exit Sub
        
        If Check����δִ�м۸�(lngMediId) = True Then
            MsgBox "��ҩƷ����δִ�м۸񣬲����ٽ��е��۲���!", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        dbl��װ = GetModulus(lngMediId)
        
        '���ָ�����ȵ��۵�ҩƷ����ֱ�ӽ���ҩƷ����
        gstrSql = "select P.ҩ��ID,P.����ϵ��,I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,P.ҩ�ⵥλ,decode(I.�Ƿ���,1,'ʱ��','����') ����,Nvl(P.�ɱ���,0) �ɱ���,P.ָ��������,P.ָ�����ۼ�,p.���������,1/(1-p.ָ�������/100)-1 �ӳ���" & _
                 " from �շ���ĿĿ¼ I,ҩƷ��� P" & _
                 " where I.ID=[1] And I.ID=P.ҩƷID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
        
        With rsTemp
            If .BOF Or .EOF Then Exit Sub
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ҩƷid) = !ID
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.Ʒ��) = "[" & !���� & "]" & !����
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.���) = IIf(IsNull(!���), "", !���)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.����) = IIf(IsNull(!����), "", !����)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ԭ�ɱ���) = FormatEx(Val(IIf(IsNull(!�ɱ���), "", !�ɱ���)) * dbl��װ, mintCostDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�ֳɱ���) = FormatEx(Val(IIf(IsNull(!�ɱ���), "", !�ɱ���)) * dbl��װ, mintCostDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ԭ�ɹ��޼�) = FormatEx(Val(IIf(IsNull(!ָ��������), "", !ָ��������)) * dbl��װ, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�ֲɹ��޼�) = FormatEx(Val(IIf(IsNull(!ָ��������), "", !ָ��������)) * dbl��װ, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ԭָ���ۼ�) = FormatEx(Val(IIf(IsNull(!ָ�����ۼ�), "", !ָ�����ۼ�)) * dbl��װ, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��ָ���ۼ�) = FormatEx(Val(IIf(IsNull(!ָ�����ۼ�), "", !ָ�����ۼ�)) * dbl��װ, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ҩ��ID) = !ҩ��ID
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.����ϵ��) = !����ϵ��
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��װϵ��) = dbl��װ
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.���������) = IIf(IsNull(!���������), 0, !���������)
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�ӳ���) = !�ӳ���
            
            If intҩ�ⵥλ = 0 Then
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��λ) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            Else
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��λ) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
            End If
            Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.����) = IIf(IsNull(!����), "", !����)
            
            Call zlGetPrice(.AbsolutePosition, lngMediId, IIf(!���� = "ʱ��", True, False))
            
            If mint���� = 0 Then
                Me.BillPrice.Col = �ۼ��б�.�ּ�
            ElseIf mint���� = 1 Or mint���� = 2 Then
                Me.BillPrice.Col = �ۼ��б�.�ֳɱ���
            ElseIf mint���� = 3 Then
                Me.BillPrice.Col = �ۼ��б�.��������
            End If
        
            Call GetDrugStore(1, lngMediId)
            
'            If mint���� = 1 Or mint���� = 2 Then
'                If BillStore.Rows = 1 Then
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�Ƿ��п��) = 0
'                ElseIf BillStore.TextMatrix(1, 0) = "" Then
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�Ƿ��п��) = 0
'                Else
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�Ƿ��п��) = 1
'                End If
'            End If
        End With
    Else
        '���������ʾ״̬
        Me.BillPrice.Active = False
        Me.BillStore.Active = False
        Me.BillPay.Active = False
        Me.cmdOk.Visible = False
        Me.cmdCanc.Caption = "����(&C)"
        Me.cmdCanc.Top = Me.cmdOk.Top
        Me.txtSummary.Enabled = False
        optʱ��(1).Value = True
        optʱ��(0).Enabled = False
        optʱ��(1).Enabled = False
        Me.dtpRunDate.Enabled = False
        Me.chk�Զ�����Ӧ����䶯.Enabled = False
        Me.chk������.Enabled = False
        
        Dim strBills As String
        strBills = ""
        
        gstrSql = "select P.ID,M.id as ҩƷid,'['||M.����||']'||M.���� as Ʒ��,M.���,M.����,M.���㵥λ as ��λ,P.ҩ�ⵥλ," & _
            "        P.ԭ��,P.�ּ�,P.������Ŀid,I.���� as ��������," & _
            "        To_Char(P.ִ������,'yyyy-MM-dd hh24:mi:ss') ִ������,P.�䶯ԭ��,P.����˵��,P.������,p.���������,1/(1-p.ָ�������/100)-1 �ӳ���," & _
            " from �շѼ�Ŀ P,�շ���ĿĿ¼ M,������Ŀ I,ҩƷ��� P" & _
            " where P.�շ�ϸĿid=M.id and P.������Ŀid=I.id And M.ID=P.ҩƷID and P.ID=[1] " & _
            GetPriceClassString("P") & _
            " order by P.id"                            '�����IDȡ���Ǽ۸��¼ID����һ��ID
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngBillId)
        
        With rsTemp
            Me.BillPrice.Rows = .RecordCount + 1
            Do While Not .EOF
                dbl��װ = GetModulus(Val(!ҩƷid))
                
                strBills = strBills & "," & !ID
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ҩƷid) = !ҩƷid
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.Ʒ��) = !Ʒ��
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.���) = IIf(IsNull(!���), "", !���)
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.����) = IIf(IsNull(!����), "", !����)
                If intҩ�ⵥλ = 0 Then
                    Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��λ) = IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��λ) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
                End If
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.ԭ��) = FormatEx(!ԭ�� * dbl��װ, mintPriceDigit)
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�ּ�) = FormatEx(!�ּ� * dbl��װ, mintPriceDigit)
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.������ID) = !������Ŀid
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.��������) = !��������
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.���������) = IIf(IsNull(!���������), 0, !���������)
                Me.BillPrice.TextMatrix(.AbsolutePosition, �ۼ��б�.�ӳ���) = !�ӳ���
                
                Me.txtSummary = IIf(IsNull(!����˵��), "", !����˵��)
                Me.txtValuer.Text = IIf(IsNull(!������), "", !������)
                Me.dtpRunDate.Value = !ִ������
                
                If !ִ������ <= strToday And !�䶯ԭ�� = 0 Then        'δ���е��ۼ���,��ִ�м���
                    gstrSql = "zl_ҩƷ�շ���¼_Adjust(" & !ID & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
            
            If !ִ������ > strToday Then
                '���ִ��ʱ��δ������ֻ��ģ����ʾ���䶯
                Me.lblTitle.Caption = "���䶯��(����ִ��ʱ��δ������ӳ�Ŀ����ܲ�׼ȷ)"
            Else
                'ִ��ʱ���ѵ����϶�Ҳ�����˵��ۼ��㣬ֱ�Ӵ��շ���¼��ȡ���۱䶯���
                Me.lblTitle.Caption = "���䶯��"
                gstrSql = "select S.ID,S.ҩƷID,D.���� as �ⷿ,'['||M.����||']'||M.���� as ҩƷ,M.���,M.����,M.���㵥λ as ��λ,P.ҩ�ⵥλ,S.����,S.����,S.ԭ��,S.�ּ�,S.�������" & _
                        " from (select ID,�ⷿID,ҩƷID,����,��д���� as ����,�ɱ��� as ԭ��,���ۼ� as �ּ�,���۽�� as �������" & _
                        "       from (select P.ID,N.�ⷿID,N.ҩƷID,N.����,N.��д����,N.�ɱ���,N.���ۼ�,N.���۽��" & _
                        "            from ҩƷ�շ���¼ N, (select ID,�շ�ϸĿID,ִ������,��ֹ���� from �շѼ�Ŀ where ID=[1]" & _
                        GetPriceClassString("") & ") P" & _
                        "       where N.ҩƷID=P.�շ�ϸĿID and ����=13 and N.����ID is null " & _
                        "             and N.������� Between P.ִ������ and nvl(P.��ֹ����,sysdate))) S," & _
                        "       ���ű� D,�շ���ĿĿ¼ M,ҩƷ��� P" & _
                        " where S.�ⷿid+0=D.id and S.ҩƷID=M.ID And M.ID=P.ҩƷID" & _
                        " order by M.����,S.����"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(strBills, 2)))
                    
                With rsTemp
                    If .RecordCount > 0 Then Me.BillStore.Rows = .RecordCount + 1
                    Do While Not .EOF
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.�ⷿ) = !�ⷿ
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.ҩƷ) = !ҩƷ
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.���) = IIf(IsNull(!���), "", !���)
                        If intҩ�ⵥλ = 0 Then
                            Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.��λ) = IIf(IsNull(!��λ), "", !��λ)
                        Else
                            Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.��λ) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
                        End If
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.����) = IIf(IsNull(!����), "", !����)
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.����) = Format(!���� / dbl��װ, "0.00000")
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.ԭ��) = FormatEx(!ԭ�� * dbl��װ, mintPriceDigit)
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.�ּ�) = FormatEx(!�ּ� * dbl��װ, mintPriceDigit)
                        Me.BillStore.TextMatrix(.AbsolutePosition, ����б�.�������) = Format(!�������, mstrMoneyFormat)
                        .MoveNext
                    Loop
                End With
            
            End If
            
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniGrid()
    With Me.BillPrice
        .Cols = �ۼ��б�.����
        .MsfObj.FixedCols = 0

        If intDrugType = 1 Then   '��ҩ����ҩ
            .TextMatrix(0, �ۼ��б�.ҩƷid) = "ҩƷid"
            .TextMatrix(0, �ۼ��б�.Ʒ��) = "Ʒ��"
            .TextMatrix(0, �ۼ��б�.���) = "���"
            .TextMatrix(0, �ۼ��б�.����) = "����"
            .TextMatrix(0, �ۼ��б�.��λ) = "��λ"
            .TextMatrix(0, �ۼ��б�.����) = "����"
            .TextMatrix(0, �ۼ��б�.�ϴ�����) = "�ϴ�����"
            .TextMatrix(0, �ۼ��б�.ԭ��) = "ԭ���ۼ�"
            .TextMatrix(0, �ۼ��б�.�ּ�) = "�����ۼ�"
            .TextMatrix(0, �ۼ��б�.������ID) = "����id"
            .TextMatrix(0, �ۼ��б�.ԭ����ID) = "ԭ����id"
            .TextMatrix(0, �ۼ��б�.��������) = "������Ŀ"
            .TextMatrix(0, �ۼ��б�.ԭ�ɱ���) = IIf(mint���� = 1 Or mint���� = 2, "ԭ�ɹ���", "�ɱ���")
            .TextMatrix(0, �ۼ��б�.�ֳɱ���) = "�ֲɹ���"
            .TextMatrix(0, �ۼ��б�.ԭ�ɹ��޼�) = IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, "�ɹ��޼�", "ԭ�ɹ��޼�")
            .TextMatrix(0, �ۼ��б�.�ֲɹ��޼�) = "�ֲɹ��޼�"
            .TextMatrix(0, �ۼ��б�.ԭָ���ۼ�) = IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, "ָ���ۼ�", "ԭָ���ۼ�")
            .TextMatrix(0, �ۼ��б�.��ָ���ۼ�) = "��ָ���ۼ�"
            .TextMatrix(0, �ۼ��б�.�Ƿ��п��) = "�Ƿ��п��"
            .TextMatrix(0, �ۼ��б�.����ϵ��) = "����ϵ��"
            .TextMatrix(0, �ۼ��б�.ҩ��ID) = "ҩ��ID"
            .TextMatrix(0, �ۼ��б�.��װϵ��) = "��װϵ��"
            .TextMatrix(0, �ۼ��б�.���������) = "���������"
            .TextMatrix(0, �ۼ��б�.�ӳ���) = "�ӳ���"
            
            .ColWidth(�ۼ��б�.ҩƷid) = 0
            .ColWidth(�ۼ��б�.Ʒ��) = 2600
            .ColWidth(�ۼ��б�.���) = 1200
            .ColWidth(�ۼ��б�.����) = 1000
            .ColWidth(�ۼ��б�.��λ) = 600
            .ColWidth(�ۼ��б�.����) = 0
            .ColWidth(�ۼ��б�.�ϴ�����) = 0
            .ColWidth(�ۼ��б�.ԭ��) = 900
            .ColWidth(�ۼ��б�.�ּ�) = 900
            .ColWidth(�ۼ��б�.������ID) = 0
            .ColWidth(�ۼ��б�.ԭ����ID) = 0
            .ColWidth(�ۼ��б�.��������) = 900
            .ColWidth(�ۼ��б�.ԭ�ɱ���) = 975
            .ColWidth(�ۼ��б�.�ֳɱ���) = IIf(mint���� = 1 Or mint���� = 2, 975, 0)
            .ColWidth(�ۼ��б�.ԭ�ɹ��޼�) = 0
            .ColWidth(�ۼ��б�.�ֲɹ��޼�) = 0 'IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 0, 1000)
            .ColWidth(�ۼ��б�.ԭָ���ۼ�) = 0
            .ColWidth(�ۼ��б�.��ָ���ۼ�) = 0 'IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 0, 1000)
            .ColWidth(�ۼ��б�.�Ƿ��п��) = 0
            .ColWidth(�ۼ��б�.����ϵ��) = 0
            .ColWidth(�ۼ��б�.ҩ��ID) = 0
            .ColWidth(�ۼ��б�.��װϵ��) = 0
            .ColWidth(�ۼ��б�.���������) = 0
            .ColWidth(�ۼ��б�.�ӳ���) = 0
            
        Else    '�в�ҩ
            .TextMatrix(0, �ۼ��б�.ҩƷid) = "ҩƷid"
            .TextMatrix(0, �ۼ��б�.Ʒ��) = "Ʒ��"
            .TextMatrix(0, �ۼ��б�.���) = "���"
            .TextMatrix(0, �ۼ��б�.����) = "����"
            .TextMatrix(0, �ۼ��б�.��λ) = "��λ"
            .TextMatrix(0, �ۼ��б�.����) = "����"
            .TextMatrix(0, �ۼ��б�.�ϴ�����) = "�ϴ�����"
            .TextMatrix(0, �ۼ��б�.ԭ��) = "ԭ���ۼ�"
            .TextMatrix(0, �ۼ��б�.�ּ�) = "�����ۼ�"
            .TextMatrix(0, �ۼ��б�.������ID) = "����id"
            .TextMatrix(0, �ۼ��б�.ԭ����ID) = "ԭ����id"
            .TextMatrix(0, �ۼ��б�.��������) = "������Ŀ"
            .TextMatrix(0, �ۼ��б�.ԭ�ɱ���) = IIf(mint���� = 1 Or mint���� = 2, "ԭ�ɹ���", "�ɱ���")
            .TextMatrix(0, �ۼ��б�.�ֳɱ���) = "�ֲɹ���"
            .TextMatrix(0, �ۼ��б�.ԭ�ɹ��޼�) = "ԭ�ɹ��޼�"
            .TextMatrix(0, �ۼ��б�.�ֲɹ��޼�) = "�ֲɹ��޼�"
            .TextMatrix(0, �ۼ��б�.ԭָ���ۼ�) = "ԭָ���ۼ�"
            .TextMatrix(0, �ۼ��б�.��ָ���ۼ�) = "��ָ���ۼ�"
            .TextMatrix(0, �ۼ��б�.�Ƿ��п��) = "�Ƿ��п��"
            .TextMatrix(0, �ۼ��б�.����ϵ��) = "����ϵ��"
            .TextMatrix(0, �ۼ��б�.ҩ��ID) = "ҩ��ID"
            .TextMatrix(0, �ۼ��б�.��װϵ��) = "��װϵ��"
            .TextMatrix(0, �ۼ��б�.���������) = "���������"
            .TextMatrix(0, �ۼ��б�.�ӳ���) = "�ӳ���"
                        
            .ColWidth(�ۼ��б�.ҩƷid) = 0
            .ColWidth(�ۼ��б�.Ʒ��) = 2800
            .ColWidth(�ۼ��б�.���) = 1200
            .ColWidth(�ۼ��б�.����) = 1000
            .ColWidth(�ۼ��б�.��λ) = 600
            .ColWidth(�ۼ��б�.����) = 0
            .ColWidth(�ۼ��б�.�ϴ�����) = 0
            .ColWidth(�ۼ��б�.ԭ��) = 1200
            .ColWidth(�ۼ��б�.�ּ�) = 1200
            .ColWidth(�ۼ��б�.������ID) = 0
            .ColWidth(�ۼ��б�.ԭ����ID) = 0
            .ColWidth(�ۼ��б�.��������) = 1200
            .ColWidth(�ۼ��б�.ԭ�ɱ���) = 975
            .ColWidth(�ۼ��б�.�ֳɱ���) = IIf(mint���� = 1 Or mint���� = 2, 975, 0)
            .ColWidth(�ۼ��б�.ԭ�ɹ��޼�) = 0
            .ColWidth(�ۼ��б�.�ֲɹ��޼�) = 0 'IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 0, 1000)
            .ColWidth(�ۼ��б�.ԭָ���ۼ�) = 0
            .ColWidth(�ۼ��б�.��ָ���ۼ�) = 0 'IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 0, 1000)
            .ColWidth(�ۼ��б�.�Ƿ��п��) = 0
            .ColWidth(�ۼ��б�.����ϵ��) = 0
            .ColWidth(�ۼ��б�.ҩ��ID) = 0
            .ColWidth(�ۼ��б�.��װϵ��) = 0
            .ColWidth(�ۼ��б�.���������) = 0
            .ColWidth(�ۼ��б�.�ӳ���) = 0
        End If
        
        If lngBillId <> 0 Then
            .ColWidth(�ۼ��б�.ԭ�ɱ���) = 0
            .ColWidth(�ۼ��б�.�ֳɱ���) = 0
        End If
        
        .ColData(�ۼ��б�.ҩƷid) = 5
        .ColData(�ۼ��б�.Ʒ��) = 1
        .ColData(�ۼ��б�.���) = 5
        .ColData(�ۼ��б�.����) = 5
        .ColData(�ۼ��б�.��λ) = 5
        .ColData(�ۼ��б�.����) = 5
        .ColData(�ۼ��б�.�ϴ�����) = 5
        .ColData(�ۼ��б�.ԭ��) = 5
        .ColData(�ۼ��б�.�ּ�) = IIf(mint���� = 3, 5, IIf(mint���� = 1, 0, 4))
        .ColData(�ۼ��б�.������ID) = 5
        .ColData(�ۼ��б�.ԭ����ID) = 5
        .ColData(�ۼ��б�.��������) = 1
        .ColData(�ۼ��б�.ԭ�ɱ���) = 5
        .ColData(�ۼ��б�.�ֳɱ���) = IIf(mint���� = 1 Or mint���� = 2, 4, 0)
        .ColData(�ۼ��б�.ԭ�ɹ��޼�) = 5
        .ColData(�ۼ��б�.�ֲɹ��޼�) = IIf(mint���� = 3, 5, IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 5, 4))
        .ColData(�ۼ��б�.ԭָ���ۼ�) = 5
        .ColData(�ۼ��б�.��ָ���ۼ�) = IIf(mint���� = 3, 5, IIf(InStr(1, mstrPrivs, "ָ���۸����") = 0, 5, 4))

        .ColAlignment(�ۼ��б�.ҩƷid) = 1
        .ColAlignment(�ۼ��б�.Ʒ��) = 1
        .ColAlignment(�ۼ��б�.���) = 1
        .ColAlignment(�ۼ��б�.����) = 1
        .ColAlignment(�ۼ��б�.��λ) = 4
        .ColAlignment(�ۼ��б�.����) = 1
        .ColAlignment(�ۼ��б�.�ϴ�����) = 1
        .ColAlignment(�ۼ��б�.ԭ��) = 7
        .ColAlignment(�ۼ��б�.�ּ�) = 7
        .ColAlignment(�ۼ��б�.������ID) = 1
        .ColAlignment(�ۼ��б�.ԭ����ID) = 1
        .ColAlignment(�ۼ��б�.��������) = 1
        .ColAlignment(�ۼ��б�.ԭ�ɱ���) = 7
        .ColAlignment(�ۼ��б�.�ֳɱ���) = 7
        .ColAlignment(�ۼ��б�.ԭ�ɹ��޼�) = 7
        .ColAlignment(�ۼ��б�.�ֲɹ��޼�) = 7
        .ColAlignment(�ۼ��б�.ԭָ���ۼ�) = 7
        .ColAlignment(�ۼ��б�.��ָ���ۼ�) = 7
        
        .PrimaryCol = �ۼ��б�.Ʒ��
        .LocateCol = �ۼ��б�.Ʒ��
    End With
    
    With Me.BillStore
        .Rows = 2
        .MsfObj.FixedCols = 0
        .Cols = ����б�.����
        .TextMatrix(0, ����б�.�ⷿ) = "�ⷿ"
        .TextMatrix(0, ����б�.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, ����б�.ҩƷ) = "ҩƷ"
        .TextMatrix(0, ����б�.���) = "���"
        .TextMatrix(0, ����б�.��λ) = "��λ"
        .TextMatrix(0, ����б�.����) = "����"
        .TextMatrix(0, ����б�.Ч��) = "Ч��"
        .TextMatrix(0, ����б�.����) = "����"
        .TextMatrix(0, ����б�.����) = "����"
        .TextMatrix(0, ����б�.ԭ��) = "ԭ���ۼ�"
        .TextMatrix(0, ����б�.�ּ�) = "�����ۼ�"
        .TextMatrix(0, ����б�.�������) = "�������"
        .TextMatrix(0, ����б�.�ӳ���) = "�ӳ���(%)"
        .TextMatrix(0, ����б�.ԭ�ɱ���) = "ԭ�ɹ���"
        .TextMatrix(0, ����б�.�ֳɱ���) = "�ֲɹ���"
        .TextMatrix(0, ����б�.��۲�) = "��۲�"
        .TextMatrix(0, ����б�.����) = "����"
        .TextMatrix(0, ����б�.���) = "���"
        .TextMatrix(0, ����б�.ҩƷid) = "ҩƷID"
        .TextMatrix(0, ����б�.�ⷿid) = "�ⷿID"
        .TextMatrix(0, ����б�.��Ӧ��ID) = "��Ӧ��ID"
        
        .ColData(����б�.�ⷿ) = 5
        .ColData(����б�.��Ӧ��) = 5
        .ColData(����б�.ҩƷ) = 5
        .ColData(����б�.���) = 5
        .ColData(����б�.��λ) = 5
        .ColData(����б�.����) = 5
        .ColData(����б�.Ч��) = 5
        .ColData(����б�.����) = 5
        .ColData(����б�.����) = 5
        .ColData(����б�.ԭ��) = 5
        .ColData(����б�.�ּ�) = 0
        .ColData(����б�.�������) = 5
        .ColData(����б�.�ӳ���) = 4
        .ColData(����б�.ԭ�ɱ���) = 5
        .ColData(����б�.�ֳɱ���) = 4
        .ColData(����б�.��۲�) = 5
        .ColData(����б�.����) = 5
        .ColData(����б�.���) = 5
        .ColData(����б�.ҩƷid) = 5
        .ColData(����б�.�ⷿid) = 5
        .ColData(����б�.��Ӧ��ID) = 5
        
        .ColWidth(����б�.�ⷿ) = 1000
        .ColWidth(����б�.��Ӧ��) = 1500
        .ColWidth(����б�.ҩƷ) = 2800
        .ColWidth(����б�.���) = 1350
        .ColWidth(����б�.��λ) = 600
        .ColWidth(����б�.����) = 800
        .ColWidth(����б�.Ч��) = 1000
        .ColWidth(����б�.����) = 1000
        .ColWidth(����б�.����) = 1000
        .ColWidth(����б�.ԭ��) = 900
        .ColWidth(����б�.�ּ�) = 900
        .ColWidth(����б�.�������) = 1050
        .ColWidth(����б�.����) = 0
        .ColWidth(����б�.���) = 0
        .ColWidth(����б�.ҩƷid) = 0
        .ColWidth(����б�.�ⷿid) = 0
        .ColWidth(����б�.��Ӧ��ID) = 0
        
        If mint���� = 0 Then
            .ColWidth(����б�.�ӳ���) = 0
            .ColWidth(����б�.ԭ�ɱ���) = 0
            .ColWidth(����б�.�ֳɱ���) = 0
            .ColWidth(����б�.��۲�) = 0
            .ColWidth(����б�.ԭ��) = 900
            .ColWidth(����б�.�ּ�) = 900
        ElseIf mint���� = 1 Then
            .ColWidth(����б�.�ӳ���) = 900
            .ColWidth(����б�.ԭ�ɱ���) = 900
            .ColWidth(����б�.�ֳɱ���) = 900
            .ColWidth(����б�.��۲�) = 900
            .ColWidth(����б�.ԭ��) = 0
            .ColWidth(����б�.�ּ�) = 0
        Else
            .ColWidth(����б�.�ӳ���) = 900
            .ColWidth(����б�.ԭ�ɱ���) = 900
            .ColWidth(����б�.�ֳɱ���) = 900
            .ColWidth(����б�.��۲�) = 900
            .ColWidth(����б�.ԭ��) = 900
            .ColWidth(����б�.�ּ�) = 900
        End If
        
        .ColAlignment(����б�.�ⷿ) = 1
        .ColAlignment(����б�.��Ӧ��) = 1
        .ColAlignment(����б�.ҩƷ) = 1
        .ColAlignment(����б�.���) = 1
        .ColAlignment(����б�.��λ) = 4
        .ColAlignment(����б�.����) = 1
        .ColAlignment(����б�.Ч��) = 1
        .ColAlignment(����б�.����) = 1
        .ColAlignment(����б�.����) = 7
        .ColAlignment(����б�.ԭ��) = 7
        .ColAlignment(����б�.�ּ�) = 7
        .ColAlignment(����б�.�������) = 7
        .ColAlignment(����б�.�ӳ���) = 7
        .ColAlignment(����б�.ԭ�ɱ���) = 7
        .ColAlignment(����б�.�ֳɱ���) = 7
        .ColAlignment(����б�.��۲�) = 7
        .ColAlignment(����б�.����) = 7
        .ColAlignment(����б�.���) = 7
        .ColAlignment(����б�.ҩƷid) = 7
        
        .PrimaryCol = ����б�.�ⷿ
        .LocateCol = ����б�.�ⷿ
    End With
    
    
    With BillPay
        .Rows = 2
        .Cols = Ӧ������.����
        .MsfObj.FixedCols = 0
        
        .TextMatrix(0, Ӧ������.ҩƷid) = "ҩƷid"
        .TextMatrix(0, Ӧ������.Ʒ��) = "Ʒ��"
        .TextMatrix(0, Ӧ������.��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, Ӧ������.��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, Ӧ������.��Ʊ���) = "��Ʊ���"
        
        .ColWidth(Ӧ������.ҩƷid) = 0
        .ColWidth(Ӧ������.Ʒ��) = 3000
        .ColWidth(Ӧ������.��Ʊ��) = 1000
        .ColWidth(Ӧ������.��Ʊ����) = 2000
        .ColWidth(Ӧ������.��Ʊ���) = 1000
        
        .ColData(Ӧ������.ҩƷid) = 5
        .ColData(Ӧ������.Ʒ��) = 5
        .ColData(Ӧ������.��Ʊ��) = 4
        .ColData(Ӧ������.��Ʊ����) = 2
        .ColData(Ӧ������.��Ʊ���) = 4

        .ColAlignment(Ӧ������.ҩƷid) = 1
        .ColAlignment(Ӧ������.Ʒ��) = 1
        .ColAlignment(Ӧ������.��Ʊ��) = 1
        .ColAlignment(Ӧ������.��Ʊ����) = 4
        .ColAlignment(Ӧ������.��Ʊ���) = 7
        
'        .PrimaryCol = Ӧ������.Ʒ��
'        .LocateCol = Ӧ������.Ʒ��
    End With
End Sub

    
Private Sub BillPay_EnterCell(Row As Long, Col As Long)
    With BillPay
        Select Case Col
            Case Ӧ������.��Ʊ��
                .TxtCheck = False
                .MaxLength = 20
            Case Ӧ������.��Ʊ���
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
            Case Ӧ������.��Ʊ����
                .TxtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
        End Select
   End With
End Sub


Private Sub BillPay_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> 13 Then Exit Sub
    
    With BillPay
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case Ӧ������.��Ʊ���
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬷�Ʊ������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0.001 Then
                        MsgBox "�Բ��𣬷�Ʊ���������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "��Ʊ������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = FormatEx(strKey, 2)
                    .Text = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                    
                End If
            Case Ӧ������.��Ʊ����
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "�Բ���Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ��������(2000-10-10) �� ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

'ת����ֵΪ����
Public Function TranNumToDate(ByVal strNum As String) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim StrDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    TranNumToDate = StrDate
End Function

Private Sub BillPrice_AfterAddRow(Row As Long)
    Call SetColor
End Sub

Private Sub BillPrice_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Call GetDrugStore(Row, 0, Val(BillPrice.TextMatrix(Row, �ۼ��б�.ҩƷid)))
End Sub

Private Sub BillPrice_CommandClick()
    Dim strSqlType As String
    Dim dbl��װ As Double
    
    On Error GoTo errHandle
    If intDrugType = 1 Then
        If InStr(1, mstrPrivs, "��������ҩ") > 0 And InStr(1, mstrPrivs, "�����г�ҩ") > 0 Then
            strSqlType = "In('5','6')"
        ElseIf InStr(1, mstrPrivs, "��������ҩ") > 0 Then
            strSqlType = "='5'"
        ElseIf InStr(1, mstrPrivs, "�����г�ҩ") > 0 Then
            strSqlType = "='6'"
        End If
    Else
        strSqlType = "='7'"
    End If
    
    Select Case Me.BillPrice.Col
    Case �ۼ��б�.Ʒ��
        gstrSql = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,P.ҩ�ⵥλ,decode(I.�Ƿ���,1,'ʱ��','����') ����,Nvl(P.�ɱ���,0) �ɱ���,P.ָ��������,P.ָ�����ۼ�,P.����ϵ��,P.ҩ��ID " & _
                 " from �շ���ĿĿ¼ I,ҩƷ��� P" & _
                 " where I.��� " & strSqlType & " And I.ID=P.ҩƷID" & _
                 "       and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "δ����ҩƷ��", vbInformation, gstrSysName: Exit Sub
            End If
            
            Me.lvwItem.Tag = �ۼ��б�.Ʒ��

            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "����", "����", 900
                .Add , "����", "����", 2000
                .Add , "���", "���", 1200
                .Add , "����", "����", 1200
                .Add , "��λ", "��λ", 500
                .Add , "����", "����", 600
                .Add , "�ɱ���", "�ɱ���", 600
                .Add , "�ɹ��޼�", "�ɹ��޼�", 0
                .Add , "ָ���ۼ�", "ָ���ۼ�", 0
                .Add , "����ϵ��", "����ϵ��", 0
                .Add , "ҩ��ID", "ҩ��ID", 0
            End With
            Me.lvwItem.Width = 6500
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                dbl��װ = GetModulus(Val(!ID))
                
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
                objItem.SubItems(Me.lvwItem.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                If intҩ�ⵥλ = 0 Then
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                Else
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
                End If
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɱ���").Index - 1) = FormatEx(Val(IIf(IsNull(!�ɱ���), "", !�ɱ���)) * dbl��װ, mintCostDigit)
                
                objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɹ��޼�").Index - 1) = FormatEx(Val(IIf(IsNull(!ָ��������), "", !ָ��������)) * dbl��װ, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("ָ���ۼ�").Index - 1) = FormatEx(Val(IIf(IsNull(!ָ�����ۼ�), "", !ָ�����ۼ�)) * dbl��װ, mintPriceDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����ϵ��").Index - 1) = !����ϵ��
                objItem.SubItems(Me.lvwItem.ColumnHeaders("ҩ��ID").Index - 1) = !ҩ��ID
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.BillPrice.Left
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    Case �ۼ��б�.��������
        
        gstrSql = "select id,����,����" & _
                " from ������Ŀ" & _
                " where (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) and ĩ��=1"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "BillPrice_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "û�����ú�������Ŀ", vbExclamation, gstrSysName: Exit Sub
            End If
            
            Me.lvwItem.Tag = �ۼ��б�.��������
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "����", "����", 600
                .Add , "����", "����", 1000
            End With
            Me.lvwItem.Width = 1800
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Exit Sub
            End If
        End With
        
        With Me.lvwItem
            .Left = BillPrice.Left + BillPrice.MsfObj.CellLeft
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDrugRepeat(ByVal lngҩƷID As Long) As Boolean
    Dim n As Integer
    
    With BillPrice
        For n = 1 To .Rows - 1
            If .TextMatrix(n, �ۼ��б�.ҩƷid) <> "" Then
                If Val(.TextMatrix(n, �ۼ��б�.ҩƷid)) = lngҩƷID Then
                    MsgBox "�Բ������и�ҩƷ�������ظ����룡", vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckDrugRepeat = True
End Function

Private Sub BillPrice_EditKeyPress(KeyAscii As Integer)
    With BillPrice
        If .Col = �ۼ��б�.�ֳɱ��� Then
            mdbl�ɱ��� = Val(.TextMatrix(.Row, .Col))
        End If
    End With
End Sub

Private Sub BillPrice_EnterCell(Row As Long, Col As Long)
    Dim n As Integer
    
    Select Case Col
    Case �ۼ��б�.�ֲɹ��޼�, �ۼ��б�.��ָ���ۼ�
'        BillPrice.TxtCheck = True
        BillPrice.MaxLength = 11
        BillPrice.TextMask = ".1234567890"
    Case �ۼ��б�.Ʒ��
        Me.lblHelp.Caption = "��ʾ������ҩƷ���롢����ѡ�����ҩƷ"
    Case �ۼ��б�.�ּ�
        Me.lblHelp.Caption = "��ʾ��F3����ҩ�۸������㣬���ݳɱ��ۼ�������µ��ۼ�"
        
        If mint���� <> 3 Then
            If mint���� = 1 Or (Me.BillPrice.TextMatrix(Row, �ۼ��б�.����) = "ʱ��" And mblnʱ��ҩƷ����) Then
                Me.BillPrice.ColData(�ۼ��б�.�ּ�) = 0
            Else
                Me.BillPrice.ColData(�ۼ��б�.�ּ�) = 4
                BillPrice.MaxLength = 11
                BillPrice.TextMask = ".1234567890"
            End If
        End If
    Case �ۼ��б�.�ֳɱ���
        Me.BillPrice.ColData(�ۼ��б�.�ֳɱ���) = 0
        If mint���� = 1 Or mint���� = 2 Then
            Me.BillPrice.ColData(�ۼ��б�.�ֳɱ���) = 4
            BillPrice.MaxLength = 11
            BillPrice.TextMask = ".1234567890"
        End If
    Case �ۼ��б�.��������
        Me.BillPrice.TextMatrix(Row, �ۼ��б�.�ּ�) = FormatEx(Me.BillPrice.TextMatrix(Row, �ۼ��б�.�ּ�), mintPriceDigit)
        Me.lblHelp.Caption = "��ʾ����ȷ����ҩƷ��������Ŀ���Ա���Ч��ɲ����Ŀ����"

    Case Else
        Me.lblHelp.Caption = ""
    End Select
    
    If BillStore.Rows > 1 Then
        If Trim(BillStore.TextMatrix(1, 0)) <> "" Then
            For n = 1 To BillStore.Rows - 1
                If Val(Me.BillPrice.TextMatrix(Row, �ۼ��б�.ҩƷid)) = Val(BillStore.TextMatrix(n, ����б�.ҩƷid)) Then
                    BillStore.MsfObj.TopRow = n
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub BillPrice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    Dim strSqlType As String
    Dim lngҩƷID As Long
    Dim dbl��װ As Double
    Dim dblSalePrice As Double
    
    If KeyCode = 13 And Not BillPrice.Active Then
        Cancel = True: Call OS.PressKey(vbKeyTab): Exit Sub
    End If
    
    If KeyCode <> 13 Then Exit Sub
    
    On Error GoTo errHandle
    If intDrugType = 1 Then
        If InStr(1, mstrPrivs, "��������ҩ") > 0 And InStr(1, mstrPrivs, "�����г�ҩ") > 0 Then
            strSqlType = "In('5','6')"
        ElseIf InStr(1, mstrPrivs, "��������ҩ") > 0 Then
            strSqlType = "='5'"
        ElseIf InStr(1, mstrPrivs, "�����г�ҩ") > 0 Then
            strSqlType = "='6'"
        End If
    Else
        strSqlType = "='7'"
    End If
    
    Select Case Me.BillPrice.Col
    Case �ۼ��б�.Ʒ��
        If Trim(Me.BillPrice.Text) = "" Then Exit Sub
        If Me.BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.Ʒ��) = UCase(Trim(Me.BillPrice.Text)) Then Exit Sub
        strInput = UCase(Trim(Me.BillPrice.Text))
        
        gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,P.ҩ�ⵥλ,decode(I.�Ƿ���,1,'ʱ��','����') ����,Nvl(P.�ɱ���,0) �ɱ���,P.ָ��������,P.ָ�����ۼ�,P.����ϵ��,P.ҩ��ID " & _
                 " from �շ���ĿĿ¼ I,�շ���Ŀ���� N,ҩƷ��� P" & _
                 " where I.ID=N.�շ�ϸĿID and I.��� " & strSqlType & " And I.ID=P.ҩƷID " & _
                 "       and (I.���� like [1] " & _
                 "            or N.���� Like [2] " & _
                 "            or N.���� Like [2])" & _
                 "       and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strInput & "%", gstrMatch & strInput & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "δ�ҵ����ҩƷ�����������룡", vbInformation, gstrSysName
'                Cancel = True
                Exit Sub
            End If
            
            Me.lvwItem.Tag = �ۼ��б�.Ʒ��
            Me.lvwItem.Tag = �ۼ��б�.Ʒ��
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "����", "����", 900
                .Add , "����", "����", 2000
                .Add , "���", "���", 1200
                .Add , "����", "����", 1200
                .Add , "��λ", "��λ", 500
                .Add , "����", "����", 600
                .Add , "�ɱ���", "�ɱ���", 800
                .Add , "�ɹ��޼�", "�ɹ��޼�", 0
                .Add , "ָ���ۼ�", "ָ���ۼ�", 0
                .Add , "����ϵ��", "����ϵ��", 0
                .Add , "ҩ��ID", "ҩ��ID", 0
            End With
            Me.lvwItem.Width = 6500
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                dbl��װ = GetModulus(Val(!ID))
                
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
                objItem.SubItems(Me.lvwItem.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                If intҩ�ⵥλ = 0 Then
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                Else
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
                End If
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɱ���").Index - 1) = FormatEx(Val(IIf(IsNull(!�ɱ���), "", !�ɱ���)) * dbl��װ, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɹ��޼�").Index - 1) = FormatEx(Val(IIf(IsNull(!ָ��������), "", !ָ��������)) * dbl��װ, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("ָ���ۼ�").Index - 1) = FormatEx(Val(IIf(IsNull(!ָ�����ۼ�), "", !ָ�����ۼ�)) * dbl��װ, mintPriceDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����ϵ��").Index - 1) = !����ϵ��
                objItem.SubItems(Me.lvwItem.ColumnHeaders("ҩ��ID").Index - 1) = !ҩ��ID
                
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Cancel = True: Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.BillPrice.Left
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        Cancel = True
    
    Case �ۼ��б�.��������
        If Trim(Me.BillPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.BillPrice.Text)
        
        gstrSql = "select id,����,����" & _
                " from ������Ŀ" & _
                " where (���� like [1] or ���� like [2] or ���� like [2])" & _
                "       and (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) and ĩ��=1"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strInput & "%", gstrMatch & strInput & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "����Ŀ������", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            
            Me.lvwItem.Tag = �ۼ��б�.��������
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "����", "����", 600
                .Add , "����", "����", 1000
            End With
            Me.lvwItem.Width = 1800
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Cancel = True: Exit Sub
            End If
        End With
        
        With Me.lvwItem
            .Left = BillPrice.Left + BillPrice.MsfObj.CellLeft
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        Cancel = True
    
    Case �ۼ��б�.�ּ�
        With BillPrice
            If .Text = "" Then Exit Sub
            
            lngҩƷID = Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid))
            If lngҩƷID = 0 Then Exit Sub
            
            '�ּ۴���ָ���ۼ�ʱ����ʾ�Ƿ����
            If mbln�޼���ʾ = True Then
                If BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.����) = "����" And Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.��ָ���ۼ�)) Then
                    MsgBox "�ּ۸���ָ�����ۼ�" & Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.��ָ���ۼ�)) & "��ָ���۸񽫺��ۼ�һ�£�", vbInformation, gstrSysName
                End If
            End If
            
            If Val(.Text) < 0 Then
                MsgBox "�ۼ۲���Ϊ������", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            .TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�) = .Text
            If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.��ָ���ۼ�)) Then
                BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.��ָ���ۼ�) = .Text
            End If
            
            '�����������
            Call ChangeDrugStore(BillPrice.Row, lngҩƷID, Val(.Text))
            
            '�в�ҩ��������������ּ�
            Call BatchAdjustPriceByItem(BillPrice.Row, Val(.Text))
            
            blnModify = True
        End With
        
    Case �ۼ��б�.�ֲɹ��޼�
        With BillPrice
            If Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid)) = 0 Then Exit Sub
            
            If .Text = "" Then Exit Sub
            
            If Val(.Text) < 0 Then
                MsgBox "�۸���Ϊ������", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            If mbln�޼���ʾ = True Then
                If Val(.Text) < Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֳɱ���)) Then
                    If MsgBox("��ָ���ɹ��޼۵����ּ�" & Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֳɱ���)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
            
            .TextMatrix(BillPrice.Row, �ۼ��б�.�ֲɹ��޼�) = .Text
            
            blnModify = True
        End With
    Case �ۼ��б�.��ָ���ۼ�
        With BillPrice
            If Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid)) = 0 Then Exit Sub
            
            If .Text = "" Then Exit Sub
            
            '��ָ���ۼ�С��ָ���ۼ�ʱ����ʾ�Ƿ����
            If mbln�޼���ʾ = True Then
                If BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.����) = "����" And Val(.Text) < Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�)) Then
                    If MsgBox("��ָ�����ۼ۵����ּ�" & Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
            
            If Val(.Text) < 0 Then
                MsgBox "�۸���Ϊ������", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            .TextMatrix(BillPrice.Row, �ۼ��б�.��ָ���ۼ�) = .Text
            
            blnModify = True
        End With
    Case �ۼ��б�.�ֳɱ���
        With BillPrice
            If .Text = "" Then Exit Sub
            
            If Val(.Text) < 0 Then
                MsgBox "�۸���Ϊ������", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            If mbln�޼���ʾ = True Then
                If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֲɹ��޼�)) Then
                    MsgBox "�ֳɱ��۸���ָ���ɹ��޼�" & Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֲɹ��޼�)) & "���ɹ��޼۽��Ͳɹ���һ�£�", vbInformation, gstrSysName
                End If
            End If
            
            .TextMatrix(BillPrice.Row, �ۼ��б�.�ֳɱ���) = .Text
            If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֲɹ��޼�)) Then
                BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ֲɹ��޼�) = .Text
            End If
            
            If cbo�ۼۼ��㷽ʽ = "�ۼ۰��ֶμӳɼ���" And .TextMatrix(.Row, �ۼ��б�.����) = "ʱ��" And mint���� = 2 Then
                Call get�ֶμӳ��ۼ�(Val(.TextMatrix(BillPrice.Row, �ۼ��б�.�ֳɱ���)), dblSalePrice)
                If dblSalePrice = 0 Then
                    .Text = mdbl�ɱ���
                    .TextMatrix(BillPrice.Row, �ۼ��б�.�ֳɱ���) = .Text
                    .TxtSetFocus
                    Exit Sub
                End If
                dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, �ۼ��б�.���������)) / 100)
'                If dblSalePrice > Val(.TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�)) Then dblSalePrice = Val(.TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�))
                .TextMatrix(.Row, �ۼ��б�.�ּ�) = FormatEx(dblSalePrice, mintPriceDigit)
            ElseIf cbo�ۼۼ��㷽ʽ = "�ۼ۰��̶���������" And .TextMatrix(.Row, �ۼ��б�.����) = "ʱ��" And mint���� = 2 Then
                dblSalePrice = Val(.Text) * (1 + Val(.TextMatrix(.Row, �ۼ��б�.�ӳ���)))
                If dblSalePrice > Val(.TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�)) Then dblSalePrice = Val(.TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�))
                .TextMatrix(.Row, �ۼ��б�.�ּ�) = FormatEx(dblSalePrice, mintPriceDigit)
            End If
            
            CaculateCost Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid)), Val(.Text)
            
            
            '�в�ҩ��������������ּ�
            Call BatchAdjustCostByItem(BillPrice.Row, Val(.Text))
            
            blnModify = True
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrice_LostFocus()
    Me.lblHelp.Caption = ""
End Sub

Private Sub BillStore_EnterCell(Row As Long, Col As Long)
    Dim i As Integer
    
    With BillStore
        If Row = 0 Then Exit Sub
        If .TextMatrix(Row, 0) = "" Or .TextMatrix(Row, ����б�.ҩƷ) = "" Then Exit Sub
        If mint���� = 3 Then
            .ColData(����б�.�ּ�) = 0
            .ColData(����б�.�ֳɱ���) = 0
            .ColData(����б�.�ӳ���) = 0
            Exit Sub
        End If
        Select Case Col
            Case ����б�.�ּ�
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                
                If Val(.TextMatrix(Row, ����б�.���)) = 1 And mblnʱ��ҩƷ���� And mint���� <> 1 Then
                    .ColData(����б�.�ּ�) = 4
                Else
                    .ColData(����б�.�ּ�) = 0
                End If
                
                If BillPrice.Rows = 1 Then Exit Sub
                If BillPrice.TextMatrix(1, �ۼ��б�.ҩƷid) = "" Then Exit Sub
                
                For i = 1 To BillPrice.Rows - 1
                    If Val(BillPrice.TextMatrix(i, �ۼ��б�.ҩƷid)) = Val(.TextMatrix(Row, ����б�.ҩƷid)) Then
                        BillPrice.Row = i
                        Exit For
                    End If
                Next
                
            Case ����б�.�ֳɱ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case ����б�.�ӳ���
                .TxtCheck = True
                .MaxLength = 8
                .TextMask = ".1234567890"
        End Select
    End With
End Sub


Private Sub BillStore_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl��Ʊ��� As Double
    Dim dbl���� As Double
    Dim dbl��� As Double
    Dim dbl�ֳɱ��� As Double
    
    If KeyCode <> 13 Then Exit Sub
    
    With BillStore
        If .Text = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case ����б�.�ּ�
                If Not IsNumeric(.Text) Then
                    MsgBox "�������µ��ۼۡ�", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .Text = FormatEx(.Text, mintPriceDigit)
                End If
                .TextMatrix(intRow, ����б�.�������) = Format(Val(.TextMatrix(intRow, ����б�.����)) * (Val(.Text) - Val(.TextMatrix(intRow, ����б�.ԭ��))), mstrMoneyFormat)
                .TextMatrix(intRow, ����б�.�ּ�) = FormatEx(Val(.Text), mintPriceDigit)
                .TextMatrix(intRow, ����б�.�ֳɱ���) = FormatEx(Val(.TextMatrix(intRow, ����б�.�ּ�)) / (1 + Val(.TextMatrix(intRow, ����б�.�ӳ���)) / 100), mintCostDigit)
                .TextMatrix(intRow, ����б�.��۲�) = Format((Val(.TextMatrix(intRow, ����б�.�ֳɱ���)) - Val(.TextMatrix(intRow, ����б�.ԭ�ɱ���))) * Val(.TextMatrix(intRow, ����б�.����)), mstrMoneyFormat)
                
                For n = 1 To .Rows - 1
                    If BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid) = .TextMatrix(n, ����б�.ҩƷid) Then
                        If Val(.TextMatrix(intRow, ����б�.����)) <> 0 And Val(.TextMatrix(intRow, ����б�.����)) = Val(.TextMatrix(n, ����б�.����)) Then
                            .TextMatrix(n, ����б�.�ּ�) = .TextMatrix(intRow, ����б�.�ּ�)
                            .TextMatrix(n, ����б�.�������) = Format(Val(.TextMatrix(n, ����б�.����)) * (Val(.Text) - Val(.TextMatrix(n, ����б�.ԭ��))), mstrMoneyFormat)
                            .TextMatrix(n, ����б�.�ֳɱ���) = FormatEx(Val(.TextMatrix(n, ����б�.�ּ�)) / (1 + Val(.TextMatrix(n, ����б�.�ӳ���)) / 100), mintCostDigit)
                            .TextMatrix(n, ����б�.��۲�) = Format((Val(.TextMatrix(n, ����б�.�ֳɱ���)) - Val(.TextMatrix(n, ����б�.ԭ�ɱ���))) * Val(.TextMatrix(n, ����б�.����)), mstrMoneyFormat)
                        End If
                        dbl���� = dbl���� + .TextMatrix(n, ����б�.����)
                        dbl��� = dbl��� + .TextMatrix(n, ����б�.����) * Val(.TextMatrix(n, ����б�.�ּ�))
                    End If
                Next
                
                BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�) = FormatEx(dbl��� / dbl����, mintPriceDigit)
                
                If mint���� > 0 Then
                    For n = 1 To .Rows - 1
                        If .TextMatrix(n, ����б�.ҩƷid) <> "" Then
                            If Val(.TextMatrix(n, ����б�.ҩƷid)) = Val(.TextMatrix(intRow, ����б�.ҩƷid)) Then
                                dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(n, ����б�.�ֳɱ���)) - Val(.TextMatrix(n, ����б�.ԭ�ɱ���))) * Val(.TextMatrix(n, ����б�.����))
                            End If
                        End If
                    Next
    
                    If chk�Զ�����Ӧ����䶯.Value = 1 Then
                        For n = 1 To BillPay.Rows - 1
                            If BillPay.TextMatrix(1, 0) <> "" Then
                                If Val(BillPay.TextMatrix(n, Ӧ������.ҩƷid)) = Val(BillStore.TextMatrix(intRow, ����б�.ҩƷid)) Then
                                    BillPay.TextMatrix(n, Ӧ������.��Ʊ���) = FormatEx(dbl��Ʊ���, 2)
                                End If
                            End If
                        Next
                    End If
                End If
            Case ����б�.�ӳ���
                If Val(.Text) < 0 Then Exit Sub
                
                .TextMatrix(intRow, ����б�.�ӳ���) = FormatEx(Val(.Text), 5)
                .TextMatrix(intRow, ����б�.�ֳɱ���) = FormatEx(Val(.TextMatrix(intRow, ����б�.�ּ�)) / (1 + Val(.TextMatrix(intRow, ����б�.�ӳ���)) / 100), mintCostDigit)
                .TextMatrix(intRow, ����б�.��۲�) = Format((Val(.TextMatrix(intRow, ����б�.�ֳɱ���)) - .TextMatrix(intRow, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(intRow, ����б�.����)), mstrMoneyFormat)
                dbl��Ʊ��� = (Val(.TextMatrix(intRow, ����б�.�ֳɱ���)) - .TextMatrix(intRow, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(intRow, ����б�.����))
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(n, ����б�.ҩƷid) <> "" Then
                        If Val(.TextMatrix(n, ����б�.ҩƷid)) = Val(.TextMatrix(intRow, ����б�.ҩƷid)) And n <> intRow Then
                            If chk������.Value = 0 Or (Val(.TextMatrix(intRow, ����б�.����)) <> 0 And Val(.TextMatrix(intRow, ����б�.����)) = Val(.TextMatrix(n, ����б�.����))) Then
                                .TextMatrix(n, ����б�.�ӳ���) = FormatEx(.TextMatrix(intRow, ����б�.�ӳ���), 5)
                                .TextMatrix(n, ����б�.�ֳɱ���) = .TextMatrix(intRow, ����б�.�ֳɱ���)
                                .TextMatrix(n, ����б�.��۲�) = Format((Val(.TextMatrix(n, ����б�.�ֳɱ���)) - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����)), mstrMoneyFormat)
                            End If
                        End If
                        dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(n, ����б�.�ֳɱ���)) - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����))
                    End If
                Next

                If chk�Զ�����Ӧ����䶯.Value = 1 Then
                    For n = 1 To BillPay.Rows - 1
                        If BillPay.TextMatrix(1, 0) <> "" Then
                            If Val(BillPay.TextMatrix(n, Ӧ������.ҩƷid)) = Val(BillStore.TextMatrix(intRow, ����б�.ҩƷid)) Then
                                BillPay.TextMatrix(n, Ӧ������.��Ʊ���) = FormatEx(dbl��Ʊ���, 2)
                            End If
                        End If
                    Next
                End If
            Case ����б�.�ֳɱ���
                If Val(.Text) > Val(.TextMatrix(.Row, ����б�.�ּ�)) Then
                    MsgBox "ע�⣬�³ɱ��۴��������ۼۣ�", vbExclamation, gstrSysName
                End If
                
                If Val(.Text) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                End If
                
                .TextMatrix(intRow, ����б�.�ֳɱ���) = FormatEx(Val(.Text), mintCostDigit)
                If Val(.Text) <> 0 Then
                    .TextMatrix(intRow, ����б�.�ӳ���) = FormatEx((Val(.TextMatrix(intRow, ����б�.�ּ�)) / Val(.Text) - 1) * 100, 5)
                End If
                .TextMatrix(intRow, ����б�.��۲�) = Format((Val(.Text) - .TextMatrix(intRow, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(intRow, ����б�.����)), mstrMoneyFormat)
                dbl��Ʊ��� = (Val(.Text) - .TextMatrix(intRow, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(intRow, ����б�.����))
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(n, ����б�.ҩƷid) <> "" Then
                        If Val(.TextMatrix(n, ����б�.ҩƷid)) = Val(.TextMatrix(intRow, ����б�.ҩƷid)) And n <> intRow Then
                            If chk������.Value = 0 Or (Val(.TextMatrix(intRow, ����б�.����)) <> 0 And Val(.TextMatrix(intRow, ����б�.����)) = Val(.TextMatrix(n, ����б�.����))) Then
                                dbl�ֳɱ��� = Val(.Text)
                                .TextMatrix(n, ����б�.�ֳɱ���) = FormatEx(dbl�ֳɱ���, mintCostDigit)
                                If dbl�ֳɱ��� <> 0 Then
                                    .TextMatrix(n, ����б�.�ӳ���) = FormatEx((Val(.TextMatrix(n, ����б�.�ּ�)) / dbl�ֳɱ��� - 1) * 100, 5)
                                End If
                                .TextMatrix(n, ����б�.��۲�) = Format((dbl�ֳɱ��� - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����)), mstrMoneyFormat)
                            Else
                                dbl�ֳɱ��� = Val(.TextMatrix(n, ����б�.�ֳɱ���))
                            End If
                            dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֳɱ��� - .TextMatrix(n, ����б�.ԭ�ɱ���)) * Val(.TextMatrix(n, ����б�.����))
                        End If
                    End If
                Next

                If chk�Զ�����Ӧ����䶯.Value = 1 Then
                    For n = 1 To BillPay.Rows - 1
                        If BillPay.TextMatrix(1, 0) <> "" Then
                            If Val(BillPay.TextMatrix(n, Ӧ������.ҩƷid)) = Val(BillStore.TextMatrix(intRow, ����б�.ҩƷid)) Then
                                BillPay.TextMatrix(n, Ӧ������.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If
                
                If chk������.Value = 0 Then
                    For n = 1 To BillPrice.Rows - 1
                        If Val(.TextMatrix(intRow, ����б�.ҩƷid)) = Val(BillPrice.TextMatrix(n, �ۼ��б�.ҩƷid)) Then
                            BillPrice.TextMatrix(n, �ۼ��б�.�ֳɱ���) = .TextMatrix(intRow, ����б�.�ֳɱ���)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, ����б�.ҩƷid))
                End If
        End Select
    End With
End Sub

Private Sub cbo�ۼۼ��㷽ʽ_Click()
    Set mrs�ֶμӳ� = Nothing
    If cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" Then
        gstrSql = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵�� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zldatabase.OpenSQLRecord(gstrSql, "ҩƷ�ӳɷ���")
    End If
End Sub

Private Sub get�ֶμӳ��ۼ�(ByVal dbl�ɹ��� As Double, ByRef dbl�ۼ� As Double)
'���ܣ�ͨ���ɱ��۰��ֶμӳɷ�ʽ�����ۼ�
'�������ɱ���,�ۼ�
    Dim dbl��۶� As Double
    
    mdbl�ֶμӳ��� = 0
    If mrs�ֶμӳ�.EOF Then
        dbl�ۼ� = 0!
        MsgBox "û�����ý���Ϊ��" & dbl�ɹ��� & "  �ļӳ��ʣ�����ҩƷĿ¼�����ֶμӳ��ʣ������ã�"
        Exit Sub
    End If
    mrs�ֶμӳ�.MoveFirst
    Do Until mrs�ֶμӳ�.EOF
        If dbl�ɹ��� > mrs�ֶμӳ�!��ͼ� And dbl�ɹ��� <= mrs�ֶμӳ�!��߼� Then
            mdbl�ֶμӳ��� = mrs�ֶμӳ�!�ӳ��� / 100
            dbl��۶� = IIf(IsNull(mrs�ֶμӳ�!��۶�), 0, mrs�ֶμӳ�!��۶�)
            Exit Do
        End If
        mrs�ֶμӳ�.MoveNext
    Loop
    If mdbl�ֶμӳ��� = 0 Then
        MsgBox "û�����ý���Ϊ��" & dbl�ɹ��� & "  �ļӳ��ʣ�����ҩƷĿ¼�����ֶμӳ��ʣ������ã�"
        dbl�ۼ� = 0
        Exit Sub
    Else
        If dbl�ɹ��� <= 2000 Then
            dbl�ۼ� = dbl�ɹ��� * (1 + mdbl�ֶμӳ���) + dbl��۶�
        Else
            dbl�ۼ� = dbl�ɹ��� + dbl��۶�
        End If
    End If
End Sub

'Private Sub cboִ��ʱ��_Click()
'    If cboִ��ʱ��.Text = "������Ч" Then
'       cboִ��ʱ��.ListIndex = IIf(Check����δִ�м۸�, 1, 0)
'    End If
'
'    If Me.cboִ��ʱ��.Text = "������Ч" Then
'        Me.dtpRunDate.Enabled = False
'    Else
'        Me.dtpRunDate.Enabled = True
'    End If
'
'    On Error Resume Next
'    Me.BillPrice.SetFocus
'End Sub

Private Sub ChkSelect_Click()
    Dim lngRow As Long
    
    With vsfSpec
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ҩƷID"))) > 0 Then
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(ChkSelect.Value = 1, 1, 0)
            End If
        Next
    End With
End Sub

Private Sub cmdAdd_Click()
    Call GetBatchData(False)
End Sub

Private Sub cmdCanc_Click()
    Dim strTemp As String
    Dim i As Integer
    Dim j As Integer
    
    With BillPrice
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    strTemp = strTemp & "|" & txtSummary.Text & "|" & txtValuer.Text & "|" & optʱ��(0).Value & "|" & optʱ��(1).Value & "|" & dtpRunDate.Value & "|" & Chk����.Value & "|" & chk��ҩ��������.Value & "|" & _
                    chk������ & "|" & chk�Զ�����Ӧ����䶯.Value & "|" & chk�Զ����ɱ���.Value
    
    If strTemp <> mstr���м�¼ Then
        If MsgBox("�����ݱ��޸��ˣ��Ƿ��˳���", vbYesNo, gstrSysName) = vbYes Then
            lngBillId = 0
            lngMediId = 0
            lngItemID = 0
            Unload Me
        Else
            Exit Sub
        End If
    Else
        lngBillId = 0
        lngMediId = 0
        lngItemID = 0
        Unload Me
    End If
End Sub

Private Sub CmdExit_Click()
    txtItem.Text = ""
    vsfSpec.Rows = 1
    vsfSpec.Rows = 2
    ChkSelect.Value = 0
    picItem.Visible = False
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdItem_Click()
    picItem.Visible = True
    
    picItem.Left = fraCondition.Left + lblSummary.Left
    picItem.Top = fraCondition.Top + lblSummary.Top
    picItem.Width = fraCondition.Left + txtSummary.Left + txtSummary.Width
    picItem.Height = (Me.Height - fraCondition.Top) * 2 / 3
    
    txtItem.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strID As String, LngCurID As Long
    Dim ArrayID
    Dim lngAdjId As Long
    Dim strOldId As String
    Dim strNewId As String
    
    Dim Array���μ۸�
    Dim str���μ۸� As String
    Dim lngCurrBatch As Long
    Dim strTmp As String
    Dim strʱ�۷��� As String
    Dim n As Integer
    Dim i As Integer
    
    Dim lng�ⷿID As Long
    Dim lng��Ӧ��ID As Long
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str���� As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    
    Dim dbl��װ As Double
    Dim strUpdate As String
    Dim rsTmp As ADODB.Recordset
    
    Dim blnPrint As Boolean
    Dim blnIgnore As Boolean
    Dim inProc As Integer
    Dim blnOne As Boolean
    Dim blnCancel As Boolean
    
    If Me.BillPrice.Rows = 1 Then Exit Sub
    If Me.BillPrice.TextMatrix(0, �ۼ��б�.ҩƷid) = "" Then Exit Sub
    
    If Me.BillPrice.Text <> "" Then
        Call BillPrice_KeyDown(13, 0, blnCancel)
        If blnCancel = True Then Exit Sub
    End If

    '����������Ϸ���
    If CheckPrice = False Then Exit Sub
    
    '����ǽ�����������Ŀ����ôִֻ�����
    
    Err = 0: On Error GoTo ErrHand
    If mint���� = 3 Then
        gcnOracle.BeginTrans
        With Me.BillPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                If Val(.TextMatrix(intCount, �ۼ��б�.ԭ����ID)) <> Val(.TextMatrix(intCount, �ۼ��б�.������ID)) Then
                    gstrSql = "Select �շ�ϸĿid, ������Ŀid, ԭ��, �ּ�, �����շ���, �Ӱ�Ӽ���, ����˵��, ����id, ȱʡ�۸� " & _
                        " From �շѼ�Ŀ " & _
                        " Where �շ�ϸĿid = [1] And Decode(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, ��ֹ����) Is Null" & _
                        GetPriceClassString("")
                        
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, "ȡ��Ŀ��Ϣ", Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)))
                    
                    If Not rsTmp.EOF Then
                        gstrSql = "zl_�շѼ�Ŀ_update("
                        '�շ�ϸĿid_In
                        gstrSql = gstrSql & Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid))
                        '������Ŀid_In
                        gstrSql = gstrSql & "," & Val(.TextMatrix(intCount, �ۼ��б�.������ID))
                        'ԭ��_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!ԭ��), "Null", rsTmp!ԭ��)
                        '�ּ�_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!�ּ�), "Null", rsTmp!�ּ�)
                        '�����շ���_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!�����շ���), "Null", rsTmp!�����շ���)
                        '�Ӱ�Ӽ���_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!�Ӱ�Ӽ���), "Null", rsTmp!�Ӱ�Ӽ���)
                        '����˵��_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!����˵��), "Null", "'" & rsTmp!����˵�� & "'")
                        '����id_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!����id), "Null", rsTmp!����id)
                        '������_In
                        gstrSql = gstrSql & ",'" & gstrUserName & "'"
                        'ȱʡ�۸�_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!ȱʡ�۸�), "Null", rsTmp!ȱʡ�۸�)
                        gstrSql = gstrSql & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                End If
            Next
        End With
        gcnOracle.CommitTrans
        
        lngBillId = 0
        lngMediId = 0
        lngItemID = 0
        
        blnModify = False
        Unload Me
        Exit Sub
    End If
    
    dtToday = Sys.Currentdate()

    gstrSql = "select �շѼ�Ŀ_ID.nextval from dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "ȡ�շѼ�Ŀ���")
        
    lngAdjId = rsTemp.Fields(0).Value
    
    '�ٴμ���Ƿ����δִ�м۸񣬷�ֹ����
'    If chkImmediately.Value = 1 Then
        If Check����δִ�м۸� Then
            Exit Sub
        End If
'    End If
    
    mstrNo = Sys.GetNextNo(9)
    
    gcnOracle.BeginTrans
    With Me.BillPrice
        strOldId = ""
        strNewId = ""
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            If Val(.TextMatrix(intCount, �ۼ��б�.ԭ����ID)) <> Val(.TextMatrix(intCount, �ۼ��б�.������ID)) Or _
                Val(.TextMatrix(intCount, �ۼ��б�.�ּ�)) <> Val(.TextMatrix(intCount, �ۼ��б�.ԭ��)) Then
                    
                LngCurID = Sys.NextId("�շѼ�Ŀ")
                strID = strID & IIf(strID = "", "", ",") & LngCurID
                
                dbl��װ = GetModulus(Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)))
                
                If .TextMatrix(intCount, �ۼ��б�.����) = "ʱ��" And mblnʱ��ҩƷ���� And mint���� <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To BillStore.Rows - 1
                        If Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)) = Val(BillStore.TextMatrix(n, ����б�.ҩƷid)) Then
                            If InStr(1, "|" & strTmp, "|" & BillStore.TextMatrix(n, ����б�.����) & ",") = 0 Then
                                lngCurrBatch = BillStore.TextMatrix(n, ����б�.����)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & BillStore.TextMatrix(n, ����б�.����) & "," & BillStore.TextMatrix(n, ����б�.�ּ�) / dbl��װ
                            End If
                        End If
                    Next
                    str���μ۸� = str���μ۸� & strTmp
                End If
                str���μ۸� = str���μ۸� & ";"
                
                If CLng(.RowData(intCount)) <> 0 Then
                    If .RowData(intCount) <> -1 And InStr(1, strOldId & ",", "," & .RowData(intCount) & ",") > 0 Then
                        MsgBox "��һ�ε����в��ܶ���ͬƷ��(" & .TextMatrix(intCount, �ۼ��б�.Ʒ��) & ")�ظ�����", vbExclamation, gstrSysName
                        gcnOracle.RollbackTrans: .SetFocus: Exit Sub
                    End If
                    If .RowData(intCount) = -1 And InStr(1, strNewId & ",", "," & .TextMatrix(intCount, �ۼ��б�.ҩƷid) & ",") > 0 Then
                        MsgBox "���ܶ���ͬƷ��(" & .TextMatrix(intCount, �ۼ��б�.Ʒ��) & ")�ظ����ü۸�", vbExclamation, gstrSysName
                        gcnOracle.RollbackTrans: .SetFocus: Exit Sub
                    End If
                    If .RowData(intCount) <> -1 Then
                        strOldId = strOldId & "," & .RowData(intCount)
                    Else
                        strNewId = strNewId & "," & .TextMatrix(intCount, �ۼ��б�.ҩƷid)
                    End If
                    
                    '������һ�εļ۸��¼��ִֹ��
                    gstrSql = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(intCount, �ۼ��б�.ҩƷid) & ","
                    If optʱ��(0).Value = True Then
                        gstrSql = gstrSql & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSql = gstrSql & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSql = gstrSql & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    
                    '�����۸��¼
                    gstrSql = "zl_�շѼ�Ŀ_Insert(" & LngCurID & "," & IIf(.RowData(intCount) = -1, "NUll", .RowData(intCount)) & _
                              "," & .TextMatrix(intCount, �ۼ��б�.ҩƷid) & "," & Val(.TextMatrix(intCount, �ۼ��б�.������ID)) & "," & _
                              Round(Val(.TextMatrix(intCount, �ۼ��б�.ԭ��)) / dbl��װ, gtype_MaxDigits.dig_���ۼ�) & "," & _
                              Round(Val(.TextMatrix(intCount, �ۼ��б�.�ּ�)) / dbl��װ, gtype_MaxDigits.dig_���ۼ�) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If optʱ��(0).Value = True Then
                        gstrSql = gstrSql & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSql = gstrSql & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSql = gstrSql & ",0,'" & mstrNo & "'," & intCount & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    blnPrint = True
                End If
            End If
        Next
    End With
    
    '�ɱ��۵��۴���
    If mint���� = 1 Or mint���� = 2 Then
        If BillStore.TextMatrix(1, 0) <> "" Then
            If BillPrice.Rows = 2 Then
                blnOne = True
            ElseIf BillPrice.Rows = 3 Then
                If BillPrice.TextMatrix(BillPrice.Rows - 1, 0) = "" Then
                    blnOne = True
                End If
            End If
            
            For n = 1 To BillStore.Rows - 1
                If BillStore.TextMatrix(n, 0) = "" Then Exit For
                
                '���δ��˵���
                If blnOne = True Then
                    If CheckUnVerify(Val(BillStore.TextMatrix(n, ����б�.ҩƷid))) = True Then
                        If MsgBox(BillStore.TextMatrix(n, ����б�.ҩƷ) & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                            vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            gcnOracle.RollbackTrans
                            Exit Sub
                        End If
                    End If
                Else
                    If blnIgnore = False Then
                        If CheckUnVerify(Val(BillStore.TextMatrix(n, ����б�.ҩƷid))) = True Then
                            inProc = frmMsgBox.ShowMsgBox(BillStore.TextMatrix(n, ����б�.ҩƷ) & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                                vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", Me)
                            
                            If inProc = vbNo Or inProc = vbCancel Then
                                gcnOracle.RollbackTrans
                                Exit Sub
                            ElseIf inProc = vbIgnore Then
                                blnIgnore = True
                            End If
                        End If
                    End If
                End If
                
                For i = 1 To BillPay.Rows - 1
                    If BillPay.TextMatrix(i, 0) = "" Then Exit For
                    If Val(BillStore.TextMatrix(n, ����б�.ҩƷid)) = Val(BillPay.TextMatrix(i, Ӧ������.ҩƷid)) Then
                        lng�ⷿID = Val(BillStore.TextMatrix(n, ����б�.�ⷿid))
                        lng��Ӧ��ID = Val(BillStore.TextMatrix(n, ����б�.��Ӧ��ID))
                        lngҩƷID = Val(BillStore.TextMatrix(n, ����б�.ҩƷid))
                        lng���� = Val(BillStore.TextMatrix(n, ����б�.����))
                        str���� = BillStore.TextMatrix(n, ����б�.����)
                        strЧ�� = IIf(Trim(BillStore.TextMatrix(n, ����б�.Ч��)) = "", "", BillStore.TextMatrix(n, ����б�.Ч��))
                        str���� = BillStore.TextMatrix(n, ����б�.����)
                        dblOldCost = FormatEx(Val(BillStore.TextMatrix(n, ����б�.ԭ�ɱ���)) / GetModulus(lngҩƷID), gtype_MaxDigits.dig_�ɱ���)
                        dblNewCost = FormatEx(Val(BillStore.TextMatrix(n, ����б�.�ֳɱ���)) / GetModulus(lngҩƷID), gtype_MaxDigits.dig_�ɱ���)
                        str��Ʊ�� = BillPay.TextMatrix(i, Ӧ������.��Ʊ��)
                        str��Ʊ���� = Format(BillPay.TextMatrix(i, Ӧ������.��Ʊ����), "yyyy-mm-dd")
                        dbl��Ʊ��� = Val(BillPay.TextMatrix(i, Ӧ������.��Ʊ���))
                                                
                        gstrSql = "Zl_�ɱ��۵�����Ϣ_Insert(" & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID) & "," & lng�ⷿID & "," & lngҩƷID & "," & lng���� & ",'" & str���� & "'" & _
                                "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str���� & "',Null," & dblOldCost & ", " & dblNewCost & "," & IIf(str��Ʊ�� <> "", "'" & str��Ʊ�� & "'", "NULL") & "," & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl��Ʊ��� & "," & IIf(mblnӦ����¼ = True, 1, 0) & ")"
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                Next
            Next
        End If
    End If
    
    '�޿��ʱ�����ɱ���
    If mint���� = 1 Or mint���� = 2 Then
        With Me.BillPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                If .TextMatrix(intCount, �ۼ��б�.�Ƿ��п��) = "0" And Val(.TextMatrix(intCount, �ۼ��б�.ԭ�ɱ���)) <> Val(.TextMatrix(intCount, �ۼ��б�.�ֳɱ���)) Then
                    dbl��װ = GetModulus(Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)))

                    lngҩƷID = Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, �ۼ��б�.ԭ�ɱ���)) / dbl��װ, gtype_MaxDigits.dig_�ɱ���))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, �ۼ��б�.�ֳɱ���)) / dbl��װ, gtype_MaxDigits.dig_�ɱ���))
                    
                    gstrSql = "Zl_�ɱ��۵�����Ϣ_Insert(Null,Null," & lngҩƷID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0)"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
            Next
        End With
    End If
    
    '����ִ��
    If mint���� = 1 Then
        '�����ɱ��۵���ʱ
        If optʱ��(0).Value = True Then
            With Me.BillPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                    gstrSql = "zl_ҩƷ�շ���¼_Adjust(0,0,Null," & Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)) & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                Next
            End With
        End If
    Else
        '���ۼ�
        ArrayID = Split(strID, ",")
        Array���μ۸� = Split(str���μ۸�, ";")
        For intCount = 0 To UBound(ArrayID)
            If optʱ��(0).Value = True Or BillPrice.RowData(intCount + 1) = -1 Then
                gstrSql = "zl_ҩƷ�շ���¼_Adjust(" & ArrayID(intCount) & "," & Me.Chk����.Value & ",'" & Array���μ۸�(intCount) & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        Next
    End If
    
    '����ָ���۸�
    With Me.BillPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            dbl��װ = GetModulus(Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)))
            
            '����ָ�����ۼ�
            If Val(.TextMatrix(intCount, �ۼ��б�.ԭָ���ۼ�)) <> Val(.TextMatrix(intCount, �ۼ��б�.��ָ���ۼ�)) And Val(.TextMatrix(intCount, �ۼ��б�.��ָ���ۼ�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, �ۼ��б�.��ָ���ۼ�)) / dbl��װ, mintSalePriceDigit))
                
                gstrSql = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)) & ",'ָ�����ۼ�=" & strUpdate & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
            
            '���²ɹ��޼�
            If Val(.TextMatrix(intCount, �ۼ��б�.ԭ�ɹ��޼�)) <> Val(.TextMatrix(intCount, �ۼ��б�.�ֲɹ��޼�)) And Val(.TextMatrix(intCount, �ۼ��б�.�ֲɹ��޼�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, �ۼ��б�.�ֲɹ��޼�)) / dbl��װ, mintSalePriceDigit))
                                
                gstrSql = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, �ۼ��б�.ҩƷid)) & ",'ָ��������=" & strUpdate & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        Next
    End With
    
    gcnOracle.CommitTrans
    
    If blnPrint = True Then
        If MsgBox("����Ҫ��ӡ����֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1023_3", Me, "NO=" & mstrNo, "��װ��λ=" & intҩ�ⵥλ, 2)
        End If
    End If
                
    lngBillId = 0
    lngMediId = 0
    lngItemID = 0
    
    blnModify = False
    
    BillPrice.ClearBill
    BillStore.ClearBill
    BillPay.ClearBill
    
    BillPrice.SetFocus
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    Me.BillPrice.SetFocus
End Sub

Private Sub cmdPrint_Click()
   Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1023_3", Me, "NO=" & mstrNo, "��װ��λ=" & intҩ�ⵥλ, 1)
End Sub

Private Sub cmdPstor_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.BillStore.TextMatrix(1, ����б�.�ⷿ)) = "" Then Exit Sub
    
    objPrint.Title.Text = "���ۿ��䶯��"
    
    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(optʱ��(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.BillStore.MsfObj
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub CmdSelecter_Click()
    Call GetItem("")
End Sub

Private Sub dtpRunDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOk.SetFocus
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Dim j As Integer
    Dim strBillPrice As String
    
    If Not blnFirst Then Exit Sub
    blnFirst = False
    
    '-----------------������ʾ����---------------------------------
    Select Case Me.Tag
    Case "5", "6"
        intDrugType = 1
        Me.Caption = "��ҩ����"
        cmdItem.Visible = False
        chk��ҩ��������.Visible = False
    Case "7"
        intDrugType = 2
        Me.Caption = "�в�ҩ����"
        cmdItem.Visible = True
        chk��ҩ��������.Visible = True
    End Select
    
'    With cboִ��ʱ��
'        .AddItem "������Ч"
'        .AddItem "ָ��������Ч"
'    End With
    '-----------------------------------------------------------
    If lngBillId = 0 Then
        If InStr(1, mstrPrivs, "������������Ŀ") > 0 Then
            '���ж�����������Ȩ�޾Ͳ�������Ȩ����
            mint���� = 3
            optʱ��(0).Value = True
            optʱ��(0).Enabled = False
            optʱ��(1).Value = False
            chk������.Enabled = False
            dtpRunDate.Enabled = False
        ElseIf InStr(1, mstrPrivs, "�ɱ��۹���") = 0 Then
            mint���� = 0
        Else
            If frmMediPriceNavigation.GetCondition(Me, mstrPrivs, mint����, mlng��Ӧ��ID, mdbl�ӳ���, mblnӦ����¼) = False Then
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    Call GetMaxDigit    '��ȡ��󾫶�

    With cbo�ۼۼ��㷽ʽ
        .AddItem "�ۼ���ɱ��۲���������"
        .AddItem "�ۼ۰��̶���������"
        .AddItem "�ۼ۰��ֶμӳɼ���"
        .ListIndex = 0
    End With
    
    Call IniGrid
    
    If lngItemID > 0 Then
        Call IniBatchData
    Else
        Call IniData
    End If
    
    If mint���� = 0 Then
        sstabDetail.TabVisible(1) = False
        chk������.Visible = False
    ElseIf mblnӦ����¼ = False Then
        sstabDetail.TabVisible(1) = False
    End If
    
    If mint���� = 1 Then
        optʱ��(0).Value = True
        optʱ��(0).Enabled = False
        optʱ��(1).Enabled = False
    End If
    
    chk�Զ�����Ӧ����䶯.Visible = sstabDetail.TabVisible(1)
    
    If mint���� = 2 Then
        chk�Զ����ɱ���.Left = IIf(chk�Զ�����Ӧ����䶯.Visible, chk�Զ�����Ӧ����䶯.Left + chk�Զ�����Ӧ����䶯.Width + 1000, dtpRunDate.Left)
    Else
        chk�Զ����ɱ���.Visible = False
    End If
    
    If mint���� = 0 Then
        fraCondition.Height = 800
        fraCondition.Top = (fraLine.Top - fraCondition.Height - lblHelp.Height) + 80
        BillPrice.Height = fraCondition.Top - 250
    End If
    If InStr(1, mstrPrivs, "������������Ŀ") > 0 Then
        If gstrDBUser = "ZLHIS" Then
            lblInfo.Visible = True
'            If chk�Զ����ɱ���.Visible = False Then
'                lblInfo.Left = chk������.Left + chk������.Width + 1000
'            Else
                lblInfo.Left = dtpRunDate.Left
'            End If
        Else
            lblInfo.Visible = False
        End If
    Else
        lblInfo.Visible = False
'        If chk�Զ����ɱ���.Visible = False Then
'            lblInfo.Left = chk������.Left + chk������.Width + 1000
'        Else
            lblInfo.Left = dtpRunDate.Left
'        End If
    End If
    
    lbl���۷�ʽ.Left = IIf(lblInfo.Visible = True, lblInfo.Left + lblInfo.Width + 410, chk�Զ����ɱ���.Left + chk�Զ����ɱ���.Width + 410)
    lbl���۷�ʽ.Top = chk�Զ����ɱ���.Top
    cbo�ۼۼ��㷽ʽ.Left = lbl���۷�ʽ.Left + lbl���۷�ʽ.Width + 50
    cbo�ۼۼ��㷽ʽ.Top = lbl���۷�ʽ.Top - 50
    If mint���� = 2 Then
        lbl���۷�ʽ.Visible = True
        cbo�ۼۼ��㷽ʽ.Visible = True
    Else
        lbl���۷�ʽ.Visible = False
        cbo�ۼۼ��㷽ʽ.Visible = False
    End If
    
    If optʱ��(0).Value <> True Then
        optʱ��(1).Value = True
    End If
    With BillPrice
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strBillPrice = strBillPrice & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    mstr���м�¼ = ""
    mstr���м�¼ = strBillPrice & "|" & txtSummary.Text & "|" & txtValuer.Text & "|" & optʱ��(0).Value & "|" & optʱ��(1).Value & "|" & dtpRunDate.Value & "|" & Chk����.Value & "|" & chk��ҩ��������.Value & "|" & _
                    chk������ & "|" & chk�Զ�����Ӧ����䶯.Value & "|" & chk�Զ����ɱ���.Value
    
    Call SetColor
    Call RestoreWinState(Me)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cur���ۼ� As Currency
    
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        If Me.ActiveControl.Name = "lvwItem" Then
            lvwItem.Visible = False
            BillPrice.SetFocus
        Else
            cmdCanc_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If BillPrice.Col <> �ۼ��б�.�ּ� Then Exit Sub
        cur���ۼ� = frmMediPriceCpt.ShowMe(intҩ�ⵥλ, Val(BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.ҩƷid)))
        If cur���ۼ� <> 0 Then
            BillPrice.TextMatrix(BillPrice.Row, �ۼ��б�.�ּ�) = FormatEx(cur���ۼ�, mintPriceDigit)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnʱ��ҩƷ���� = (zldatabase.GetPara("ʱ��ҩƷ�����ε���", glngSys, 1023, 0) = 1)
    mbln�޼���ʾ = (zldatabase.GetPara("�޼���ʾ", glngSys, 1023, 1) = 1)
    
    mstrPrivs = ";" & GetPrivFunc(glngSys, 1023) & ";"
    
    '�ж��Ƿ���ҩ�ⵥλ��ʾ
    intҩ�ⵥλ = Val(zldatabase.GetPara(29, glngSys))
    
    mintCostDigit = GetDigit(1, 1, IIf(intҩ�ⵥλ = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(intҩ�ⵥλ = 0, 1, 4))
    mintNumberDigit = GetDigit(1, 3, IIf(intҩ�ⵥλ = 0, 1, 4))
    mintMoneyDigit = GetDigit(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    
    mintSalePriceDigit = GetDigit(1, 2, 1)
    
    blnFirst = True
End Sub

Private Sub SetColor()
    '���ƽ����б�����ɫ
    Dim i As Long
    
    For i = 1 To BillPrice.Cols - 1
        If BillPrice.ColData(i) = 5 Or BillPrice.ColData(i) = 0 Then
            BillPrice.SetColColor i, &HE7CFBA
        Else
            BillPrice.SetColColor i, vbWhite
        End If
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 8100 Then
        Me.Height = 8100
    End If
    If Me.Width < 12075 Then
        Me.Width = 12075
    End If
    
    Me.cmdOk.Left = Me.ScaleWidth - Me.cmdOk.Width - 150
    Me.cmdCanc.Left = Me.cmdOk.Left
    Me.cmdPrint.Left = Me.cmdOk.Left
    Me.cmdHelp.Left = Me.cmdOk.Left
    Me.cmdItem.Left = Me.cmdOk.Left
    
    Me.BillPrice.Width = Me.cmdOk.Left - 150
    lblHelp.Left = BillPrice.Left + 80
    lblHelp.Top = BillPrice.Top + BillPrice.Height + 80
    lblHelp.Height = 450
    lblִ��ʱ��.Left = lblHelp.Left
    optʱ��(0).Left = txtSummary.Left
    optʱ��(1).Left = optʱ��(0).Left + 1000 + optʱ��(0).Width
    Me.fraCondition.Width = Me.BillPrice.Width
    Me.fraLine.Left = Me.BillPrice.Left
    Me.fraLine.Width = Me.Width
    fraLine.Top = fraCondition.Top + fraCondition.Height + 50
    Me.txtValuer.Left = Me.fraCondition.Width - Me.txtValuer.Width
    Me.lblValuer.Left = txtValuer.Left - lblValuer.Width - 50
    Me.txtSummary.Width = lblValuer.Left - txtSummary.Left - 300
    
    Me.chk������.Left = Me.lblSummary.Left
        
    Me.dtpRunDate.Left = optʱ��(1).Left + optʱ��(1).Width + 1000
    
    If dtpRunDate.Visible = True Then
        Me.Chk����.Left = dtpRunDate.Left + dtpRunDate.Width + 1000
    Else
        Me.Chk����.Left = optʱ��(1).Left + optʱ��(1).Width + 1000
    End If
    Me.chk��ҩ��������.Left = Chk����.Left + Chk����.Width + 1000

    Me.chk�Զ�����Ӧ����䶯.Left = IIf(chk�Զ�����Ӧ����䶯.Visible = True, chk�Զ�����Ӧ����䶯.Left, Chk����.Left)
'    If Me.chk�Զ�����Ӧ����䶯.Left < chk������.Left + chk������.Width + 100 Then
'        Me.chk�Զ�����Ӧ����䶯.Left = chk������.Left + chk������.Width + 100
'    End If
     
    Me.cmdPstor.Left = Me.cmdOk.Left + Me.cmdOk.Width - Me.cmdPstor.Width
    cmdPstor.Top = fraLine.Top + fraLine.Height + 10
    sstabDetail.Top = fraLine.Top + fraLine.Height + 50
    Me.sstabDetail.Width = Me.ScaleWidth - 50
    Me.sstabDetail.Height = Me.ScaleHeight - Me.sstabDetail.Top - 50
    
    Me.BillStore.Width = sstabDetail.Width - 200
    Me.BillStore.Height = sstabDetail.Height - 500
    
    Me.BillPay.Width = Me.BillStore.Width
    Me.BillPay.Height = Me.BillStore.Height
    lbl���۷�ʽ.Left = IIf(lblInfo.Visible = True, lblInfo.Left + lblInfo.Width + 410, chk�Զ����ɱ���.Left + chk�Զ����ɱ���.Width + 410)
    lbl���۷�ʽ.Top = chk�Զ����ɱ���.Top
    cbo�ۼۼ��㷽ʽ.Left = lbl���۷�ʽ.Left + lbl���۷�ʽ.Width + 50
    cbo�ۼۼ��㷽ʽ.Top = lbl���۷�ʽ.Top - 50
        
'    lblHelp.Top = IIf(Me.cmdItem.Visible = True, Me.cmdItem.Top + Me.cmdItem.Height + 50, Me.CmdHelp.Top + Me.CmdHelp.Height + 50)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnModify Then If MsgBox("��ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    
    mstrAdjMsg = ""
    
    SaveWinState Me
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwItem_DblClick()
    Dim lngOldDrugId As Long
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    Set objItem = Me.lvwItem.SelectedItem
    If Me.lvwItem.Tag = �ۼ��б�.Ʒ�� Then
        If CheckDrugRepeat(Val(Mid(objItem.Key, 2))) = False Then Exit Sub
        
        With Me.BillPrice
            lngOldDrugId = Val(.TextMatrix(.Row, �ۼ��б�.ҩƷid))
            .TextMatrix(.Row, �ۼ��б�.ҩƷid) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, �ۼ��б�.Ʒ��) = "[" & objItem.Text & "]" & objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.���) = objItem.SubItems(Me.lvwItem.ColumnHeaders("���").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.����) = objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.��λ) = objItem.SubItems(Me.lvwItem.ColumnHeaders("��λ").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.����) = objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.ԭ�ɱ���) = objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɱ���").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.�ֳɱ���) = objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɱ���").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.ԭ�ɹ��޼�) = objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɹ��޼�").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.�ֲɹ��޼�) = objItem.SubItems(Me.lvwItem.ColumnHeaders("�ɹ��޼�").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.ԭָ���ۼ�) = objItem.SubItems(Me.lvwItem.ColumnHeaders("ָ���ۼ�").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.��ָ���ۼ�) = objItem.SubItems(Me.lvwItem.ColumnHeaders("ָ���ۼ�").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.����ϵ��) = objItem.SubItems(Me.lvwItem.ColumnHeaders("����ϵ��").Index - 1)
            .TextMatrix(.Row, �ۼ��б�.ҩ��ID) = objItem.SubItems(Me.lvwItem.ColumnHeaders("ҩ��ID").Index - 1)
            
            Call zlGetPrice(.Row, .TextMatrix(.Row, �ۼ��б�.ҩƷid), IIf(.TextMatrix(.Row, �ۼ��б�.����) = "ʱ��", True, False))
            .CmdVisible = False
            
            If mint���� = 0 Then
                .Col = �ۼ��б�.�ּ�
            ElseIf mint���� = 1 Or mint���� = 2 Then
                .Col = �ۼ��б�.�ֳɱ���
            ElseIf mint���� = 3 Then
                .Col = �ۼ��б�.��������
            End If
            
            Call GetDrugStore(.Row, Val(Mid(objItem.Key, 2)), lngOldDrugId)
        End With
    Else
        With Me.BillPrice
            .TextMatrix(.Row, �ۼ��б�.������ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, �ۼ��б�.��������) = objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1)
            .CmdVisible = False
            .Col = �ۼ��б�.��������
        End With
    End If
    Me.lvwItem.Visible = False
    BillPrice.SetFocus
    blnModify = True
End Sub

Private Sub lvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    lvwItem_DblClick
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub zlGetPrice(ByVal lngRow As Long, ByVal lngMediId As Long, ByVal blnSeason As Boolean)
    '----------------------------------------------------
    '���ܣ���дָ��ҩƷid�Ķ�Ӧ�۸���Ϣ
    '��Σ�lngMediIdҩƷID��blnSeason�Ƿ�ʱ��ҩƷ
    '----------------------------------------------------
    On Error GoTo errHandle
    If blnSeason Then
        Me.Chk����.Enabled = True
        '��ʾʱ��ҩƷ���ۣ�ȡ�����/���������Ϊ��۸�
        gstrSql = "select P.id,Decode(Nvl(K.�������,0),0,P.�ּ�,K.�����/Nvl(K.�������,1)) �ּ�,P.ִ������,P.������Ŀid,I.���� as ��������" & _
                " from �շѼ�Ŀ P,������Ŀ I," & _
                "   (Select Sum(ʵ�ʽ��) �����,Sum(ʵ������) �������" & _
                "    From ҩƷ��� Where ����=1 and ҩƷID=[1]) K" & _
                " where P.������Ŀid=I.id and P.�շ�ϸĿid=[1] " & _
                "       and (P.��ֹ���� is null or SYSDATE BETWEEN P.ִ������ AND P.��ֹ����)" & _
                GetPriceClassString("P")
    Else
        '��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
        gstrSql = "select P.id,P.�ּ�,P.ִ������,P.������Ŀid,I.���� as ��������" & _
                " from �շѼ�Ŀ P,������Ŀ I" & _
                " where P.������Ŀid=I.id and P.�շ�ϸĿid=[1] " & _
                "       and (P.��ֹ���� is null or SYSDATE BETWEEN P.ִ������ AND P.��ֹ����)" & _
                GetPriceClassString("P")
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.BillPrice.RowData(lngRow) = !ID
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.�ϴ�����) = Format(!ִ������, "YYYY-MM-DD HH:MM:SS")
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.ԭ��) = FormatEx(!�ּ� * GetModulus(lngMediId), mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.�ּ�) = FormatEx(!�ּ� * GetModulus(lngMediId), mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.������ID) = !������Ŀid
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.ԭ����ID) = !������Ŀid
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.��������) = !��������
        Else
            Me.BillPrice.RowData(lngRow) = -1
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.�ϴ�����) = Format(!ִ������, "YYYY-MM-DD HH:MM:SS")
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.ԭ��) = FormatEx(0, mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.�ּ�) = FormatEx(0, mintPriceDigit)
            If lngRow > 1 Then
                Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.������ID) = Me.BillPrice.TextMatrix(lngRow - 1, �ۼ��б�.������ID)
                Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.ԭ����ID) = Me.BillPrice.TextMatrix(lngRow - 1, �ۼ��б�.������ID)
                Me.BillPrice.TextMatrix(lngRow, �ۼ��б�.��������) = Me.BillPrice.TextMatrix(lngRow - 1, �ۼ��б�.��������)
            Else
            
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub optʱ��_Click(Index As Integer)
    Dim blnδִ�� As Boolean
    
    If optʱ��(0).Value = True Then
        blnδִ�� = Check����δִ�м۸�
        If blnδִ�� = True Then
            optʱ��(0).Value = False
            optʱ��(1).Value = True
        Else
            optʱ��(0).Value = True
            optʱ��(1).Value = False
        End If
    End If
    
    If optʱ��(0).Value = True Then
        Me.dtpRunDate.Enabled = False
    Else
        Me.dtpRunDate.Enabled = True
    End If
    
    On Error Resume Next
    Me.BillPrice.SetFocus
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    
    With CmdExit
        .Top = picItem.Height - .Height - 50
        .Left = picItem.Width - .Width - 100
    End With
    
    With cmdAdd
        .Top = CmdExit.Top
        .Left = CmdExit.Left - .Width - 50
    End With
    
    With vsfSpec
        .Left = lblItem.Left
        .Top = lblItem.Top + lblItem.Height + 100
        .Width = picItem.Width - .Left - 100
        .Height = CmdExit.Top - .Top - 50
    End With
End Sub


Private Sub txtItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtItem.Text) = "" Then Exit Sub
    Call GetItem(Trim(txtItem.Text))
End Sub


Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtpRunDate.Enabled Then Me.dtpRunDate.SetFocus
End Sub

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim bln�޿�� As Boolean
    
    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    CheckPrice = False
    With BillPrice
        For IntCheck = 1 To .Rows - 1
            If Val(.TextMatrix(IntCheck, �ۼ��б�.ҩƷid)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, �ۼ��б�.�ּ�))) Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�ּ��к��зǷ��ַ���", vbInformation, gstrSysName
                    Exit Function
                End If
'                If Val(.TextMatrix(IntCheck, �ۼ��б�.�ּ�)) = 0 Then
'                    MsgBox "��" & IntCheck & "�е�ҩƷ�ּ۲���Ϊ�գ�", vbInformation, gstrSysName
'                    Exit Function
'                End If
                
                If mint���� <> 1 Then
                    If Val(.TextMatrix(IntCheck, �ۼ��б�.ԭ����ID)) = Val(.TextMatrix(IntCheck, �ۼ��б�.������ID)) Then
                        If Val(.TextMatrix(IntCheck, �ۼ��б�.�ּ�)) = Val(.TextMatrix(IntCheck, �ۼ��б�.ԭ��)) Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�ּ���ԭ����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                If .TextMatrix(IntCheck, �ۼ��б�.����) = "ʱ��" And optʱ��(0).Value <> True And mint���� <> 1 Then
                    MsgBox "��" & IntCheck & "��Ϊʱ��ҩƷ����������Ϊ����ִ�У�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckPrice = True
End Function

Private Function Check����δִ�м۸�(Optional ByVal lngDrugId As Long = 0) As Boolean
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer
    
    Err = 0
    On Error GoTo ErrHand
    
    If lngDrugId = 0 Then
        'ѭ���ж�����ҩƷ
        For IntCheck = 1 To BillPrice.Rows - 1
            LngmediIDThis = Val(BillPrice.TextMatrix(IntCheck, �ۼ��б�.ҩƷid))
            If LngmediIDThis <> 0 Then
                If mint���� = 0 Or mint���� = 2 Then
                    '�ж��Ƿ���δִ�е���ʷ�۸�
                    gstrSql = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, LngmediIDThis)
                    
                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    MsgBox "ҩƷ" & BillPrice.TextMatrix(IntCheck, �ۼ��б�.Ʒ��) & "����δִ�м۸񣬲������ñ��ε��ۣ�", vbInformation, gstrSysName
                                    Check����δִ�м۸� = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If
                
                If mint���� = 1 Or mint���� = 2 Then
                    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
                    gstrSql = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
                    Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, LngmediIDThis)
                    
                    If RecCheck.RecordCount > 0 Then
                        MsgBox "ҩƷ" & BillPrice.TextMatrix(IntCheck, �ۼ��б�.Ʒ��) & "����δִ�гɱ��ۣ��������ñ��ε��ۣ�", vbInformation, gstrSysName
                        Check����δִ�м۸� = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint���� = 0 Or mint���� = 2 Then
            '�ж��Ƿ���δִ�е���ʷ�۸�
            gstrSql = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]" & _
                    GetPriceClassString("")
            
            Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrugId)
            
            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            Check����δִ�м۸� = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If
        
        If mint���� = 1 Or mint���� = 2 Then
            '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
            gstrSql = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
            Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrugId)
            
            If RecCheck.RecordCount > 0 Then
                Check����δִ�м۸� = True
                Exit Function
            End If
        End If
    End If
    
   
    Check����δִ�м۸� = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.BillPrice.SetFocus

End Function

Private Function GetModulus(ByVal lngҩƷID As Long) As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '����ָ��ҩƷ�ĵ�λϵ��
    If intҩ�ⵥλ = 0 Then GetModulus = 1: Exit Function
    
    '��ȡҩ���װϵ��
    gstrSql = "Select Nvl(ҩ���װ,1) ϵ�� From ҩƷ��� Where ҩƷID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
    
    If Not rsTemp.EOF Then GetModulus = rsTemp!ϵ��
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfSpec_EnterCell()
    With vsfSpec
        .Editable = flexEDNone
        If .Col = .ColIndex("ѡ��") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


