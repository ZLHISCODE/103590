VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm�ƻ�����༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ƻ����"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList img32 
      Left            =   2595
      Top             =   5955
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�ƻ�����༭.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1875
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�ƻ�����༭.frx":064A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "��һ��(&N)"
      Height          =   350
      Left            =   6570
      TabIndex        =   13
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7695
      TabIndex        =   14
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8805
      TabIndex        =   15
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   16
      Top             =   5910
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   6375
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�ƻ�����༭.frx":0C94
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12965
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "��һ��(&B)"
      Height          =   350
      Left            =   5415
      TabIndex        =   12
      Top             =   5895
      Width           =   1100
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5685
      Index           =   0
      Left            =   0
      ScaleHeight     =   5685
      ScaleWidth      =   9930
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   105
      Width           =   9930
      Begin VB.CommandButton cmd���� 
         Caption         =   "��������(&R)"
         Height          =   330
         Left            =   60
         TabIndex        =   0
         Top             =   0
         Width           =   1320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   2565
         Left            =   6570
         TabIndex        =   9
         ToolTipText     =   "Ԥ�����嵥"
         Top             =   3090
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   4524
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ListView lvwMain 
         Height          =   2430
         Left            =   -15
         TabIndex        =   2
         Top             =   375
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   4286
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "���֤��"
            Object.Tag             =   "���֤��"
            Text            =   "���֤��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "���֤Ч��"
            Object.Tag             =   "���֤Ч��"
            Text            =   "���֤Ч��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "ִ�պ�"
            Object.Tag             =   "ִ�պ�"
            Text            =   "ִ�պ�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "ִ��Ч��"
            Object.Tag             =   "ִ��Ч��"
            Text            =   "ִ��Ч��"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "˰��ǼǺ�"
            Object.Tag             =   "˰��ǼǺ�"
            Text            =   "˰��ǼǺ�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "��ַ"
            Object.Tag             =   "��ַ"
            Text            =   "��ַ"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "�绰"
            Object.Tag             =   "�绰"
            Text            =   "�绰"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "��������"
            Object.Tag             =   "��������"
            Text            =   "��������"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "�ʺ�"
            Object.Tag             =   "�ʺ�"
            Text            =   "�ʺ�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "��ϵ��"
            Object.Tag             =   "��ϵ��"
            Text            =   "��ϵ��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "������"
            Text            =   "������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "���ö�"
            Text            =   "���ö�"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMain 
         Height          =   2565
         Left            =   0
         TabIndex        =   8
         ToolTipText     =   "δ�����嵥"
         Top             =   3090
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   4524
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDATE 
         AutoSize        =   -1  'True
         Caption         =   "���ڷ�Χ:"
         Height          =   180
         Left            =   1500
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ӧ��:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   5
         Left            =   4125
         TabIndex        =   7
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ԥ��:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   8385
         TabIndex        =   6
         Top             =   2865
         Width           =   630
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���ۼ�:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   3
         Left            =   6570
         TabIndex        =   4
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ�Ӧ��:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   2
         Left            =   2100
         TabIndex        =   5
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000010&
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   2805
         Width           =   9945
      End
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   5685
      Index           =   1
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   9855
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   9915
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshԤ�� 
         Height          =   2625
         Left            =   5205
         TabIndex        =   38
         Top             =   1500
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   4630
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483628
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   975
         TabIndex        =   11
         Top             =   4500
         Width           =   8820
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4875
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4875
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   5235
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5250
         Width           =   3240
      End
      Begin ZL9BillEdit.BillEdit mshEdit 
         Height          =   2610
         Left            =   165
         TabIndex        =   10
         Top             =   1500
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4604
         Appearance      =   0
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����֪ͨ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   30
         TabIndex        =   24
         Top             =   90
         Width           =   9780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�ϼ�:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   39
         Top             =   4095
         Width           =   5055
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���γ�Ԥ����:"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   5205
         TabIndex        =   37
         Top             =   4110
         Width           =   4605
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��λ����:"
         Height          =   180
         Index           =   9
         Left            =   390
         TabIndex        =   34
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��ַ�绰:"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   33
         Top             =   825
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   32
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˰��ǼǺ�:"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   31
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Index           =   4
         Left            =   165
         TabIndex        =   30
         Top             =   4560
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   5
         Left            =   345
         TabIndex        =   29
         Top             =   4935
         Width           =   570
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   6
         Left            =   5745
         TabIndex        =   28
         Top             =   4935
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   7
         Left            =   345
         TabIndex        =   27
         Top             =   5310
         Width           =   570
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   8
         Left            =   5745
         TabIndex        =   26
         Top             =   5310
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   10
         Left            =   8055
         TabIndex        =   25
         Top             =   450
         Width           =   315
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8355
         TabIndex        =   23
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���θ���:"
         Height          =   180
         Index           =   1
         Left            =   7950
         TabIndex        =   36
         Top             =   1260
         Width           =   810
      End
   End
   Begin VB.Menu mnuIco 
      Caption         =   "�����˵�(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuHandle 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "ѡ��(&S)"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "ȡ��ѡ��(&D)"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "ȫ��ѡ��(&A)"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "ȫ��ȡ��(&C)"
      End
   End
End
Attribute VB_Name = "frm�ƻ�����༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownY As Single, msngDownX As Single

Private mintStep As Integer

Private mstrNo As String                                '���ݺ�
Private mlng��λID As Long
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mlngID As Long                                  '����ID
Private mfrmMain  As Object

Private mEditType As gEditType
Private mint��¼״̬ As RecBillStatus                   '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mErrBillStatusInfor As ErrBillStatusInfor       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mblnEdit As Boolean                             '�༭״̬
Private mblnSuccess As Boolean                          '�Ƿ��е��ݱ���ɹ�
Private mstrPrivs  As String
Private mbln��� As Boolean                           '�Ƿ�����ͨ�ĸ��,ΪFalse �Ǽƻ�����
Private mlng������� As Long                            '�������

Private mdbl�ۼ�Ӧ�� As Long
Private mdbl����Ӧ�� As Double
Private mdbl����Ԥ�� As Double
Private mdbl�ۼ�Ԥ�� As Double
Private mstrSelectTag As String
Private mstrStartDate As String
Private mstrEndDate As String
Private mintPreCol As Integer
Private mintsort As Integer

Private Enum PayHeadCol
        ���ʽ = 0
        ������
        �������
End Enum
Private Const mlngModule = 1322

'�������������
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsTemp As New Recordset

    GetDepend = False
    '��ȡ���㷽ʽ
    'by lesfeng 2009-12-2 �����Ż�
    strSQL = "Select Ӧ�ó���,���㷽ʽ,ȱʡ��־ From ���㷽ʽӦ�� Where Ӧ�ó���='������' Order by ȱʡ��־ desc"
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "���㷽ʽӦ����Ϣ��ȫ,���ڽ��㷽ʽ�����н������ã�"
        Exit Function
    End If
    
    '��ʼ������
    With rsTemp
        mshEdit.Clear
        Do While Not .EOF
                mshEdit.AddItem !���㷽ʽ
            .MoveNext
        Loop
        'mshEdit.ListIndex = 0
        .Close
    End With
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub initPayGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ�����ͷ��Ϣ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    With mshEdit
        .Cols = 3
        .TextMatrix(0, PayHeadCol.���ʽ) = "���ʽ"
        .TextMatrix(0, PayHeadCol.������) = "������"
        .TextMatrix(0, PayHeadCol.�������) = "�������"
                
        If Not RestoreFlexState(mshEdit, Me.Caption) Then
            .ColWidth(PayHeadCol.���ʽ) = 1600
            .ColWidth(PayHeadCol.������) = 1200
            .ColWidth(PayHeadCol.�������) = 1000
        End If
        
        .ColAlignment(PayHeadCol.���ʽ) = 1
        .ColAlignment(PayHeadCol.������) = 7
        .ColAlignment(PayHeadCol.�������) = 1
        
        .ColData(PayHeadCol.���ʽ) = 3
        .ColData(PayHeadCol.������) = 4
        .ColData(PayHeadCol.�������) = 4
        .LocateCol = PayHeadCol.���ʽ
        .PrimaryCol = PayHeadCol.���ʽ
        .Active = True
    End With
End Sub

Private Sub SetԤ����ͷ()
    '��ʼ��Ԥ�����
    With mshԤ��
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "���ʽ"
        .TextMatrix(0, 2) = "������"
        .TextMatrix(0, 3) = "�������"
                
        If Not RestoreFlexState(mshԤ��, Me.Caption) Then
            .ColWidth(0) = 0
            .ColWidth(1) = 1400
            .ColWidth(2) = 1200
            .ColWidth(3) = 1000
        End If
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 1
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As New Recordset
    Dim lngLoop As Long
    Dim itmTemp As ListItem
    Dim strTmp As String
    Dim str���� As String
    Dim intR As Integer
    '��ʼ���
    Call initPayGrd
    On Error GoTo errHandle
    Select Case mEditType
        Case g����
                txtInfo(1).Text = UserInfo.����
                txtInfo(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                txtInfo(3).Text = ""
                txtInfo(4).Text = ""
                lblDATE.Visible = True
                cmd����.Enabled = True
        Case g���, g�޸�, g�鿴, gȡ��
            lblDATE.Visible = False
            '��ȡ�������
            'by lesfeng 2009-12-2 �����Ż�  ȡ�� select * from ���Ӱ󶨱���
            strSQL = "Select ID,��¼״̬,NO,���,Ԥ����,��λID,���,���㷽ʽ,�������,ժҪ," & _
                     "       ������,��������,�����,�������,������� " & _
                     "  From �����¼ Where NO=[1] And ��¼״̬=[2] order by ���"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, mint��¼״̬)
            
            If rsTemp.EOF Then
                mErrBillStatusInfor = �Ѿ�ɾ��
                Exit Sub
            End If
            mlng������� = Nvl(rsTemp!�������, 0)
            mlng��λID = Nvl(rsTemp!��λID, 0)
            
            txtInfo(0).Text = Nvl(rsTemp!ժҪ)
            txtInfo(1).Text = Nvl(rsTemp!������)
            txtInfo(2).Text = Format(rsTemp!��������, "yyyy-MM-dd hh:mm:ss")
            txtInfo(3).Text = Nvl(rsTemp!�����)
            txtInfo(4).Text = Format(rsTemp!�������, "yyyy-MM-dd hh:mm:ss")
            txtNo = Nvl(rsTemp!NO)
            If mEditType = g��� Or mEditType = gȡ�� Then
                txtInfo(3).Text = UserInfo.����
                txtInfo(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            End If
                        
            With mshEdit
                .Rows = rsTemp.RecordCount + 1
                lngLoop = 1
                Do While Not rsTemp.EOF
                    .TextMatrix(lngLoop, 0) = Nvl(rsTemp!���㷽ʽ)
                    .TextMatrix(lngLoop, 1) = Format(rsTemp!���, "###0.00;-###0.00; ;")
                    .TextMatrix(lngLoop, 2) = Nvl(rsTemp!�������)
                    lngLoop = lngLoop + 1
                    rsTemp.MoveNext
                Loop
            End With
            cmd����.Enabled = False
    End Select
    
    If mlng��λID <> 0 Then
        '����ṩ�˹�Ӧ��ID���ȡ�ù�Ӧ����Ϣ
        'by lesfeng 2009-12-2 �����Ż�  ȡ�� select * from �󶨱���
        strSQL = "Select ID,�ϼ�ID,����,����,����,ĩ��,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,�绰,��������," & _
                  "       �ʺ�,��ϵ��,����ʱ��,����ʱ��,����,������,���ö�,����ί����,����ί������,������֤��,������֤����," & _
                  "       ҩ��ֱ�����,ҩ��ֱ�������,��Ȩ��,��Ȩ��,վ��" & _
                  "  From ��Ӧ�� where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID)
        
        If Not rsTemp.EOF Then
            With rsTemp
                Set itmTemp = Me.lvwMain.ListItems.Add(, "K" & !ID, Nvl(!����) & "--" & Nvl(!����), 1, 1)
                  i = 1
                  itmTemp.SubItems(i) = Nvl(!���֤��)
                  i = i + 1
                  itmTemp.SubItems(i) = Format(!���֤Ч��, "yyyy-mm-dd")
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!ִ�պ�)
                  i = i + 1
                  itmTemp.SubItems(i) = Format(!ִ��Ч��, "yyyy-mm-dd")
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!˰��ǼǺ�)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!��ַ)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!�绰)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!��������)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!�ʺ�)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!��ϵ��)
                  i = i + 1
                  strTmp = Nvl(!����)
                  str���� = ""
                  For intR = 1 To Len(strTmp)
                      If Mid(Nvl(!����), intR, 1) = 1 Then
                          Select Case intR
                              Case 1
                                  str���� = str���� & " " & "ҩƷ"
                              Case 2
                                  str���� = str���� & " " & "����"
                              Case 3
                                  str���� = str���� & " " & "�豸"
                              Case 4
                                  str���� = str���� & " " & "����"
                          End Select
                      End If
                  Next
                  itmTemp.SubItems(i) = str����
                  i = i + 1
                  itmTemp.SubItems(i) = IIf(Nvl(!������, 0) = 0, "", Nvl(!������) & "����")
                  i = i + 1
                  itmTemp.SubItems(i) = Format(Nvl(!���ö�, 0), "####0.00;-####0.00; ;")
                  If lvwMain.SelectedItem Is Nothing Then
                      itmTemp.Selected = True
                  End If
            End With
            
            lblInfo(9).Caption = "��λ����:" & rsTemp!����
            lblInfo(1).Caption = "��ַ�绰:" & IIf(IsNull(rsTemp!��ַ), "", rsTemp!��ַ) & IIf(IsNull(rsTemp!��ַ), "", "  TEL:") & IIf(IsNull(rsTemp!�绰), "", rsTemp!�绰)
            lblInfo(2).Caption = "��������:" & IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
            lblInfo(3).Caption = "˰��ǼǺ�:" & IIf(IsNull(rsTemp!˰��ǼǺ�), "", rsTemp!˰��ǼǺ�)
        End If
    End If
    If mbln��� Then
        '��������
        Call LoadPayMoney
    Else
        '���ؼƻ���������
        Call GetPlanPayMoney
    End If
    cmdBack.Enabled = False
    SetEditPro
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal bln��� As Boolean, _
    ByVal int�༭״̬ As gEditType, ByVal strPrivs As String, _
    Optional strNO As String = "", _
    Optional lng��λID As Long = 0, _
    Optional int��¼״̬ As RecBillStatus = 1, _
    Optional blnSuccess As Boolean = False)
    
    mstrNo = strNO
    mbln��� = bln���
    mblnSave = False
    mblnSuccess = False
    mEditType = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mstrPrivs = strPrivs

    mlng��λID = lng��λID
    
    mblnChange = False
    mErrBillStatusInfor = �������
    Set mfrmMain = FrmMain
        
    '�������������ϵ
    If Not GetDepend Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
     
    If mEditType = g���� Then
        mblnEdit = True
    ElseIf mEditType = g�޸� Then
        mblnEdit = True
    ElseIf mEditType = g��� Then
        mblnEdit = False
        cmdOK.Caption = "���(&V)"
    ElseIf mEditType = gȡ�� Then
        mblnEdit = False
        cmdOK.Caption = "����(&O)"
    ElseIf mEditType = g�鿴 Then
        mblnEdit = False
        cmdOK.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, ";����֪ͨ��;") = 0 Then
            cmdOK.Visible = False
        Else
            cmdOK.Visible = True
        End If
    End If
    lblTitle.Caption = GetUnitName & lblTitle.Caption
     Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub LoadPayMoney()
    '--------------------------------------------------------------
    '���ܣ���乩ѡ���Ӧ��������
    '������
    '���أ�
    '˵����
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim lngLoop As Long, lngJLoop As Long
    Dim sngAllCount As Single, sngCount As Single
    Dim lng������� As Long
    
    '��־,��Ʊ��,��ⵥ��,Ʒ��,���,��λ,����,��Ʊ���
    Call zlcommfun.ShowFlash("�������������¼,���Ժ� ...", Me)
    
    mshMain.Redraw = False
    DoEvents
    Screen.MousePointer = vbHourglass
    
    '���ݲ��������趨��¼��ȡ����
    'by lesfeng 2009-12-2 �����Ż�  �޸İ󶨱���
    lng������� = mlng�������
    If IsNull(lng�������) Then lng������� = 0
    If mEditType = g���� Then
        '����ʱ��ȡ�������Ϊ�յ�Ӧ���ѡ��
        strWhere = " and ������� Is Null"
    ElseIf mEditType = g�޸� Then
        '�༭ʱ��ȡ�������Ϊ�ջ�ǰ�༭�ĸ����������Ӧ��Ӧ����
        strWhere = " And (������� Is Null Or �������=[2])"
    Else
        '�鿴�����ʱ����ȡ��ǰ�༭�ĸ������Ӧ��Ӧ����
        strWhere = " And �������=[2]"
    End If
    '��ȡӦ�����¼
    'lblTemp(0).Caption = "δ���Ʊ�嵥"
    strSQL = "" & _
        "   Select Decode(�������,Null,'','��') As ��־,ID,�ƻ�����,��Ʊ��,��ⵥ�ݺ�," & _
        "           Ʒ��,���,������λ,to_char(����,'99999999999.9999') as ����,to_char(��Ʊ���,'99999999999.99') as ��Ʊ��� " & _
        "   From Ӧ����¼ " & _
        "   Where �ƻ����� Is Null AND ��¼״̬=1 and ������� is not null And ��¼����<>-1 And ��λID=[1]" & strWhere & _
        "   Order By ��Ʊ��"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, lng�������)
    '��ʼ�����������
    With mshMain
        .Clear
        If rsTemp.EOF Then
            Set .Recordset = Nothing
            .Rows = 2
        Else
            Set .Recordset = rsTemp
        End If
    
        .FormatString = "^��־|||^��Ʊ��|^��ⵥ��|^Ʒ��|^���|^��λ|^����|^��Ʊ���"
        .ColAlignment(0) = 4
        .ColWidth(1) = 0: .ColWidth(2) = 0
        .ColWidth(3) = 1000: .ColAlignment(3) = 1
        .ColWidth(4) = 1000: .ColAlignment(4) = 4
        .ColWidth(5) = 1800: .ColAlignment(5) = 1
        .ColWidth(6) = 800: .ColAlignment(6) = 1
        .ColWidth(7) = 800: .ColAlignment(7) = 4
        .ColWidth(8) = 1200: .ColAlignment(8) = 7
        .ColWidth(9) = 1200: .ColAlignment(9) = 7
        
        mdbl����Ӧ�� = 0
        mdbl�ۼ�Ӧ�� = 0
        
        sngCount = 0: sngAllCount = 0
        For lngLoop = 1 To .Rows - 1
            mdbl�ۼ�Ӧ�� = mdbl�ۼ�Ӧ�� + Val(.TextMatrix(lngLoop, .Cols - 1))
            If .TextMatrix(lngLoop, 0) <> "" Then
                mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(lngLoop, .Cols - 1))
            End If
        Next
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    Call SetMoneyLbl
    Call GetԤ������             '��ȡԤ����
    Call SetCmdEn
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub GetPlanPayMoney()
    '--------------------------------------------------------------
    '���ܣ����ƻ�����ʱ��ȡ�����ƻ������¼��ѡ��
    '������
    '���أ�
    '˵����
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long, lngJLoop As Long
    Dim sngAllCount As Single, sngCount As Single
    Dim lng������� As Long
    Dim strStartDate As String
    Dim strEndDate As String
    
    '��־���ƻ����ڣ��ƻ����
'    lblTemp(0).Caption = "����ƻ��嵥"
    Call zlcommfun.ShowFlash("�������������¼,���Ժ� ...", Me)
    
    mshMain.Redraw = False
    Screen.MousePointer = vbHourglass
    
    '���ݲ��������趨��¼��ȡ����
    'by lesfeng 2009-12-2 �����Ż�  �޸İ󶨱���
    lng������� = mlng�������
    If IsNull(lng�������) Then lng������� = 0
    If mEditType = g���� Then
        '����ʱ��ȡ�������Ϊ�յ�Ӧ����ƻ���ѡ��
        strWhere = " And ������� Is Null and ID in (Select ID From Ӧ����¼ where (��¼״̬=1   or ��¼״̬=3) and ��¼����<>-1 and �ƻ����� is not null) "
        strWhere = strWhere & "  and �ƻ�����  between [3] and [4]" '+1-1/24/60/60
    ElseIf mEditType = g�޸� Then
        '�༭ʱ��ȡ�������Ϊ�ջ�ǰ�༭�ĸ����������Ӧ��Ӧ����ƻ�
        strWhere = "  and ID in (Select ID From Ӧ����¼ where (��¼״̬=1  or ��¼״̬=3) and ��¼����<>-1 and �ƻ����� is not null) and (������� Is Null Or �������=[2])"
    Else
        '�鿴�����ʱ����ȡ��ǰ�༭�ĸ������Ӧ��Ӧ����ƻ�
        strWhere = " and �������=[2]"
    End If
    '����29231 by lesfeng 2010-04-23
    strStartDate = mstrStartDate & " 00:00:00"
    strEndDate = mstrEndDate & " 23:59:59"
    
    '��ȡӦ����ƻ�����
    strSQL = "" & _
        "   Select  Decode(�������,Null,'','��') As ��־,ID,�ƻ����," & _
        "           TO_CHAR(�ƻ�����,'yyyy-MM-dd') As �ƻ�����,to_char(�ƻ����,'99999999999.99') as �ƻ����,��Ʊ��,��ⵥ�ݺ�," & _
        "           Ʒ��,���,������λ,to_char(����,'999999999999.9999') as ����,ժҪ " & _
        "   From Ӧ����¼ " & _
        "   Where ��¼����=-1   And ��λID=[1]" & strWhere & _
        "   Order By ��Ʊ��"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, lng�������, CDate(strStartDate), CDate(strEndDate))
    
    With mshMain
        .Clear
        If rsTemp.EOF Then
            Set .Recordset = Nothing
            .Rows = 2
        Else
            Set .Recordset = rsTemp
        End If
    
        .FormatString = "^��־|||^�ƻ�����|^�ƻ����|^��Ʊ��|^��ⵥ��|^Ʒ��|^���|^��λ|^����|^ժҪ"
        
        .ColAlignment(0) = 4
        .ColWidth(1) = 0: .ColWidth(2) = 0
        .ColWidth(3) = 1000: .ColAlignment(3) = 4
        .ColWidth(4) = 1200: .ColAlignment(4) = 7
        .ColWidth(5) = 1000: .ColAlignment(5) = 4
        .ColWidth(6) = 1000: .ColAlignment(6) = 4
        .ColWidth(7) = 1800: .ColAlignment(7) = 1
        .ColWidth(8) = 800: .ColAlignment(8) = 1
        .ColWidth(9) = 800: .ColAlignment(9) = 4
        .ColWidth(10) = 1200: .ColAlignment(10) = 7
        .ColWidth(11) = 2000: .ColAlignment(11) = 1
        
        mdbl�ۼ�Ӧ�� = 0
        mdbl����Ӧ�� = 0
        sngCount = 0: sngAllCount = 0
        For lngLoop = 1 To .Rows - 1
            mdbl�ۼ�Ӧ�� = mdbl�ۼ�Ӧ�� + Val(.TextMatrix(lngLoop, 4))
            If Trim(.TextMatrix(lngLoop, 0)) <> "" Then
                mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(lngLoop, 4))
            End If
        Next
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    Call SetMoneyLbl
    Call SetCmdEn
    
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    GetԤ������         '��ȡԤ����
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub GetԤ������()
    '--------------------------------------------------------------
    '���ܣ���ȡ�����Ԥ�����¼��ѡ��
    '������
    '���أ�
    '˵����
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long
    Dim lng������� As Long
    
    '��־,���㷽ʽ,������,�������
    'by lesfeng 2009-12-2 �����Ż�  �޸İ󶨱���
    lng������� = mlng�������
    If IsNull(lng�������) Then lng������� = 0
    Call zlcommfun.ShowFlash("��������Ԥ�����¼,���Ժ� ...", Me)
    Screen.MousePointer = vbHourglass
    
    If mEditType = g���� Then
        strWhere = " And ������� Is Null"
    ElseIf mEditType = g�޸� Then
        strWhere = " and (������� Is Null Or �������=[2])"
    Else
        strWhere = " And �������=[2]"
    End If
    
    strSQL = "" & _
        "   Select Decode(�������,Null,'','��') As ��־,ID,���㷽ʽ,���,������� " & _
        "   From �����¼ " & _
        "   Where ������� Is not  Null And ( ��¼״̬=1 and Ԥ����=1)  And ��λID=[1]" & strWhere & _
        "   Order By ID"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, lng�������)
    
    mshList.Redraw = False
    mshList.Clear
    mshList.Tag = 0
    
    If rsTemp.EOF Then
        Set mshList.Recordset = Nothing
        mshList.Rows = 2
    Else
        Set mshList.Recordset = rsTemp
        mshList.Row = 1: mshList.Col = 1
    End If
    
    With mshList
        .FormatString = "^��־||^���㷽ʽ|^������|^�������"
        .ColAlignment(0) = 4
        .ColWidth(1) = 0
        .ColWidth(2) = 1000: .ColAlignment(2) = 4
        .ColWidth(3) = 1200: .ColAlignment(3) = 7
        .ColWidth(4) = 1000: .ColAlignment(4) = 1
        mdbl�ۼ�Ԥ�� = 0
        mdbl����Ԥ�� = 0
        For lngLoop = 1 To .Rows - 1
            mdbl�ۼ�Ԥ�� = mdbl�ۼ�Ԥ�� + Val(.TextMatrix(lngLoop, 3))
            If Trim(.TextMatrix(lngLoop, 0)) = "��" Then
                mdbl����Ԥ�� = mdbl����Ԥ�� + Val(.TextMatrix(lngLoop, 3))
            End If
            
            If Val(.TextMatrix(lngLoop, 3)) < 0 Then
                    Call SetMshRowColor(mshList, lngLoop, vbRed)
            Else
                    Call SetMshRowColor(mshList, lngLoop, &H0&)
            End If
            
        Next
    End With
    mshList.Redraw = True
    
    Call SetMoneyLbl
    SetCmdEn
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub FullԤ��()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��䱾��Ԥ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim lngRow As Long
    With mshԤ��
        .Clear
        Call SetԤ����ͷ
        .Rows = 2
        lngRow = 1
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(lngLoop, 0)) = "��" Then
                .TextMatrix(lngRow, 0) = mshList.TextMatrix(lngLoop, 1)
                .TextMatrix(lngRow, 1) = mshList.TextMatrix(lngLoop, 2)
                .TextMatrix(lngRow, 2) = mshList.TextMatrix(lngLoop, 3)
                .TextMatrix(lngRow, 3) = mshList.TextMatrix(lngLoop, 4)
                If Val(.TextMatrix(lngRow, 2)) < 0 Then
                    Call SetMshRowColor(mshԤ��, lngRow, vbRed)
                Else
                    Call SetMshRowColor(mshԤ��, lngRow, &H0&)
                End If
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
        Next
    End With
End Sub

Private Sub SetMshRowColor(ByVal mshGrid As MSHFlexGrid, ByVal lngRow As Long, ByVal oleColor As OLE_COLOR)
    '����:����ָ���е���ɫ
    Dim lngOldRow As Long, lngoldCol As Long
    Dim i As Long
    With mshGrid
        lngOldRow = .Row: lngoldCol = .Col
        .Row = lngRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = oleColor
        Next
        .Row = lngOldRow: .Col = lngoldCol
    End With
End Sub

Private Sub cmdBack_Click()
    ChangeMode 1
    cmdDown.Enabled = True
    cmdBack.Enabled = False
    SetCmdEn
    mshMain.SetFocus
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim dblCount As Double
    Dim lngRow As Long, i As Long, j As Long
    If mEditType = g���� Or mEditType = g�޸� Then
        If mdbl����Ԥ�� < 0 Then
            MsgBox "���γ�Ԥ�����ܶ��С����", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        '�������㷽ʽ��Ԥ�����ܶ���ۼ��Ƿ�Ϊ����
        Dim str���㷽ʽ As String
        Dim dbl��� As Double
        str���㷽ʽ = ","
        With mshList
            For i = 1 To .Rows - 1
                dbl��� = 0
                
                If InStr(1, str���㷽ʽ, "," & .TextMatrix(i, 2) & ",") = 0 And Trim(.TextMatrix(i, 0)) = "��" Then
                    For j = 1 To .Rows - 1
                        If .TextMatrix(i, 2) = .TextMatrix(j, 2) And Trim(.TextMatrix(j, 0)) = "��" Then
                            dbl��� = dbl��� + Val(.TextMatrix(j, 3))
                        End If
                    Next
                    If dbl��� < 0 Then
                        MsgBox "���㷽ʽΪ:" & .TextMatrix(i, 2) & "���ܶ��Ϊ����!", vbInformation + vbDefaultButton1, gstrSysName
                        Exit Sub
                    End If
                    str���㷽ʽ = str���㷽ʽ & .TextMatrix(i, 2) & ","
                End If
            Next
        End With
    
        With mshEdit
            If .Rows <= 2 And Trim(.TextMatrix(1, 0)) = "" Then
                .Rows = 2
                .PrimaryCol = 0
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                If Trim(.TextMatrix(1, 0)) = "" Then
                    .TextMatrix(1, 0) = .CboText
                End If
            End If
            If .Rows <= 2 Then
                If Val(.TextMatrix(1, 2)) = 0 Then
                    .TextMatrix(1, 1) = mdbl����Ӧ�� - mdbl����Ԥ��
                End If
            End If
            .Active = True
        End With
    End If
    '���б��γ�Ԥ��������
    Call FullԤ��
    
    Call �ϼ�
    ChangeMode 2
    If mshEdit.Enabled And mshEdit.Visible Then mshEdit.SetFocus
    cmdDown.Enabled = False
    cmdBack.Enabled = True
    SetCmdEn
    If mshEdit.Enabled Then mshEdit.SetFocus
End Sub

Private Sub cmdHelp_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strReg As String
    Dim blnSuccess As Boolean
    
    If mEditType = g�鿴 Then    '�鿴
        '��ӡ
        printbill
        Unload Me
        Exit Sub
    End If
    
    If mEditType = g��� Then        '���
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, ";����֪ͨ��;") <> 0 Then
                    printbill
                End If
            End If
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
    
    If ValidData = False Then Exit Sub
    
    If mEditType = gȡ�� Then
        If SaveStrike() = True Then
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
    
    blnSuccess = SaveCard
    mblnChange = False
    If blnSuccess = True Then
        If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, ";����֪ͨ��;") <> 0 Then
                printbill
            End If
        End If
        mblnSuccess = True
        If mEditType = g�޸� Then    '�޸�
            Unload Me
            Exit Sub
        End If
        
        GetPrivoder
    Else
        Exit Sub
    End If
    
    txtInfo(0).Text = ""
    Me.Tag = "-1"
    
    mshEdit.ClearBill
    mshԤ��.Clear
    mshԤ��.Rows = 2
    SetԤ����ͷ
    
    ChangeMode 1
    FillDeptDue
      
      
    mblnSave = False
    mblnEdit = True
    cmdBack.Enabled = False
    cmdDown.Enabled = True
End Sub

Private Sub cmd����_Click()
        Dim blnOk As Boolean
        
        If frmTimeSel.GetTimeScope(mstrStartDate, mstrEndDate, Me) = False Then Exit Sub
        
        lblDATE.Caption = "���ڷ�Χ:" & mstrStartDate & " �� " & mstrEndDate
        
        'ȷ����صĹ�Ӧ��
        Call GetPrivoder
End Sub

Private Function GetPrivoder() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ع�Ӧ��
    '--�����:
    '--������:
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim itmTemp As ListItem
    Dim intR As Integer
    Dim strWhere  As String
    Dim lng������� As Long
    Dim strStartDate As String
    Dim strEndDate As String
    
    'by lesfeng 2009-12-2 �����Ż�  �޸İ󶨱���
    lng������� = mlng�������
    If IsNull(lng�������) Then lng������� = 0
    '���ݲ��������趨��¼��ȡ����
    If mEditType = g���� Then
        '����ʱ��ȡ�������Ϊ�յ�Ӧ����ƻ���ѡ��
        strWhere = " And ������� Is Null  and ID in (Select ID From Ӧ����¼ where (��¼״̬=1 or ��¼״̬=3) and ��¼����<>-1 and �ƻ����� is not null) "
        strWhere = strWhere & " and �ƻ����� between [3] and [4]" '+1-1/24/60/60
        
    ElseIf mEditType = g�޸� Then
        '�༭ʱ��ȡ�������Ϊ�ջ�ǰ�༭�ĸ����������Ӧ��Ӧ����ƻ�
        strWhere = " and (������� Is Null Or �������=[2]) And ��λid=[1]"
    Else
        '�鿴�����ʱ����ȡ��ǰ�༭�ĸ������Ӧ��Ӧ����ƻ�
        strWhere = " and �������=[2] And ��λid=[1]"
    End If
    
    Dim strȨ�� As String
    '����29231 by lesfeng 2010-04-23
    strStartDate = mstrStartDate & " 00:00:00"
    strEndDate = mstrEndDate & " 23:59:59"
    
    strȨ�� = " and " & Get����Ȩ��(gstrPrivs)
    On Error GoTo errHandle
    strSQL = "Select ID,�ϼ�ID,����,����,����,ĩ��,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,�绰,��������," & _
                  "       �ʺ�,��ϵ��,����ʱ��,����ʱ��,����,������,���ö�,����ί����,����ί������,������֤��,������֤����," & _
                  "       ҩ��ֱ�����,ҩ��ֱ�������,��Ȩ��,��Ȩ��,վ��" & _
                  "  From ��Ӧ�� where id in (Select distinct ��λid  from Ӧ����¼ where ��¼����=-1 " & strWhere & ") " & strȨ��
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, lng�������, CDate(strStartDate), CDate(strEndDate))
    
    Dim i As Long
    Dim strTmp As String
    Dim str���� As String
    With rsTemp
        Me.lvwMain.ListItems.Clear
        Do While Not .EOF
            Set itmTemp = Me.lvwMain.ListItems.Add(, "K" & !ID, Nvl(!����) & "--" & Nvl(!����), 1, 1)
            i = 1
            itmTemp.SubItems(i) = Nvl(!���֤��)
            i = i + 1
            itmTemp.SubItems(i) = Format(!���֤Ч��, "yyyy-mm-dd")
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!ִ�պ�)
            i = i + 1
            itmTemp.SubItems(i) = Format(!ִ��Ч��, "yyyy-mm-dd")
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!˰��ǼǺ�)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!��ַ)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!�绰)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!��������)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!�ʺ�)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!��ϵ��)
            i = i + 1
            strTmp = Nvl(!����)
            str���� = ""
            For intR = 1 To Len(strTmp)
                If Mid(Nvl(!����), intR, 1) = 1 Then
                    Select Case intR
                        Case 1
                            str���� = str���� & " " & "ҩƷ"
                        Case 2
                            str���� = str���� & " " & "����"
                        Case 3
                            str���� = str���� & " " & "�豸"
                        Case 4
                            str���� = str���� & " " & "����"
                    End Select
                End If
            Next
            itmTemp.SubItems(i) = str����
            i = i + 1
            itmTemp.SubItems(i) = IIf(Nvl(!������, 0) = 0, "", Nvl(!������) & "����")
            i = i + 1
            itmTemp.SubItems(i) = Format(Nvl(!���ö�, 0), "####0.00;-####0.00; ;")
            
            If lvwMain.SelectedItem Is Nothing Then
                itmTemp.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '��ȡ�������
    If Me.lvwMain.SelectedItem Is Nothing Then
        mlng��λID = 0
    Else
        mlng��λID = Val(Mid(lvwMain.SelectedItem.Key, 2))
    End If
    Call FillDeptDue
    GetPlanPayMoney
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    SetEditPro
'    If mEditType = g���� Or mEditType = g�޸� Then
'        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
'    End If
  SetCmdEn
  mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
        mblnFirst = True
        mintStep = 0
        mstrStartDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        mstrEndDate = mstrStartDate
        lblDATE.Caption = "���ڷ�Χ:" & mstrStartDate & " �� " & mstrEndDate
        Call initCard
End Sub

Private Sub Form_Resize()
'    If Me.WindowState = 1 Then Exit Sub
'
'    cmdHelp.Move 90, Me.ScaleHeight - cmdHelp.Height - 90
'    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 90, cmdHelp.Top
'    fraTemp.Move -150, cmdHelp.Top - 90, Me.ScaleWidth + 300
'    cmdOK.Move cmdCancel.Left - 1100, cmdHelp.Top
'    cmdBack.Move cmdOK.Left - 1100, cmdHelp.Top
'    cmdDown.Move cmdOK.Left, cmdOK.Top
'    Pic_Resize 0
End Sub

Private Sub ChangeMode(intMode As Integer)

    If intMode = mintStep Then Exit Sub
    
    mintStep = intMode
    
    If mintStep = 1 Then
        Pic(1).Enabled = False
        Pic(1).Visible = False
        Pic(0).Visible = True
        Pic(0).Enabled = True
    ElseIf mintStep = 2 Then
        Pic(0).Enabled = False
        Pic(0).Visible = False
        Pic(1).Visible = True
        Pic(1).Enabled = True
    ElseIf mintStep = 3 Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim blnYes As Boolean
    If mblnChange = False Then Exit Sub
    ShowMsgbox "���Ѿ������˵�����Ϣ,�������˳��Ļ�," & vbCrLf & "�����ĵ����ݽ����ܱ���,���Ҫ�˳���?", True, blnYes
    If blnYes = True Then Exit Sub
    Cancel = 1
End Sub

Private Sub lblInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
    msngDownX = X
End Sub '

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim blnYes As Boolean
    If mlng��λID = Val(Mid(Item.Key, 2)) Then Exit Sub
    If mblnChange Then
        ShowMsgbox "���Ѿ��޸��˵�ǰ����,���ѡ����������λ," & vbCrLf & "�������������õ�����,���Ҫ�ı䵥λ��?", True, blnYes
        If blnYes = False Then
            Err = 0
            On Error GoTo ErrHand:
            lvwMain.ListItems("K" & mlng��λID).Selected = True
            Exit Sub
        End If
        mblnChange = False
    Else
        mlng��λID = Val(Mid(Item.Key, 2))
    End If
ErrHand:
    Call FillDeptDue
    '������
    Call GetPlanPayMoney
End Sub

Private Sub lvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuIco
End Sub

Private Sub mnuClear_Click()
    If Me.ActiveControl Is mshMain Then
        mshMain_DblClick
    Else
        mshList_DblClick
    End If
End Sub

Private Sub mnuClearAll_Click()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If Not (Me.ActiveControl Is mshMain) And Not (Me.ActiveControl Is mshList) Then Exit Sub
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, 0) = ""
    Next
    If objTemp Is mshMain Then
        mdbl����Ӧ�� = 0
    Else
        mdbl����Ԥ�� = 0
    End If
    
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub mnuSelect_Click()
    If Me.ActiveControl Is mshMain Then
        mshMain_DblClick
    Else
        mshList_DblClick
    End If
End Sub

Private Sub mnuSelectAll_Click()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If Not (Me.ActiveControl Is mshMain) And Not (Me.ActiveControl Is mshList) Then Exit Sub
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, 0) = "��"
    Next
    If objTemp Is mshList Then
        mdbl����Ԥ�� = mdbl�ۼ�Ԥ��
    Else
        mdbl����Ӧ�� = mdbl�ۼ�Ӧ��
    End If
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
    Next
    
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
    lvwMain.Refresh
End Sub

Private Sub mshEdit_AfterDeleteRow()
    Dim Cur��� As Currency
    Dim intLop As Integer
    
    Cur��� = 0
    
    For intLop = 1 To mshEdit.Rows - 1
        If intLop <> mshEdit.Row Then
            Cur��� = Cur��� + Val(mshEdit.TextMatrix(intLop, 1))
        End If
    Next
    
    Cur��� = (mdbl����Ӧ�� - mdbl����Ԥ��) - Cur���
    
    If Cur��� <> 0 Then
        mshEdit.TextMatrix(mshEdit.Row, 1) = Format(Cur���, "#####0.00;-#####0.00; ;")
        mshEdit.TextMatrix(mshEdit.Row, 0) = mshEdit.CboText
    End If
    Call �ϼ�
End Sub

Private Sub mshEdit_cboClick(ListIndex As Long)
    With mshEdit
        If .Col <> 0 Then Exit Sub
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshEdit_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshEdit
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshEdit_EditChange(curText As String)
    mblnChange = True
    SetCmdEn
End Sub

Private Sub mshEdit_EnterCell(Row As Long, Col As Long)
    With mshEdit
        Select Case Col
            Case 1
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case 2
                .TxtCheck = True
                .MaxLength = 10
        End Select
    End With
End Sub

Private Sub mshEdit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intLop As Integer, Cur��� As Currency, curImprest As Currency
    If mEditType <> g���� And mEditType <> g�޸� Then
        If KeyCode = vbKeyReturn And mshEdit.Row = mshEdit.Rows - 1 Then
            zlcommfun.PressKey vbKeyTab
        End If
        Exit Sub
    End If
      
    With mshEdit
        If mEditType = g���� Or mEditType = g�޸� Then mblnChange = True
        If .Col = 2 Then
            If KeyCode <> vbKeyReturn Then
                .ColData(2) = 4
                .TxtCheck = False
            Else
                .ColData(2) = 0
                .TxtCheck = True
                .TextLen = 10
            End If
        End If
        If .Col = 1 And .Row = .Rows - 1 And KeyCode = vbKeyReturn Then
            If txtInfo(0).Enabled And txtInfo(0).Visible Then txtInfo(0).SetFocus
        End If
        
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .TxtVisible = False Then Exit Sub
        
        If .Col = 1 Then
            Cur��� = 0
            For intLop = 1 To .Rows - 1
                If intLop <> .Row Then
                    Cur��� = Cur��� + Val(.TextMatrix(intLop, 1))
                End If
            Next
            
            Cur��� = (mdbl����Ӧ�� - mdbl����Ԥ��) - Cur���
            
            If Val(.Text) = 0 And Cur��� > 0 Then
                MsgBox "�������Ϊ��!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Not IsNumeric(.Text) And Trim(.Text) <> "" Then
                MsgBox "�������к��зǷ��ַ�!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(.Text) < 0 Then
                MsgBox "�����¼����Ϊ����!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(.Text) >= 10 ^ 14 - 1 Then
                MsgBox "���������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Trim(.Text) = "" Then Exit Sub
            
            Cur��� = Cur��� - IIf(Trim(.Text) = "", 0, .Text)
            If Cur��� < 0 Then
                MsgBox "��������ܶ�!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If .Row >= .Rows - 1 And Cur��� > 0 Then
                .Rows = .Rows + 1
            End If
                    
            .Text = GetFormat(.Text, 2)
            .TextMatrix(.Row, .Col) = .Text
            If Cur��� > 0 Then
                .TextMatrix(.Row + 1, 1) = GetFormat(Cur���, 2)
                .TextMatrix(.Row + 1, 0) = .CboText
            End If
            Call �ϼ�
        End If
    End With
End Sub

Private Sub mshList_DblClick()
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    With mshList
        If .Recordset Is Nothing Then Exit Sub
        
        .TextMatrix(.Row, 0) = IIf(Trim(.TextMatrix(.Row, 0)) = "", "��", "")
        If Trim(.TextMatrix(.Row, 0)) = "" Then
            mdbl����Ԥ�� = mdbl����Ԥ�� - Val(.TextMatrix(.Row, 3))
        Else
            mdbl����Ԥ�� = mdbl����Ԥ�� + Val(.TextMatrix(.Row, 3))
        End If
    End With
    Call SetMoneyLbl
    
    Call SetCmdEn
End Sub

Private Sub SetMoneyLbl()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ�ǩ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    lbl(2).Caption = "��Ԥ����ϼƣ�" & Format(mdbl����Ԥ��, "###0.00;-###0.00;0;0") & "Ԫ"
    lbl(1).Caption = "���θ��" & Format(mdbl����Ӧ��, "###0.00;-###0.00;0;0") & "Ԫ"
    lbl���(1).Caption = "�ۼ�Ӧ��:" & Format(mdbl�ۼ�Ӧ��, "###0.00;-###0.00;0.00;0.00") & ""
    lbl���(2).Caption = "������:" & Format(mdbl����Ӧ��, "###0.00;-###0.00;0.00;0.00") & ""
    lbl���(3).Caption = "Ԥ���ۼ�:" & Format(mdbl�ۼ�Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
    lbl���(4).Caption = "��Ԥ��:" & Format(mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
    lbl���(5).Caption = "����Ӧ��:" & Format(mdbl����Ӧ�� - mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
End Sub

Private Sub �ϼ�()
    Dim lngRow As Long
    Dim dblCount As Double
   '��ȡ�ϼ���
    With mshEdit
        For lngRow = 1 To .Rows - 1
            dblCount = dblCount + Val(.TextMatrix(lngRow, 1))
        Next
    End With
    lbl(3).Caption = "����ϼ�:" & Format(dblCount, "###0.00;-###0.00;0;0") & "Ԫ"
End Sub

Private Sub mshList_GotFocus()
    '
    Err = 0
    On Error Resume Next
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    Err = 0
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        mshList_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub mshList_LostFocus()
    Err = 0
    On Error Resume Next
    mshList.Col = 0
    mshList.ColSel = 0
    Err = 0
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Or mshList.Recordset Is Nothing Then Exit Sub
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub

    SetEnabled 1
    Me.PopupMenu mnuHandle
    If mshList.Enabled Then mshList.SetFocus
End Sub

Private Sub mshMain_DblClick()
    Dim intCol As Integer
    With mshMain
        If .Recordset Is Nothing Then Exit Sub
        If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
        
        .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "", "��", "")
        intCol = IIf(mbln���, .Cols - 1, 4)
        
        If Trim(.TextMatrix(.Row, 0)) = "" Then
            mdbl����Ӧ�� = mdbl����Ӧ�� - Val(.TextMatrix(.Row, intCol))
        Else
            mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(.Row, intCol))
        End If
        mblnChange = True
    End With
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub SetCmdEn()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ÿؼ�����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
   ' cmdDown.Enabled = mdbl����Ӧ�� <> 0 And mdbl����Ӧ�� - mdbl����Ԥ�� > 0
    If mEditType = g��� Or mEditType = gȡ�� Then
        cmdOK.Enabled = Me.cmdBack.Enabled
    ElseIf mEditType = g�鿴 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = Me.cmdBack.Enabled And mblnChange
    End If
End Sub

Private Sub mshMain_GotFocus()
    Err = 0
    On Error Resume Next
    mshMain.Col = 0
    mshMain.ColSel = mshMain.Cols - 1
    Err = 0
    
End Sub

Private Sub mshMain_LostFocus()
    Err = 0
    On Error Resume Next
    mshMain.Col = 0
    mshMain.ColSel = 0
    Err = 0
End Sub

Private Sub mshMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then        '
        mshMain_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub mshMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    If mshMain.Recordset Is Nothing Then Exit Sub
    If mshMain.Enabled Then mshMain.SetFocus

    SetEnabled 0
    Me.PopupMenu mnuHandle
End Sub

Private Sub mshԤ��_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyReturn Then
            zlcommfun.PressKey vbKeyTab
       End If
End Sub

Private Sub SetEnabled(iControl As Integer)
    If iControl = 1 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mnuSelect.Enabled = True
            mnuClear.Enabled = False
        Else
            mnuSelect.Enabled = False
            mnuClear.Enabled = True
        End If
    Else
        If mshMain.TextMatrix(mshMain.Row, 0) = "" Then
            mnuSelect.Enabled = True
            mnuClear.Enabled = False
        Else
            mnuSelect.Enabled = False
            mnuClear.Enabled = True
        End If
    End If
End Sub

Private Sub FillDeptDue()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ز�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ID,�ϼ�ID,����,����,����,ĩ��,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,�绰,��������," & _
                  "       �ʺ�,��ϵ��,����ʱ��,����ʱ��,����,������,���ö�,����ί����,����ί������,������֤��,������֤����," & _
                  "       ҩ��ֱ�����,ҩ��ֱ�������,��Ȩ��,��Ȩ��,վ��" & _
                  "  From ��Ӧ�� where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID)
    
    If Not rsTemp.EOF Then
        lblInfo(9).Caption = "��λ����:" & rsTemp!����
        lblInfo(1).Caption = "��ַ�绰:" & IIf(IsNull(rsTemp!��ַ), "", rsTemp!��ַ) & IIf(IsNull(rsTemp!��ַ), "", "  TEL:") & IIf(IsNull(rsTemp!�绰), "", rsTemp!�绰)
        lblInfo(2).Caption = "��������:" & IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
        lblInfo(3).Caption = "˰��ǼǺ�:" & IIf(IsNull(rsTemp!˰��ǼǺ�), "", rsTemp!˰��ǼǺ�)
    End If
    If mshMain.Enabled And mshMain.Visible Then mshMain.SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtDept_LostFocus()
    ImeLanguage False
End Sub

Private Sub txtInfo_Change(Index As Integer)
    mblnChange = True
    SetCmdEn
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 4 Then
            cmdOK.SetFocus
        Else
            txtInfo(Index + 1).SetFocus
        End If
    End If
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim strTemp As String
    Dim dblCount As Double
    If mlng��λID = 0 Then
         ShowMsgbox "��Ӧ��ѡ������,������ѡ��!"
         Call cmdBack_Click
         Exit Function
    End If
        
    With mshEdit
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, 0)) <> "" And Trim(.TextMatrix(lngRow, 1)) <> "" Then
                strTemp = Trim(.TextMatrix(lngRow, 1))
                If strTemp = "" Then
                    ShowMsgbox "�������������!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox "�������������,������!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                If Val(strTemp) < 0 Then
                    ShowMsgbox "�������С����,������!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 999999999.99 Then
                    ShowMsgbox "������ܴ���999999999.99,������!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                dblCount = dblCount + Val(strTemp)
                strTemp = Trim(.TextMatrix(lngRow, 2))
                If strTemp <> "" Then
                    If LenB(StrConv(strTemp, vbFromUnicode)) > 10 Then
                        ShowMsgbox "������볬��,���������5�����ֻ�10���ַ�!"
                        .Row = lngRow
                        .Col = 2
                        If mshEdit.Enabled Then mshEdit.SetFocus
                        Exit Function
                    End If
                    If InStr(1, strTemp, "'") <> 0 Then
                        ShowMsgbox "������벻�����뵥����!"
                        .Row = lngRow
                        .Col = 2
                        If mshEdit.Enabled Then mshEdit.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    If CCur(mdbl����Ӧ�� - (dblCount + mdbl����Ԥ��)) <> 0 Then
        ShowMsgbox "�����ƽ,���鸶��������ⵥ" & vbCrLf & "��Ʊ����Ԥ����֮���Ƿ���ͬ!"
        If mshEdit.Enabled Then mshEdit.SetFocus
        Exit Function
    End If
    If mdbl����Ӧ�� = 0 Then
        ShowMsgbox "���β������κ�Ӧ����¼,����!"
        Exit Function
    End If
    If LenB(StrConv(txtInfo(0).Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "����˵���ĳ��ȳ���!(���Ϊ50���ַ���25������)"
        txtInfo(0).SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim strNO_IN As String
    Dim int���_IN As Integer
    Dim dbl���_IN As Double
    Dim str���㷽ʽ_IN As String
    Dim str�������_IN As String
    Dim intCol   As Integer
    Dim str������_IN As String
    Dim str��������_IN As String
    Dim lng�������_IN As Long
    Dim strժҪ_IN As String
    Dim lngRow As Long
    
    SaveCard = False
    
    'txtNo = NextNo(31)
    strNO_IN = txtNo
    str������_IN = UserInfo.����
    str��������_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    strժҪ_IN = txtInfo(0).Text

    
    On Error GoTo errHandle:
    
    '��ʼ����
    gcnOracle.BeginTrans
    
    If mEditType = g���� Then
        strNO_IN = NextNo(31)
        lng�������_IN = zlDatabase.GetNextId("�����¼")
    Else
        lng�������_IN = mlng�������
        gstrSQL = "zl_�����¼_DELETE('" & strNO_IN & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
       
     Dim blnData As Boolean
     blnData = False
    'ѭ������ÿ������
    With mshEdit
        'zl_�������_INSERT( /*strNO_IN*/, /*int���_IN*/, /*intԤ����_IN*/, /*lng��λID_IN*/,
            '/*dbl���_IN*/, /*str���㷽ʽ_IN*/, /*str�������_IN*/, /*str������_IN*/, /*str��������_IN*/,
            '/*lng�������_IN*/, /*strժҪ_IN*/ );
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) <> 0 And Trim(.TextMatrix(lngRow, 0)) <> "" Then
                blnData = True
                dbl���_IN = .TextMatrix(lngRow, 1)
                str���㷽ʽ_IN = .TextMatrix(lngRow, 0)
                str�������_IN = .TextMatrix(lngRow, 2)
                
                gstrSQL = "" & _
                    "   zl_�������_INSERT('" & _
                    strNO_IN & "'," & _
                    lngRow & "," & _
                    0 & "," & _
                    mlng��λID & "," & _
                    dbl���_IN & ",'" & _
                    str���㷽ʽ_IN & "','" & _
                    str�������_IN & "','" & _
                    str������_IN & "',to_date('" & _
                    str��������_IN & "','yyyy-mm-dd HH24:MI:SS')," & _
                    lng�������_IN & ",'" & _
                    strժҪ_IN & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    If blnData = False Then
            gstrSQL = "" & _
                "   zl_�������_INSERT('" & _
                strNO_IN & "'," & _
                lngRow & "," & _
                0 & "," & _
                mlng��λID & "," & _
                dbl���_IN & ",'" & _
                "" & "','" & _
                "" & "','" & _
                str������_IN & "',to_date('" & _
                str��������_IN & "','yyyy-mm-dd HH24:MI:SS')," & _
                lng�������_IN & ",'" & _
                strժҪ_IN & "')"
         zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
    End If
    Dim strIdin As String
    Dim str�ƻ�IN As String
    strIdin = ""
    str�ƻ�IN = ""
    
    '��Ӧ�ɹ��嵥
    With mshMain
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
            
                '    Id_In       In Varchar2 := Null,
                '    �ƻ����_In In Varchar2 := Null, --��0,1,2,3��ʽ����
                '    �������_In In �����¼.�������%Type := Null,
                '    Ԥ����_In   In �����¼.Ԥ����%Type := 0,
                '    ���_In     In Ӧ����¼.��Ʊ���%Type := 0
  
                 intCol = IIf(mbln���, .Cols - 1, 4)
                gstrSQL = "zl_�������_UPDATE(" & _
                    "'" & Val(.TextMatrix(lngRow, 1)) & "'," & _
                    "'" & Val(.TextMatrix(lngRow, 2)) & "'," & _
                    lng�������_IN & "," & _
                    "0," & _
                    "" & Val(.TextMatrix(lngRow, intCol)) & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With

    strIdin = ""
    '����Ԥ����
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                'strIdin = strIdin & "," & Val(.TextMatrix(lngRow, 1))
                gstrSQL = "zl_�������_UPDATE(" & _
                    "'" & Val(.TextMatrix(lngRow, 1)) & "'," & _
                    "NULL" & "," & _
                    lng�������_IN & "," & _
                    "1," & _
                    Val(.TextMatrix(lngRow, 3)) & "" & _
                    ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    
    '�ύ����
    gcnOracle.CommitTrans
    Me.stbThis.Panels(2).Text = "���ŵ��ݺ�Ϊ:" & strNO_IN
    SaveCard = True
    Exit Function
errHandle:
    
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ༭����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 0 To 4
        txtInfo(intIndex).Enabled = mblnEdit
    Next
    mshEdit.Active = mblnEdit
    cmdOK.Enabled = (Not mblnEdit) And mEditType <> g�鿴
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��˵���
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    SaveCheck = False
    
    strNO_IN = txtNo
    On Error GoTo errHandle:
    '   zl_�������_VERIFY(NO_IN);
    gstrSQL = "zl_�������_VERIFY('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    
    SaveStrike = False
    
    strNO_IN = txtNo
    On Error GoTo errHandle:
    '   zl_�������_VERIFY(NO_IN);
    gstrSQL = "zl_�������_strike('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'��ӡ����
Private Sub printbill()
    ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "���ݱ��=" & txtNo, "��¼״̬=" & mint��¼״̬
End Sub

