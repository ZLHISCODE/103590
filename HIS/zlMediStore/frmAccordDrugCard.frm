VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAccordDrugCard 
   AutoRedraw      =   -1  'True
   Caption         =   "Э��ҩƷ��ⵥ"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmAccordDrugCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   5970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   6
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDrug 
         Height          =   3000
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Visible         =   0   'False
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5292
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   1230
         Left            =   195
         TabIndex        =   2
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2170
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   4
         Top             =   4920
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin ZL9BillEdit.BillEdit mshStructure 
         Height          =   1875
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   3307
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7230
         TabIndex        =   21
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   20
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Top             =   593
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
         TabIndex        =   16
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4995
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Э��ҩƷ��ⵥ"
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
         TabIndex        =   15
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   14
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   13
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   6645
         TabIndex        =   12
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8520
         TabIndex        =   11
         Top             =   5340
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
            Picture         =   "frmAccordDrugCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":01A8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0206
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0264
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":02C2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0320
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":037E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":03DC
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
            Picture         =   "frmAccordDrugCard.frx":043A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0498
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":04F6
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0554
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":05B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0610
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":066E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":06CC
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
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
            Picture         =   "frmAccordDrugCard.frx":072A
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
            Picture         =   "frmAccordDrugCard.frx":0FBE
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAccordDrugCard.frx":14C0
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
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
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
Attribute VB_Name = "frmAccordDrugCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Dim mstrPrivs As String                     'Ȩ��
Private mbln�¿������� As Boolean           '��Ƿ��¿�������

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼

Private mlng�ⷿid As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

Private Const MStrCaption As String = "Э��ҩƷ���"

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4


Private mcolUseCount As Collection

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

'=========================================================================================
Private Const mconIntColҩ�� As Integer = 1
Private Const mconIntCol��Ʒ�� As Integer = 2
Private Const mconIntCol��Դ As Integer = 3
Private Const mconIntCol����ҩ�� As Integer = 4
Private Const mconIntCol��� As Integer = 5
Private Const mconIntCol����ϵ�� As Integer = 6
Private Const mconIntColԭ���� As Integer = 7
Private Const mconIntCol��λ As Integer = 8
Private Const mconIntCol���� As Integer = 9
Private Const mconIntCol�ɹ��� As Integer = 10
Private Const mconIntCol�ɹ���� As Integer = 11
Private Const mconIntCol�ۼ� As Integer = 12
Private Const mconIntCol�ۼ۽�� As Integer = 13
Private Const mconintCol��� As Integer = 14
Private Const mconIntColҩƷ��������� As Integer = 15
Private Const mconIntColҩƷ���� As Integer = 16
Private Const mconIntColҩƷ���� As Integer = 17
Private Const mconIntColS As Integer = 18      '������
'=========================================================================================


'=========================================================================================
'����ҩƷ����
Private Const mconIntCol��ҩ�� As Integer = 0
Private Const mconIntCol����Ʒ�� As Integer = 1
Private Const mconIntCol����� As Integer = 2
Private Const mconIntCol������ As Integer = 3
Private Const mconIntCol����λ As Integer = 4
Private Const mconIntCol������ As Integer = 5
Private Const mconIntCol��������� As Integer = 6
Private Const mconIntCol���������� As Integer = 7
Private Const mconIntcol�ӳ��� As Integer = 8
Private Const mconintcol��ʵ�ʲ�� As Integer = 9
Private Const mconintcol��ʵ�ʽ�� As Integer = 10
Private Const mconintcol��ҩƷid As Integer = 11

Private Const mconIntCol���ɹ��� As Integer = 12
Private Const mconIntCol���ɹ���� As Integer = 13
Private Const mconIntCol���ۼ� As Integer = 14
Private Const mconIntCol���ۼ۽�� As Integer = 15
Private Const mconintCol����� As Integer = 16
Private Const mconIntCol��ҩƷ��������� As Integer = 17
Private Const mconIntCol��ҩƷ���� As Integer = 18
Private Const mconIntCol��ҩƷ���� As Integer = 19

Private Const mconInt��ColS As Integer = 20             '������
'=========================================================================================

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
                
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                                
                .Update
            End If
        Next
    End With
End Sub
Private Sub GetSysParm()
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)

End Sub
'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim int��ϵ�� As Integer, int��ϵ�� As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.ϵ��, b.���� " _
        & " FROM ҩƷ�������� A, ҩƷ������ B " _
        & "Where A.���id = B.ID " _
      & "AND A.���� = 3  "
'    Call SQLTest(App.Title, "Э��ҩƷ������", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
'    Call SQLTest
    
    If rsDepend.EOF Then
        MsgBox "û������Э��ҩƷ��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = "ϵ��=-1"
    If rsDepend.EOF Then
        MsgBox "û������Э��ҩƷ���ĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = "ϵ��=1"
    If rsDepend.EOF Then
        MsgBox "û������Э��ҩƷ��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = adFilterNone
    rsDepend.Close
    
    
    gstrSQL = " SELECT a.ҩƷid FROM Э��ҩƷ���� a, ҩƷ��� b Where a.ҩƷid = b.ҩƷid "
'    Call SQLTest(App.Title, "Э��ҩƷ������", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
'    Call SQLTest
    
    
    If rsDepend.EOF Then
        MsgBox "û��һ�־���Э��ҩƷ���յ�Э��ҩƷ,��鿴ҩƷĿ¼����", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Function
errHandle:
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
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1344)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    If mint�༭״̬ = 1 Then
    ElseIf mint�༭״̬ = 2 Then
    ElseIf mint�༭״̬ = 3 Then
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        CmdSave.Caption = "��ӡ(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill

                    mlng�ⷿid = .ItemData(.ListIndex)
                    Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                    
                    '��֯��ʽ����
                    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
                    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
                    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
                    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩƷ���������, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
'    mblnChange = False
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
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zldatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        mstrTime_End = GetBillInfo(3, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If

        If SaveCheck = True Then
            If Val(zldatabase.GetPara("��˴�ӡ", glngSys, 1344)) = 1 Then
                '��ӡ
                If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
       
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zldatabase.GetPara("���̴�ӡ", glngSys, 1344)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
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
    mshBill.ClearBill
    mshStructure.ClearBill
    Call ��ʾ�ϼƽ��
    txtժҪ.Text = ""
    mshBill.SetFocus
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Load()
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")

    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "Э��ҩƷ���", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    mlng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(3, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '����ϵͳ��������ҩ����Ա�鿴����ʱ���Ƿ���ʾ�ɱ���
    If mblnViewCost = False Then
        mshBill.ColWidth(mconIntCol�ɹ���) = 0
        mshBill.ColWidth(mconIntCol�ɹ����) = 0
        mshBill.ColWidth(mconintCol���) = 0
        mshStructure.ColWidth(mconIntCol���ɹ���) = 0
        mshStructure.ColWidth(mconIntCol���ɹ����) = 0
        mshStructure.ColWidth(mconintCol�����) = 0
    Else
        mshBill.ColWidth(mconIntCol�ɹ���) = 900
        mshBill.ColWidth(mconIntCol�ɹ����) = 900
        mshBill.ColWidth(mconintCol���) = 800
        mshStructure.ColWidth(mconIntCol���ɹ���) = 1200
        mshStructure.ColWidth(mconIntCol���ɹ����) = 1200
        mshStructure.ColWidth(mconintCol�����) = 1000
    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
        mshStructure.ColWidth(mconIntCol����Ʒ��) = IIf(mshStructure.ColWidth(mconIntCol����Ʒ��) = 0, 2000, mshStructure.ColWidth(mconIntCol����Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
        mshStructure.ColWidth(mconIntCol����Ʒ��) = 0
    End If
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim str��װϵ�� As String
    Dim intRow As Integer
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPricedigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    
    On Error GoTo errHandle
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = gstrUserName
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
                
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 3 and a.no=[1] "
                Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                With cboStock
                    .AddItem rsInitCard!����
                    .ItemData(.NewIndex) = rsInitCard!Id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д���� AS ����,'1' as ����ϵ��,"
                    str��װϵ�� = "1"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ����,B.�����װ as ����ϵ��, "
                    str��װϵ�� = "B.�����װ"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ����,B.סԺ��װ as ����ϵ��,"
                    str��װϵ�� = "B.סԺ��װ"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ����, b.ҩ���װ as ����ϵ��, "
                    str��װϵ�� = "B.ҩ���װ"
            End Select
            
            gstrSQL = " SELECT * FROM " & _
                "    (SELECT DISTINCT ���,A.ҩƷID, '[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��,F.���," & _
                strUnitQuantity & _
                "    (A.�ɱ���*" & str��װϵ�� & ") AS �ɱ���,A.�ɱ���� AS �ɱ����," & _
                "    (A.���ۼ�*" & str��װϵ�� & ") AS ���ۼ�,A.���۽�� AS ���۽��," & _
                "    A.��� AS ���,A.������,A.��������,A.�����,A.�������,A.ժҪ,B.���Ч��,B.ҩƷ��Դ,B.����ҩ��," & _
                "    F.�Ƿ���,B.�ӳ���/100 AS �ӳ��� ,B.ҩ������ AS ҩ���������� " & _
                "    FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
                "    WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID = F.ID " & _
                "    AND B.ҩƷID = E.�շ�ϸĿID(+) And E.����(+)=3 " & _
                "    AND ��¼״̬ = [2] AND A.���� = 3 AND ���ϵ��=1 " & _
                "    AND A.NO = [1])" & _
                " ORDER BY ��� "
            Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = gstrUserName
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            
            With mshBill
                Do While Not rsInitCard.EOF
                    
                    intRow = rsInitCard!���
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷID
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = IIf(IsNull(rsInitCard!ͨ����), "", rsInitCard!ͨ����)
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
                    .TextMatrix(intRow, mconIntCol��Դ) = IIf(IsNull(rsInitCard!ҩƷ��Դ), "", rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = IIf(IsNull(rsInitCard!����ҩ��), "", rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    
                    .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(rsInitCard!�ɱ����, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPricedigit)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�ӳ��� & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    rsInitCard.MoveNext
                Loop
                Dim dblCostPrice As Double
                
                If .TextMatrix(1, 0) <> "" Then
                    If SetStructure(.TextMatrix(1, 0)) <> False Then
                        If .TextMatrix(1, mconIntCol����) <> "" Then
                            GetStructureNum .TextMatrix(1, mconIntCol����) * .TextMatrix(1, mconIntCol����ϵ��), dblCostPrice, False
                        End If
                    End If
                End If
            End With
            rsInitCard.Close
                 
    End Select
    Call ��ʾ�ϼƽ��
    If mint�༭״̬ = 2 And mint����� <> 0 Then
        SetUseCountCol
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'�����޸�ǰԭ��ҩ��ʹ���������Ա������޸Ĺ����жԿ���������жϸ�׼ȷ
Private Sub SetUseCountCol()
    Dim rsUseCount As New Recordset
    Dim numUsedCount As Double
    Dim vardrug As Variant
    
'    gstrSQL = "select ҩƷid,��д����,����id from ҩƷ�շ���¼ where no='" & mstr���ݺ� & "' and ����=3 and ��¼״̬=1 and ���ϵ��=-1 "
'    Call SQLTest(App.Title, mstrCaption, gstrSQL)
'    rsUseCount.Open gstrSQL, gcnOracle
'    Call SQLTest
    On Error GoTo errHandle
    gstrSQL = "select ҩƷid,��д����,����id from ҩƷ�շ���¼ where no=[1] and ����=3 and ��¼״̬=1 and ���ϵ��=-1 "
    Set rsUseCount = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
    
    If rsUseCount.EOF Then Exit Sub
    Set mcolUseCount = New Collection
    With mcolUseCount
        Do While Not rsUseCount.EOF
            numUsedCount = 0
            For Each vardrug In mcolUseCount
                If vardrug(0) = rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)) Then
                    numUsedCount = vardrug(1)
                    .Remove vardrug(0)
                    Exit For
                End If
            Next
            
            .Add Array(rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)), numUsedCount + rsUseCount.Fields(1)), rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0))
            rsUseCount.MoveNext
        Loop
        rsUseCount.Close
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ɹ���) = "����"
        .TextMatrix(0, mconIntCol�ɹ����) = "���۽��"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconIntColԭ����) = "ԭ����"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        
        .TextMatrix(1, 0) = ""
        
        .ColWidth(0) = 0
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 1000
        If mblnViewCost = False Then
            .ColWidth(mconIntCol�ɹ���) = 0
            .ColWidth(mconIntCol�ɹ����) = 0
            .ColWidth(mconintCol���) = 0
        Else
            .ColWidth(mconIntCol�ɹ���) = 900
            .ColWidth(mconIntCol�ɹ����) = 900
            .ColWidth(mconintCol���) = 800
        End If
        .ColWidth(mconIntCol�ۼ�) = 900
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntColҩ��) = 1
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntColԭ����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            .ColData(mconIntCol����) = 4
            
            If cboStock.Enabled = True Then
                cboStock.Enabled = True
            End If
            txtժҪ.Enabled = True
        Else
            .ColData(mconIntCol����) = 5
            .ColData(mconIntColҩ��) = 0
            cboStock.Enabled = False
            txtժҪ.Enabled = False
        End If
            
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        
        
        .ColData(mconIntCol����ϵ��) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
    End With
    
    With mshStructure
        
        .Cols = mconInt��ColS
        
        .TextMatrix(0, mconIntCol��ҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol����Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol�����) = "���"
        .TextMatrix(0, mconIntCol������) = "����"
        .TextMatrix(0, mconIntCol����λ) = "��λ"
        .TextMatrix(0, mconIntCol������) = "����"
        .TextMatrix(0, mconIntCol���������) = "�������"
        .TextMatrix(0, mconIntCol����������) = "��������"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconintcol��ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconintcol��ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconintcol��ҩƷid) = "ҩƷid"
        
        .TextMatrix(0, mconIntCol���ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol���ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol���ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol���ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol�����) = "���"
        .TextMatrix(0, mconIntCol��ҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol��ҩƷ����) = "ҩƷ����"
        
        
        .ColWidth(mconIntCol��ҩ��) = 2500
        .ColWidth(mconIntCol����Ʒ��) = 2000
        .ColWidth(mconIntCol�����) = 1000
        .ColWidth(mconIntCol������) = 1000
        .ColWidth(mconIntCol����λ) = 500
        .ColWidth(mconIntCol������) = 1000
        .ColWidth(mconIntCol���������) = 0
        .ColWidth(mconIntCol����������) = 0
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconintcol��ʵ�ʲ��) = 0
        .ColWidth(mconintcol��ʵ�ʽ��) = 0
        .ColWidth(mconintcol��ҩƷid) = 0
        
        If mblnViewCost = False Then
            .ColWidth(mconIntCol���ɹ���) = 0
            .ColWidth(mconIntCol���ɹ����) = 0
            .ColWidth(mconintCol�����) = 0
        Else
            .ColWidth(mconIntCol���ɹ���) = 1000
            .ColWidth(mconIntCol���ɹ����) = 1200
            .ColWidth(mconintCol�����) = 1000
        End If
        .ColWidth(mconIntCol���ۼ�) = 1000
        .ColWidth(mconIntCol���ۼ۽��) = 1200
                
        .ColWidth(mconIntCol��ҩƷ���������) = 0
        .ColWidth(mconIntCol��ҩƷ����) = 0
        .ColWidth(mconIntCol��ҩƷ����) = 0
        
        .ColAlignment(mconIntCol����Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol���ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ɹ����) = flexAlignRightCenter
        .ColAlignment(mconintCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconIntCol������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�����) = flexAlignLeftCenter
    End With
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
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
    
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
    End With
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
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
        .Left = mshBill.Left + mshBill.Width - .Width
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
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With mshStructure
        .Left = mshBill.Left
        .Width = mshBill.Width
        .Top = txtժҪ.Top - 60 - .Height
    End With
    
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = mshStructure.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 50
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\Э��ҩƷ���", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mshDrug.Visible Then
        mshDrug.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
End Sub

Private Function CheckBuildupNumStore() As String
    '���Э��ҩƷ��ԭ��ҩ��������Ƿ��㹻
    '����ֵ����-��ʾ�����㹻����Ϊ��-��ʾ��������
    Dim intRow As Integer
    Dim dblNum��� As Double
    Dim dblNum As Double
    Dim rstemp As ADODB.Recordset
    Dim strKey As String
    Dim collNum As Collection
    Dim vardrug As Variant
    Dim strArray As String
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim lngҩƷid As Long
    
    With mshBill
        If .rows <= 1 Then Exit Function
        
        Set collNum = New Collection
        
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, 0)) <> 0 Then
                gstrSQL = "Select Distinct b.ҩƷid As ԭ��ҩid, (a.���� / a.��ĸ) As ���, b.����ϵ�� As ԭ��ҩ����ϵ��, c.ʵ������ As ԭ��ҩ���" & vbNewLine & _
                    "From Э��ҩƷ���� A, ҩƷ��� B, ҩƷ��� C" & vbNewLine & _
                    "Where a.Э��ҩƷid = b.ҩƷid And b.ҩƷid = c.ҩƷid(+) And a.ҩƷid = [1] And c.�ⷿid = [2]"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ���ϵ��", Val(.TextMatrix(intRow, 0)), cboStock.ItemData(cboStock.ListIndex))
                If rstemp.RecordCount > 0 Then
                    If rstemp!ԭ��ҩ����ϵ�� <> 0 Then
                        dblNum��� = rstemp!��� * Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                    End If
                    
                    For Each vardrug In collNum
                        If vardrug(0) = rstemp!ԭ��ҩid & "" Then
                            dblNum = vardrug(1)
                            collNum.Remove vardrug(0)
                            Exit For
                        End If
                    Next
                    strKey = rstemp!ԭ��ҩid
                    '����С��λ�����������������ʱ�����������ݱȽ�
                    strArray = dblNum + dblNum���
                    collNum.Add Array(strKey, strArray), strKey
                End If
            End If
        Next
        
        For Each varNum In collNum
            lngҩƷid = varNum(0)  '��ʽ��ҩƷid,����
            dblNum = varNum(1)
            
            'ֻ�����������ж�
            If dblNum > 0 Then
                gstrSQL = "Select (a.ʵ������ - [1]) As ʣ������, b.����" & vbNewLine & _
                            "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                            "Where a.ҩƷid = b.Id And a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] And b.��� In ('5', '6', '7') And a.���� = 1"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�����", dblNum, lngҩƷid, cboStock.ItemData(cboStock.ListIndex), 0)
                If rstemp.RecordCount = 0 Then
                    gstrSQL = "select ���� from �շ���ĿĿ¼ where id=[1]"
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷid)
                    CheckBuildupNumStore = rstemp!����
                    Exit Function
                Else
                    If rstemp!ʣ������ >= 0 Then
                        CheckBuildupNumStore = ""
                    Else
                        CheckBuildupNumStore = rstemp!����
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Function SaveCheck() As Boolean
    Dim strҩƷ As String
    Dim mbln��ʾ��ʽ  As Boolean
    '�����
    strҩƷ = CheckBuildupNumStore
    If strҩƷ <> "" Then
        If mint����� = 1 Then '��������
            If MsgBox("ԭ��ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                mbln��ʾ��ʽ = True
            End If
        ElseIf mint����� = 2 Then '�����ֹ
            MsgBox "ԭ��ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    mblnSave = False
    SaveCheck = False
    gstrSQL = "zl_Э�����_Verify('" & txtNo.Tag & "','" & gstrUserName & "')"
    On Error GoTo errHandle
    Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
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
    
    With mshStructure
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol��ҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Sub mshBill_AfterDeleteRow()
    With mshBill
        If .Row > 1 Then
            .Row = .Row - 1
        Else
            .Row = 1
        End If
        If .TextMatrix(.Row, 0) = "" Then
            mshStructure.ClearBill
        Else
            Dim dblCostPrice As Double
            
            If SetStructure(.TextMatrix(.Row, 0)) Then
                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                    GetStructureNum .TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��), dblCostPrice, False
                End If
            End If
            
        End If
        
    End With
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As New Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intStockID As Long
    Dim strUnitQuantity As String
    
    On Error GoTo errHandle
    mblnChange = True
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnitQuantity = "D.���㵥λ AS ��λ, trim(to_char(s.�������," & mstrNumberFormat & ")) AS ����,'1' as ����ϵ��," _
                & "trim(to_char(p.�ּ�," & mstrPriceFormat & ")) as �ۼ�,"
        Case mconint���ﵥλ
            strUnitQuantity = "d.���ﵥλ AS ��λ, trim(to_char(s.������� / d.�����װ," & mstrNumberFormat & ")) AS ����,TRIM(d.�����װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.�����װ," & mstrPriceFormat & ")) as �ۼ�, "
        Case mconintסԺ��λ
            strUnitQuantity = "d.סԺ��λ AS ��λ, trim(to_char(s.������� / d.סԺ��װ," & mstrNumberFormat & ")) AS ����,TRIM(d.סԺ��װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.סԺ��װ," & mstrPriceFormat & ")) as �ۼ�,"
        Case mconintҩ�ⵥλ
            strUnitQuantity = "d.ҩ�ⵥλ AS ��λ, trim(to_char(s.������� / d.ҩ���װ," & mstrNumberFormat & ")) AS ����,TRIM(d.ҩ���װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.ҩ���װ," & mstrPriceFormat & ")) as �ۼ� , "
    End Select

    intStockID = cboStock.ItemData(cboStock.ListIndex)

    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50

    gstrSQL = "" & _
        " SELECT DECODE(D.���,5,'����ҩ',6,'�г�ҩ','�в�ҩ') AS ���ʷ���,D.����,D.����,D.ͨ������" & _
        "   ,D.��Ʒ��,D.ҩƷ��Դ,D.����ҩ��,D.���,D.����,D.ҩƷID," & _
             strUnitQuantity & _
        "    S.�����, D.���Ч��,D.�Ƿ���,D.�ӳ���,D.ҩ����������,E.�ⷿ��λ " & _
        " FROM  " & _
        "    (SELECT DISTINCT J.���� ����,M.���,M.����,M.���� ͨ������,A.���� ��Ʒ��,D.ҩƷ��Դ,D.����ҩ��," & _
        "        M.���,M.����, D.ҩ��ID, D.ҩƷID, M.���㵥λ,NVL (TO_CHAR (D.���Ч��, '9999990'), 0) ���Ч��,D.���ﵥλ," & _
        "        TO_CHAR (D.�����װ, '999999999990.99999') �����װ,D.סԺ��λ,TO_CHAR (D.סԺ��װ, '999999999990.99999') סԺ��װ," & _
        "        D.ҩ�ⵥλ,TO_CHAR(D.ҩ���װ, '999999999990.99999') ҩ���װ,M.�Ƿ���,D.�ӳ���,D.ҩ������ AS ҩ���������� " & _
        "    FROM Э��ҩƷ���� F, ҩƷ���� C, ҩƷ��� D,�շ���ĿĿ¼ M,�շ���Ŀ���� A, ҩƷ���� J " & _
        "    WHERE F.ҩƷID = D.ҩƷID AND D.ҩƷID=M.ID AND D.ҩ��ID=C.ҩ��ID AND C.ҩƷ���� = J.����(+)" & _
        "    AND D.ҩƷID = A.�շ�ϸĿID(+) AND A.����(+)=3 AND NVL(D.Э��ҩƷ,0)=1 And (M.վ�� = '" & gstrNodeNo & "' Or M.վ�� is Null) " & _
        "    AND (EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID =[1] ) " & _
        "        OR M.��� =(SELECT DISTINCT 5 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1]) " & _
        "        OR M.��� =(SELECT DISTINCT 6 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1]) "
    gstrSQL = gstrSQL & _
        "        OR M.��� =(SELECT DISTINCT 7 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1])) " & _
        "    AND ( EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]) " & _
        "        OR EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID =[1]) " & _
        "        OR DECODE (�������,1,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =[1] AND ������� IN (1, 3)) " & _
        "        OR DECODE (�������,2,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =[1] AND ������� IN (2, 3))) " & _
        "    AND ( M.����ʱ�� IS NULL OR TO_CHAR (M.����ʱ��, 'YYYY-MM-DD') = '3000-01-01') ) D,�շѼ�Ŀ P," & _
        "    (SELECT ҩƷID,TRIM(TO_CHAR(SUM(��������)," & mstrNumberFormat & ")) ��������," & _
        "        TRIM(TO_CHAR(SUM (ʵ������), " & mstrNumberFormat & ")) �������," & _
        "        TRIM(TO_CHAR(SUM (ʵ�ʽ��), " & mstrMoneyFormat & ")) ����� " & _
        "    FROM ҩƷ��� " & _
        "    WHERE �ⷿID =[1] AND ����=1 " & _
        "    GROUP BY ҩƷID) S,ҩƷ�����޶� E,(Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1]) F " & _
        " WHERE D.ҩƷID=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE)" & _
        GetPriceClassString("P") & _
        " AND D.ҩƷID=S.ҩƷID(+) AND D.ҩƷID=E.ҩƷID(+) And D.ҩƷid = F.�շ�ϸĿid AND E.�ⷿID(+)=[1] " & _
        " ORDER BY D.����"
    Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID)

    If RecReturn.EOF Then Exit Sub
    Set mshDrug.Recordset = RecReturn
    RecReturn.Close
    Call SetDrugWidth(sngLeft, sngTop)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����ҩƷѡ�����Ŀ�ȼ��������
Private Sub SetDrugWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshDrug
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
        If RestoreFlexState(mshDrug, MStrCaption) = False Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 0
            
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000

        End If
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    With mshBill
        mshBill.Text = UCase(curText)
        mshBill.SelStart = Len(mshBill.Text)
    End With
    mblnChange = True
End Sub


Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            Dim dblCostPrice As Double
            
            If .TextMatrix(.Row, 0) <> "" Then
                If SetStructure(.TextMatrix(.Row, 0)) <> False Then
                    If IIf(.TextMatrix(.Row, mconIntCol����) = "", 0, .TextMatrix(.Row, mconIntCol����)) <> 0 Then
                        GetStructureNum .TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��), dblCostPrice, False
                    End If
                End If
            Else
                mshStructure.ClearBill
            End If
                
        End If
        
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mconIntCol�ɹ���
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol�ɹ����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        With mshBill
            strKey = UCase(Trim(.Text))
            
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
                        Dim RecReturn As New Recordset
                        Dim sngLeft As Single
                        Dim sngTop As Single
                        Dim intStockID As Long
                        
                        Select Case mintUnit
                            Case mconint�ۼ۵�λ
                                strUnitQuantity = "d.���㵥λ AS ��λ, TRIM(to_char(s.�������," & mstrNumberFormat & ")) AS ����,'1' as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�," & mstrPriceFormat & ")) as �ۼ�,"
                            Case mconint���ﵥλ
                                strUnitQuantity = "d.���ﵥλ AS ��λ, TRIM(to_char(s.������� / d.�����װ," & mstrNumberFormat & ")) AS ����,TRIM(d.�����װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.�����װ," & mstrPriceFormat & ")) as �ۼ�, "
                            Case mconintסԺ��λ
                                strUnitQuantity = "d.סԺ��λ AS ��λ, TRIM(to_char(s.������� / d.סԺ��װ," & mstrNumberFormat & ")) AS ����,TRIM(d.סԺ��װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.סԺ��װ," & mstrPriceFormat & ")) as �ۼ�,"
                            Case mconintҩ�ⵥλ
                                strUnitQuantity = "d.ҩ�ⵥλ AS ��λ, TRIM(to_char(s.������� / d.ҩ���װ," & mstrNumberFormat & ")) AS ����,TRIM(d.ҩ���װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.ҩ���װ," & mstrPriceFormat & ")) as �ۼ� , "
                        End Select
                        
                        intStockID = cboStock.ItemData(cboStock.ListIndex)
                        
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50

                        gstrSQL = "" & _
                        " SELECT DECODE(D.���,5,'����ҩ',6,'�г�ҩ','�в�ҩ') AS ���ʷ���,D.����,D.����,D.ͨ������,D.��Ʒ��," & _
                        "      D.ҩƷ��Դ,D.����ҩ��,D.���,D.����,D.ҩƷID," & _
                               strUnitQuantity & _
                        "      S.�����, D.���Ч��,D.�Ƿ���,D.�ӳ���,D.ҩ����������,E.�ⷿ��λ  " & _
                        " FROM  " & _
                        "     (SELECT DISTINCT J.���� ����,M.���,M.����,M.���� ͨ������,A.���� ��Ʒ��,d.ҩƷ��Դ,d.����ҩ��," & _
                        "         M.���,M.����, D.ҩ��ID, D.ҩƷID, M.���㵥λ,NVL (TO_CHAR (D.���Ч��, '9999990'), 0) ���Ч��,D.���ﵥλ," & _
                        "         TO_CHAR (D.�����װ, '999999999990.99999') �����װ,D.סԺ��λ,TO_CHAR (D.סԺ��װ, '999999999990.99999') סԺ��װ," & _
                        "         D.ҩ�ⵥλ,TO_CHAR(D.ҩ���װ, '999999999990.99999') ҩ���װ,M.�Ƿ���,D.�ӳ���,D.ҩ������ AS ҩ���������� " & _
                        "     FROM Э��ҩƷ���� F, ҩƷ���� C, ҩƷ��� D, ҩƷ���� J,�շ���ĿĿ¼ M," & _
                        "         (Select A.* From �շ���Ŀ���� A,�շ���ĿĿ¼ B" & _
                        "     Where A.�շ�ϸĿID=B.ID ANd (A.���� Like [2] Or A.���� Like [2] Or B.���� Like [2]) And A.����=" & IIf(gint���뷽ʽ = 1, 2, 1) & _
                        "         And (B.վ�� = '" & gstrNodeNo & "' Or B.վ�� is Null)) A,�շ���Ŀ���� N " & _
                        "     WHERE F.ҩƷID = D.ҩƷID AND D.ҩƷID=M.ID And D.ҩƷID=A.�շ�ϸĿID AND D.ҩ��ID=C.ҩ��ID AND C.ҩƷ���� = J.����(+)" & _
                        "     AND D.ҩƷID = N.�շ�ϸĿID(+) AND N.����(+)=3 AND NVL(D.Э��ҩƷ,0)=1 " & _
                        "     AND (EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID = [1])"
                        gstrSQL = gstrSQL & _
                        "         OR M.��� =(SELECT DISTINCT 5 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] ) " & _
                        "         OR M.��� =(SELECT DISTINCT 6 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] ) " & _
                        "         OR M.��� =(SELECT DISTINCT 7 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] )) " & _
                        "     AND ( EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =  [1] ) " & _
                        "         OR EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID = [1] ) " & _
                        "         OR DECODE (�������,1,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]  AND ������� IN (1, 3)) " & _
                        "         OR DECODE (�������,2,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]  AND ������� IN (2, 3))) " & _
                        "     AND ( M.����ʱ�� IS NULL OR TO_CHAR (M.����ʱ��, 'YYYY-MM-DD') = '3000-01-01') ) D,�շѼ�Ŀ P," & _
                        "     (SELECT ҩƷID,TO_CHAR(SUM(��������), " & mstrNumberFormat & ") ��������," & _
                        "         TO_CHAR (SUM (ʵ������), " & mstrNumberFormat & ") �������," & _
                        "         TO_CHAR (SUM (ʵ�ʽ��), " & mstrMoneyFormat & ") ����� " & _
                        "     FROM ҩƷ��� " & _
                        "     WHERE �ⷿID = [1]  AND ����=1 " & _
                        "     GROUP BY ҩƷID) S,ҩƷ�����޶� E,(Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1]) F " & _
                        " WHERE D.ҩƷID=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE)" & _
                        GetPriceClassString("P") & _
                        " AND D.ҩƷID=S.ҩƷID(+) AND D.ҩƷID=E.ҩƷID(+) And D.ҩƷid = F.�շ�ϸĿid AND E.�ⷿID(+)= [1]"
                        
                        Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                        
                        If RecReturn.EOF Then
                            MsgBox "û��ƥ���Э��ҩƷ��", vbInformation + vbOKOnly, gstrSysName
                            RecReturn.Close
                            Cancel = True
                            Exit Sub
                        ElseIf RecReturn.RecordCount = 1 Then
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!���� & "]", RecReturn!ͨ������, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                               "" & RecReturn!ҩƷ��Դ, "" & RecReturn!����ҩ��, IIf(IsNull(RecReturn!���), "", RecReturn!���), _
                               RecReturn!��λ, IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                               RecReturn!����ϵ��, RecReturn!�Ƿ���, RecReturn!�ӳ���, RecReturn!ҩ����������) = False Then
                               RecReturn.Close
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            RecReturn.Close
                        Else
                            Set mshDrug.Recordset = RecReturn
                            RecReturn.Close
                            Call SetDrugWidth(sngLeft, sngTop)
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    Call ��ʾ�����
                    'End If
                Case mconIntCol����
                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                        MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        If Val(strKey) = 0 Then
                            MsgBox "�Բ����������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) < 0.001 Then
                            MsgBox "�Բ��������������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        Dim dblCostPrice As Double
                        
                        
                        If .TextMatrix(.Row, 0) = "" Then Exit Sub
                        
                        If GetStructureNum(strKey * .TextMatrix(.Row, mconIntCol����ϵ��), dblCostPrice) = False Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(dblCostPrice * .TextMatrix(.Row, mconIntCol����ϵ��), mintCostDigit)
                        End If
                                
                        strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                        .Text = strKey
                        If .TextMatrix(.Row, mconIntCol�ɹ���) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɹ���) * strKey, mintMoneyDigit, , True)
                            If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɹ���) / (1 - Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1)), mintPriceDigit)
                            End If
                        End If
                        
                        If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                        End If
                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                        
                    End If
                    ��ʾ�ϼƽ��
                
            End Select
        End With
    ElseIf KeyCode = vbKeyDown And Shift = vbAltMask Then
        mshbill_CommandClick
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����ҩƷID�ж��Ƿ��й��õ����ҩƷ�����У���������Ӧ������
Private Function GetStructureNum(ByVal dblNum As Double, ByRef dblCostPrice As Double, _
         Optional bln�жϿ�� As Boolean = True) As Boolean
    Dim rsDrug As New Recordset
    Dim intReturn As Integer
    Dim blnContinue As Boolean      '�û���ѡ��0���˳���1����
    Dim dblConstruct As Double      'ʵ��������Ӧ���������
    Dim dblPurchase As Double       'Э��ҩƷ�ĳɱ��ۣ����У����ҩƷ�Ľ���*���������
    Dim intRow As Integer
    Dim dblԭ��д���� As Double
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim numUseCount As Double
    Dim vardrug As Variant
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           ��ۺͳɱ����ڳ��⴦���еĹ�ʽ
    '   ������=����*�ۼ�
    '   ������=������*��ʵ�ʲ��/ʵ�ʽ�
    '          ���ʵ�ʲ�ۺ�ʵ�ʽ�����ʱ��Ϊ��
    '       ������=������*ָ�������
    '   ���ۣ��ɱ���)=(������-������)/����
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    intCostDigit = mintCostDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    
    GetStructureNum = False
    blnContinue = False
    With mshStructure
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dblConstruct = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���������) * dblNum, intNumberDigit, , True)
                
                .TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(dblConstruct, intNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol���ۼ۽��) = zlStr.FormatEx(dblConstruct * .TextMatrix(intRow, mconIntCol���ۼ�), intNumberDigit, , True)
                If .TextMatrix(intRow, mconintcol��ʵ�ʽ��) <= "0" Then
'                    .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ۼ۽��) * Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(0) / 100, intMoneyDigit)
'                    .TextMatrix(intRow, mconIntCol���ɹ���) =Str.FormatEx((.TextMatrix(intRow, mconIntCol���ۼ۽��) - .TextMatrix(intRow, mconintCol�����)) / (IIf(dblConstruct = 0, 1, dblConstruct)), intCostDigit)
                    .TextMatrix(intRow, mconIntCol���ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ɹ���) * dblConstruct, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol���ۼ۽��)) - Val(.TextMatrix(intRow, mconIntCol���ɹ����)), intMoneyDigit, , True)
                Else
'                    .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ۼ۽��) * (.TextMatrix(intRow, mconintcol��ʵ�ʲ��) / .TextMatrix(intRow, mconintcol��ʵ�ʽ��)), intMoneyDigit)
'                    .TextMatrix(intRow, mconIntCol���ɹ���) =Str.FormatEx((.TextMatrix(intRow, mconIntCol���ۼ۽��) - .TextMatrix(intRow, mconintCol�����)) / (IIf(dblConstruct = 0, 1, dblConstruct)), intCostDigit)
                    .TextMatrix(intRow, mconIntCol���ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ɹ���) * dblConstruct, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol���ۼ۽��)) - Val(.TextMatrix(intRow, mconIntCol���ɹ����)), intMoneyDigit, , True)
                End If
                dblPurchase = zlStr.FormatEx(dblPurchase + .TextMatrix(intRow, mconIntCol���ɹ����) / dblNum, intCostDigit)
                
            End If
        Next
        dblCostPrice = dblPurchase
            
    End With
    
    GetStructureNum = True
End Function


'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal intҩƷid As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, _
    ByVal str����ҩ�� As String, ByVal str��� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal intԭЧ�� As Integer, ByVal num����ϵ�� As Double, _
    ByVal int�Ƿ��� As Integer, ByVal dbl�ӳ��� As Double, ByVal intҩ���������� As Integer) As Boolean
    
    Dim intCount As Integer
    Dim rsStructure As New Recordset
    Dim intCol As Integer
    Dim strҩ�� As String
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
        
        If Not SetStructure(intҩƷid) Then Exit Function
        .TextMatrix(intRow, 0) = intҩƷid
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = str��Ʒ��
        .TextMatrix(intRow, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ�, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dbl�ӳ��� / 100 & "||" & int�Ƿ��� & "||" & intҩ����������
            
    End With
    SetColValue = True
End Function

Private Function SetStructure(ByVal intҩƷid As Long) As Boolean
    Dim rsStructure As New Recordset
    Dim strҩ�� As String
    Dim rs�ɱ��� As ADODB.Recordset
    
    SetStructure = False
    mshStructure.Redraw = False
    
    On Error GoTo errHandle
    If mint�༭״̬ <> 4 Then
        gstrSQL = "SELECT DISTINCT B.ҩƷID,'[' || F.���� || ']' As ����,F.���� As ͨ������,E.���� AS ��Ʒ����, F.���, C.�ϴβ���,F.���㵥λ AS ��λ,c.ƽ���ɱ���," & _
                  " C.ʵ�ʲ��,C.ʵ�ʽ��, to_char(D.�ּ�, " & mstrPriceFormat & ") �ۼ�, " & _
                  " (A.���� / A.��ĸ) AS ���,C.��������,B.�ӳ���,F.�Ƿ���,B.ҩ������ AS ҩ����������, Nvl(F.�Ƿ���, 0) ���� " & _
                  " FROM Э��ҩƷ���� A,ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F,ҩƷ��� C, �շѼ�Ŀ D" & _
                  " WHERE A.Э��ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                  " AND A.Э��ҩƷID = D.�շ�ϸĿID AND (SYSDATE BETWEEN ִ������ AND NVL(��ֹ����,SYSDATE))" & _
                  GetPriceClassString("D") & _
                  " AND B.ҩƷID = E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                  " AND A.Э��ҩƷID = C.ҩƷID(+) AND C.�ⷿID(+) =[1] AND C.����(+)=1" & _
                  " AND (F.վ�� = [3] Or F.վ�� is Null) And A.ҩƷID =[2] "

        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), intҩƷid, gstrNodeNo)
        If rsStructure.EOF Then
            mshStructure.Redraw = True
            Exit Function
        End If
        With mshStructure
            .ClearBill
            Do While Not rsStructure.EOF
                If rsStructure!ҩ���������� = 1 Then
                    MsgBox "���ҩƷ��һ��ҩ������ҩƷ������ǰ�汾��֧��ҩ�����������ҩƷ�����飡", vbInformation + vbOKOnly, gstrSysName
                    mshStructure.Redraw = True
                    Exit Function
                End If
                If rsStructure!���� = 1 And Nvl(rsStructure!��������, 0) = 0 Then
                    MsgBox "��Э��ҩƷ�����ҩƷ��ʱ��ҩƷ�����ÿ��Ϊ0����ǰ�汾��֧�֣����飡", vbInformation + vbOKOnly, gstrSysName
                    mshStructure.Redraw = True
                    Exit Function
                End If
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsStructure!ͨ������
                Else
                    strҩ�� = IIf(IsNull(rsStructure!��Ʒ����), rsStructure!ͨ������, rsStructure!��Ʒ����)
                End If
                                                
                .TextMatrix(.Row, mconIntCol��ҩƷ���������) = rsStructure!���� & strҩ��
                .TextMatrix(.Row, mconIntCol��ҩƷ����) = rsStructure!����
                .TextMatrix(.Row, mconIntCol��ҩƷ����) = strҩ��
                
                If mintDrugNameShow = 0 Then
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ���������)
                ElseIf mintDrugNameShow = 1 Then
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ����)
                Else
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ����)
                End If
                
                .TextMatrix(.Row, mconIntCol����Ʒ��) = IIf(IsNull(rsStructure!��Ʒ����), "", rsStructure!��Ʒ����)
                
                .TextMatrix(.Row, mconIntCol�����) = IIf(IsNull(rsStructure!���), "", rsStructure!���)
                .TextMatrix(.Row, mconIntCol������) = IIf(IsNull(rsStructure!�ϴβ���), "", rsStructure!�ϴβ���)
                .TextMatrix(.Row, mconIntCol����λ) = rsStructure!��λ
                .TextMatrix(.Row, mconIntCol���ۼ�) = zlStr.FormatEx(rsStructure!�ۼ�, mintPriceDigit)
                .TextMatrix(.Row, mconIntCol����������) = zlStr.FormatEx(IIf(IsNull(rsStructure!��������), "0", rsStructure!��������), mintNumberDigit, , True)
                .TextMatrix(.Row, mconIntCol���������) = rsStructure!���
                .TextMatrix(.Row, mconIntcol�ӳ���) = rsStructure!�ӳ��� / 100 & "||" & IIf(IsNull(rsStructure!�Ƿ���), 0, rsStructure!�Ƿ���) & "||" & IIf(IsNull(rsStructure!ҩ����������), 0, rsStructure!ҩ����������)
                .TextMatrix(.Row, mconintcol��ʵ�ʲ��) = IIf(IsNull(rsStructure!ʵ�ʲ��), "0", rsStructure!ʵ�ʲ��)
                .TextMatrix(.Row, mconintcol��ʵ�ʽ��) = IIf(IsNull(rsStructure!ʵ�ʽ��), "0", rsStructure!ʵ�ʽ��)
                .TextMatrix(.Row, mconintcol��ҩƷid) = rsStructure!ҩƷID
                
                If IsNull(rsStructure!ƽ���ɱ���) Then
                    gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
                    Set rs�ɱ��� = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ�ɱ���", Val(rsStructure!ҩƷID))
                    If rs�ɱ���.RecordCount > 0 Then
                        .TextMatrix(.Row, mconIntCol���ɹ���) = zlStr.FormatEx(rs�ɱ���!�ɱ���, mintCostDigit, , True)
                    End If
                Else
                    .TextMatrix(.Row, mconIntCol���ɹ���) = zlStr.FormatEx(rsStructure!ƽ���ɱ���, mintCostDigit, , True)
                End If
                
                If .Row = .rows - 1 Then
                    .rows = .rows + 1
                End If
                .Row = .Row + 1
                rsStructure.MoveNext
            Loop
        End With
        rsStructure.Close
    Else            '�鿴
        gstrSQL = " SELECT DISTINCT A.ҩƷID,'[' || F.���� || ']' As ����,F.���� As ͨ������,E.���� AS ��Ʒ����,F.���," & _
                  "     A.����,F.���㵥λ AS ��λ,A.ʵ������,A.�ɱ���,A.�ɱ����,A.���ۼ�,A.���۽��,A.��� " & _
                  " FROM " & _
                  "     (SELECT ҩƷID,����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,��� FROM ҩƷ�շ���¼ " & _
                  "     WHERE NO=[1] AND ����=3 AND ��¼״̬=[3] " & _
                  "     AND ���ϵ��=-1 AND ����=[4] AND ����ID =[2]) A," & _
                  "     ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
                  " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                  " AND B.ҩƷID = E.�շ�ϸĿID(+) AND E.����(+)=3 "
        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNo.Tag, intҩƷid, mint��¼״̬, mshBill.Row)
        
        If rsStructure.EOF Then
            mshStructure.Redraw = True
            Exit Function
        End If
        With mshStructure
            .ClearBill
            Do While Not rsStructure.EOF
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsStructure!ͨ������
                Else
                    strҩ�� = IIf(IsNull(rsStructure!��Ʒ����), rsStructure!ͨ������, rsStructure!��Ʒ����)
                End If
                                                
                .TextMatrix(.Row, mconIntCol��ҩƷ���������) = rsStructure!���� & strҩ��
                .TextMatrix(.Row, mconIntCol��ҩƷ����) = rsStructure!����
                .TextMatrix(.Row, mconIntCol��ҩƷ����) = strҩ��
                
                If mintDrugNameShow = 0 Then
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ���������)
                ElseIf mintDrugNameShow = 1 Then
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ����)
                Else
                    .TextMatrix(.Row, mconIntCol��ҩ��) = .TextMatrix(.Row, mconIntCol��ҩƷ����)
                End If
                
                .TextMatrix(.Row, mconIntCol����Ʒ��) = IIf(IsNull(rsStructure!��Ʒ����), "", rsStructure!��Ʒ����)
                
                .TextMatrix(.Row, mconIntCol�����) = IIf(IsNull(rsStructure!���), "", rsStructure!���)
                .TextMatrix(.Row, mconIntCol������) = IIf(IsNull(rsStructure!����), "", rsStructure!����)
                .TextMatrix(.Row, mconIntCol����λ) = rsStructure!��λ
                .TextMatrix(.Row, mconIntCol������) = zlStr.FormatEx(rsStructure!ʵ������, mintNumberDigit, , True)
                .TextMatrix(.Row, mconIntCol���ɹ���) = zlStr.FormatEx(rsStructure!�ɱ���, mintCostDigit)
                .TextMatrix(.Row, mconIntCol���ɹ����) = zlStr.FormatEx(rsStructure!�ɱ����, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconIntCol���ۼ�) = zlStr.FormatEx(rsStructure!���ۼ�, mintPriceDigit)
                .TextMatrix(.Row, mconIntCol���ۼ۽��) = zlStr.FormatEx(rsStructure!���۽��, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintCol�����) = zlStr.FormatEx(rsStructure!���, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintcol��ҩƷid) = rsStructure!ҩƷID
                
                If .Row = .rows - 1 Then
                    .rows = .rows + 1
                End If
                .Row = .Row + 1
                rsStructure.MoveNext
            Loop
                
        End With
        rsStructure.Close
        mshStructure.Redraw = True
        Exit Function
    End If
    
    SetStructure = True
    mshStructure.Redraw = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub mshDrug_DblClick()
    mshDrug_KeyPress 13
    
End Sub

Private Sub mshDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshDrug
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshDrug_KeyPress(KeyAscii As Integer)
    With mshDrug
        If KeyAscii = 13 Then
            If Not SetColValue(mshBill.Row, .TextMatrix(.Row, 9), "[" & .TextMatrix(.Row, 2) & "]", .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), _
                .TextMatrix(.Row, 5), .TextMatrix(.Row, 6), .TextMatrix(.Row, 7), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), _
                IIf(IsNull(.TextMatrix(.Row, 15)), "0", .TextMatrix(.Row, 15)), .TextMatrix(.Row, 12), Val(.TextMatrix(.Row, 16)), _
                Val(.TextMatrix(.Row, 17)), Val(.TextMatrix(.Row, 18))) Then
                mshBill.SetFocus
                mshBill.Col = mconIntColҩ��
                .Visible = False
                Exit Sub
            End If
            .Visible = False
            mshBill.Text = "[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 4)
            mshBill.Col = mconIntCol����
            mshBill.SetFocus
        End If
    End With
End Sub

Private Sub mshDrug_LostFocus()
    SaveFlexState mshDrug, MStrCaption
    If mshDrug.Visible Then mshDrug.Visible = False
End Sub

Private Sub mshStructure_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub mshStructure_EnterCell(Row As Long, Col As Long)
    Call ��ʾ��ɿ����
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

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
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
    Dim rsDepend As New Recordset
    
    Dim lng�Ƽ��� As Long
    Dim lng������ID As Long
    Dim lng�������ID As Long
    Dim chrNo As Variant
    Dim int��� As String
    Dim lng�ⷿid As Long
    Dim lngҩƷid As Long
    Dim dbl��д���� As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim str������ As String
    Dim str�������� As String
    Dim strժҪ As String
    
    Dim intRow As Integer
    Dim n As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    
    arrSql = Array()
    
    On Error GoTo errHandle
    
    SaveCard = False
    With mshBill
        gstrSQL = "SELECT B.Id,b.ϵ��, b.���� " _
                  & " FROM ҩƷ�������� A, ҩƷ������ B " _
                  & "Where A.���id = B.ID " _
                & "AND A.���� = 3  "
        Call SQLTest(App.Title, "Э��ҩƷ������", gstrSQL)
        Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "SaveCard")
        Call SQLTest
        If rsDepend.EOF Then
            MsgBox "û������Э��ҩƷ��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        End If
        rsDepend.Filter = "ϵ��=-1"
        If rsDepend.EOF Then
            MsgBox "û������Э��ҩƷ���ĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        Else
            lng�������ID = rsDepend!Id
        End If
        rsDepend.Filter = "ϵ��=1"
        If rsDepend.EOF Then
            MsgBox "û������Э��ҩƷ��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        Else
            lng������ID = rsDepend!Id
        End If
        rsDepend.Filter = adFilterNone
        rsDepend.Close
        
        chrNo = Trim(txtNo)
        lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = Sys.GetNextNo(23, lng�ⷿid)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")

        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_Э�����_Delete('" & mstr���ݺ� & "')"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngҩƷid = .TextMatrix(intRow, 0)
                dbl��д���� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                dbl�ɱ��� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dbl�ɱ���� = .TextMatrix(intRow, mconIntCol�ɹ����)
                dbl���ۼ� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dbl���۽�� = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                dbl��� = .TextMatrix(intRow, mconintCol���)
                int��� = intRow
                              
                gstrSQL = "zl_Э�����_INSERT("
                '������ID
                gstrSQL = gstrSQL & lng������ID
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & int���
                '�ⷿID
                gstrSQL = gstrSQL & "," & lng�ⷿid
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷid
                '��д����
                gstrSQL = gstrSQL & "," & dbl��д����
                '�ɱ���
                gstrSQL = gstrSQL & "," & dbl�ɱ���
                '�ɱ����
                gstrSQL = gstrSQL & "," & dbl�ɱ����
                '���ۼ�
                gstrSQL = gstrSQL & "," & dbl���ۼ�
                '���۽��
                gstrSQL = gstrSQL & "," & dbl���۽��
                '���
                gstrSQL = gstrSQL & "," & dbl���
                '������
                gstrSQL = gstrSQL & ",'" & str������ & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS')"
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strժҪ & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gstrSQL = "zl_ҩƷЭ�����ճ���_insert('" & chrNo & "'," & lng�������ID & "," & lng�ⷿid & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        '���д�����ҩ����
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim cur���۽�� As Double, Cur�ۼ۽�� As Double, Cur��� As Double
    Dim intLop As Integer
    
    cur���۽�� = 0
    Cur�ۼ۽�� = 0
    Cur��� = 0
    
    With mshBill
        For intLop = 1 To .rows - 1
            cur���۽�� = cur���۽�� + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur�ۼ۽�� = Cur�ۼ۽�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
        Next
    End With
    
    Cur��� = Cur�ۼ۽�� - cur���۽��
    lblPurchasePrice.Caption = "���۽��ϼƣ�" & zlStr.FormatEx(cur���۽��, mintMoneyDigit, , True)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(Cur�ۼ۽��, mintMoneyDigit, , True)
    lblDifference.Caption = "��ۺϼƣ�" & zlStr.FormatEx(Cur���, mintMoneyDigit, , True)
    
End Sub


Private Sub ��ʾ�����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntColҩ��) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    
    If RecTmp.State = 1 Then RecTmp.Close
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnit = "���㵥λ"
            strQuantity = "�������� "
        Case mconint���ﵥλ
            strUnit = "���ﵥλ"
            strQuantity = "��������/�����װ "
        Case mconintסԺ��λ
            strUnit = "סԺ��λ"
            strQuantity = "��������/סԺ��װ "
        Case mconintҩ�ⵥλ
            strUnit = "ҩ�ⵥλ"
            strQuantity = "��������/ҩ���װ "
    End Select
    
    gstrSQL = " Select b.ҩƷID," & strUnit & " as ��λ, Sum(" & strQuantity & ") as ���� " & _
              " From ҩƷ��� a,ҩƷ��� b,�շ���ĿĿ¼ C " & _
              " Where a.ҩƷid=b.ҩƷid and b.ҩƷID=C.ID " & _
              " and nvl(a.��������,0)<>0 and a.����=1 And a.�ⷿID=[1] and b.ҩƷID=[2] " & _
              " Group by b.ҩƷID," & strUnit
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), intID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl���� = IIf(IsNull(RecTmp!����), 0, RecTmp!����)
    
    staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(Dbl����, mintNumberDigit, , True) & "]" & RecTmp!��λ
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ��ʾ��ɿ����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshStructure.TextMatrix(mshStructure.Row, mconIntCol��ҩ��) = "" Then
        Exit Sub
    End If
    
    intID = mshStructure.TextMatrix(mshStructure.Row, mconintcol��ҩƷid)

    gstrSQL = "Select b.ҩƷID, nvl(Sum(��������),0) as ����,c.���㵥λ as ��λ " & _
        " from ҩƷ��� a,ҩƷ��� b,�շ���ĿĿ¼ C " & _
        " Where b.ҩƷID=C.ID and b.ҩƷid=a.ҩƷid(+) AND a.�ⷿID(+)=[1] and a.����(+)=1" & _
        " and b.ҩƷID=[2] " & _
        " Group by b.ҩƷID,c.���㵥λ "
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), intID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl���� = IIf(IsNull(RecTmp!����), 0, RecTmp!����)
    
    staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & Dbl���� & "]" & RecTmp!��λ
    Exit Sub
errHandle:
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            int��λϵ�� = 4
        Case mconint���ﵥλ
            int��λϵ�� = 2
        Case mconintסԺ��λ
            int��λϵ�� = 1
        Case mconintҩ�ⵥλ
            int��λϵ�� = 3
    End Select
    strNo = txtNo.Tag
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1344", "zl8_bill_1344"), mint��¼״̬, int��λϵ��, 1344, "ҩƷЭ����ⵥ", strNo
End Sub


