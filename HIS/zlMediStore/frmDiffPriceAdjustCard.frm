VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiffPriceAdjustCard 
   Caption         =   "����۵�����"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDiffPriceAdjustCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   10
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   11
      Top             =   0
      Width           =   11715
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   5940
         TabIndex        =   31
         Top             =   945
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   1380
         TabIndex        =   1
         Top             =   615
         Width           =   2895
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "��"
         Height          =   300
         Left            =   4290
         TabIndex        =   29
         Top             =   615
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   180
         TabIndex        =   3
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
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
         TabIndex        =   5
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��λ(&G)"
         Height          =   180
         Left            =   240
         TabIndex        =   30
         Top             =   660
         Width           =   990
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "������ϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   26
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   1920
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ۺϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   22
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   19
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   4
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "����۵�����"
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
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
            Picture         =   "frmDiffPriceAdjustCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1000
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
            Picture         =   "frmDiffPriceAdjustCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   27
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
            Picture         =   "frmDiffPriceAdjustCard.frx":22EA
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
            Picture         =   "frmDiffPriceAdjustCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDiffPriceAdjustCard.frx":3080
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
      TabIndex        =   23
      Top             =   5160
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
Attribute VB_Name = "frmDiffPriceAdjustCard"
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
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mblnFirst As Boolean                '��һ����ʾ
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintҵ��ģʽ As Integer             '1-����۵���;2-�ɱ��۵���
Private mlng��ҩ��λID As Long              '��ҩ��λID

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mstrPrivs As String                     'Ȩ��

Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mlng�ⷿ As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private Const MStrCaption As String = "����۵�������"

'�Ӳ�������ȡҩƷ�۸����������С��λ�� ����
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
Private Const mconIntCol���� As Integer = 7
Private Const mconIntCol�������� As Integer = 8
Private Const mconIntCol����ϵ�� As Integer = 9
Private Const mconIntCol���� As Integer = 10
Private Const mconIntCol��λ As Integer = 11
Private Const mconIntCol���� As Integer = 12
Private Const mconIntColЧ�� As Integer = 13
Private Const mconIntCol����� As Integer = 14
Private Const mconIntCol����� As Integer = 15
Private Const mconintCol�ɱ��� As Integer = 16
Private Const mconintCol�³ɱ��� As Integer = 17
Private Const mconintCol������ As Integer = 18
Private Const mconIntColʵ������ As Integer = 19
Private Const mconIntColҩƷ��������� As Integer = 20
Private Const mconIntColҩƷ���� As Integer = 21
Private Const mconIntColҩƷ���� As Integer = 22
Private Const mconIntColS  As Integer = 23              '������

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
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
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = n
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub
Private Function CheckӦ����¼(ByVal lngҩƷID As Long, ByVal lng��ҩ��λID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
        " Where ϵͳ��ʶ=1 And ��¼����=0 And �շ�id In (Select ID From ҩƷ�շ���¼ " & _
        " Where ���� = 1 And (Mod(��¼״̬, 3) = 0 Or ��¼״̬ = 1) And ҩƷid = [1] And ��ҩ��λid = [2]) "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���Ӧ����¼]", lngҩƷID, lng��ҩ��λID)
    
    If rsCheck.EOF Then
        CheckӦ����¼ = True
        Exit Function
    Else
        CheckӦ����¼ = (rsCheck!������� = 0)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check���(ByVal lngҩƷID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    On Error GoTo errHandle
    strsql = "select Count(ҩƷid) ��� from ҩƷ��� Where ҩƷID=[1] And ����=1 And ʵ������>0 "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, MStrCaption & "[���ҩƷ���]", lngҩƷID)
    
    Check��� = (rsCheck!��� > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckͬһҩƷ(ByVal lngҩƷID As Long, ByVal intRow As Integer) As Boolean
    Dim n As Integer
        
    If intRow = 1 Then
        CheckͬһҩƷ = True
        Exit Function
    End If
    
    For n = 1 To mshBill.rows - 1
        If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
            If Val(mshBill.TextMatrix(n, 0)) = lngҩƷID And n <> intRow Then
                CheckͬһҩƷ = False
                Exit Function
            End If
        End If
    Next
    
    CheckͬһҩƷ = True
End Function
Private Function CheckҩƷ��Ӧ��(ByVal lngҩƷID As Long, ByVal lng��ҩ��λID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "Select Nvl(�ϴι�Ӧ��ID,0) �ϴι�Ӧ��ID  From ҩƷ��� Where ҩƷid=[1] And �ϴι�Ӧ��id Is Not Null Order By nvl(����,0) Desc "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, MStrCaption & "[���ҩƷ��Ӧ��]", lngҩƷID)
    If rsCheck.RecordCount = 0 Then
        CheckҩƷ��Ӧ�� = False
    Else
        CheckҩƷ��Ӧ�� = (rsCheck!�ϴι�Ӧ��ID = lng��ҩ��λID)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'=========================================================================================
'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    GetDepend = False
    strsql = "SELECT B.Id " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 5 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strsql, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ����۵���������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
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


Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False, Optional intҵ��ģʽ As Integer = 1)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mblnFirst = True
    mintҵ��ģʽ = intҵ��ģʽ
    mstrPrivs = GetPrivFunc(glngSys, 1303)
    
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
        If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    If mintҵ��ģʽ = 1 Then
        LblTitle.Caption = "����۵�����"
        LblProvider.Visible = False
        txtProvider.Visible = False
        cmdProvider.Visible = False
    Else
        LblTitle.Caption = "�ɱ��۵�����"
        LblStock.Visible = False
        txtStock.Visible = False
        If mint�༭״̬ <> 1 And mint�༭״̬ <> 2 Then
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
        End If
    End If
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(IIf(mintҵ��ģʽ = 1, 2, 1), MStrCaption, IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0))
    End If
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

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
              "Order by level,ID"
    Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "ȡҩƷ��Ӧ��", gstrNodeNo)
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "����ҩƷ��Ӧ��"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = False Then Exit Sub
        
        Me.txtProvider.Tag = .CurrentID
        Me.txtProvider = .CurrentName
    End With
    Unload FrmSelect
    mshBill.SetFocus
    
    If Val(txtProvider.Tag) <> mlng��ҩ��λID Then
        mlng��ҩ��λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol�к�) = "1"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    'mblnChange = False
    If mblnFirst = False Then Exit Sub
    
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDataBase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 1 Then
        mshBill.ClearBill
        
        Dim str��;ID As String, str���ͱ��� As String, strALL���ͱ��� As String
        Dim str���ʷ��� As String, lng�ⷿID As Long, int��۲����� As Integer
        
        If mintҵ��ģʽ = 1 Then
            If frmDiffPriceAdjustCondition.GetCondition(mfrmMain, str��;ID, str���ͱ���, lng�ⷿID, int��۲�����) = True Then
                Screen.MousePointer = 11
                SearchData str��;ID, str���ͱ���, lng�ⷿID, int��۲�����
                Screen.MousePointer = 0
            Else
                Unload Me
                Exit Sub
            End If
        Else
            Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
        End If
        
        If cmdCancel.Enabled = False Then
            cmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
        
        If txtProvider.Visible = True And mintҵ��ģʽ = 2 Then txtProvider.SetFocus
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
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
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
        mstrTime_End = GetBillInfo(5, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If

        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        If SaveCheck = True Then
            If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.��۵���)) = 1 Then
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
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.��۵���)) = 1 Then
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
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Load()
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    
    mlng�ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    'Ϊ�˴�������������ѽ��λ��Ĭ��Ϊ���λ��
'    mintMoneyDigit = gtype_UserDrugDigits.Digit_���
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "����۵�������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(5, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    If mintҵ��ģʽ = 1 Then
        mshBill.ColWidth(mconIntCol��������) = 0
        mshBill.ColWidth(mconIntCol����) = 1000
        mshBill.ColWidth(mconIntColЧ��) = 1000
        mshBill.ColWidth(mconintCol�ɱ���) = 1200
    Else
        mshBill.ColWidth(mconIntCol��������) = 1000
        mshBill.ColWidth(mconIntCol����) = 0
        mshBill.ColWidth(mconIntColЧ��) = 0
        mshBill.ColWidth(mconintCol�ɱ���) = 1200
    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim strPrice As String
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPricedigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    '�ⷿ
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.��۵���)
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
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            txtStock = .List(.ListIndex)
            txtStock.Tag = .ItemData(.ListIndex)
            
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select distinct b.id,b.���� from ҩƷ�շ���¼ a,���ű� b  " _
                    & " where a.�ⷿid=b.id and A.���� =5 and  a.no=[1]"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!����
                txtStock.Tag = rsInitCard!id
                
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д���� as ��������,'1' as ����ϵ��,"
                    strPrice = ",A.�³ɱ��� "
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��������,B.�����װ as ����ϵ��,"
                    strPrice = ",A.�³ɱ���*B.�����װ AS �³ɱ��� "
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��������,B.סԺ��װ as ����ϵ��,"
                    strPrice = ",A.�³ɱ���*B.סԺ��װ AS �³ɱ��� "
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��������,B.ҩ���װ as ����ϵ��,"
                    strPrice = ",A.�³ɱ���*B.ҩ���װ AS �³ɱ��� "
            End Select
            If mintҵ��ģʽ = 1 Then
                gstrSQL = "SELECT * " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    "     B.ҩƷ��Դ,B.����ҩ��,F.���,A.����, A.����,A.Ч��,A.����," & _
                    "     NVL(E.����,F.����) ����," & strUnitQuantity & _
                    "     A.�ɱ��� AS �����,NVL(A.���ۼ�,0) AS �����,A.��� AS ������, " & _
                    "     A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.��д���� ʵ������,A.���� As �³ɱ��� " & _
                    "     FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    "     AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    "     AND A.��¼״̬ =[2] AND A.���� =5 AND A.NO = [1]) " & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT a.*,Rownum ��� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    "     B.ҩƷ��Դ,B.����ҩ��,F.���," & _
                    "     NVL(E.����,F.����) ����," & strUnitQuantity & _
                    "     A.�ɱ��� AS �����,NVL(A.���ۼ�,0) AS �����,A.��� AS ������, " & _
                    "     A.ժҪ,������,��������,�����,�������,A.��д���� ʵ������ " & strPrice & ",G.���� ��Ӧ��,G.Id ��ҩ��λid" & _
                    "     FROM (Select Sum(��д����) ��д����, ҩƷid, Sum(�ɱ���) �ɱ���, Nvl(Sum(���ۼ�), 0) ���ۼ�, Sum(���) ���, ժҪ, ������," & _
                    "     �������� , �����, �������,���� �³ɱ���,��ҩ��λid" & _
                    "     From ҩƷ�շ���¼ " & _
                    "     Where ���� = 5 And No = [1] And ��¼״̬ = [2] " & _
                    "     Group By ҩƷid, ժҪ, ������, ��������, �����, �������,����,��ҩ��λid) A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,��Ӧ�� G " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    "     AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 And A.��ҩ��λid=G.Id ) A " & _
                    "  ORDER BY " & strSqlOrder
            End If
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û�����
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            With mshBill
                If mintҵ��ģʽ = 2 Then
                    txtProvider.Text = rsInitCard!��Ӧ��
                    mlng��ҩ��λID = rsInitCard!��ҩ��λID
                End If
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
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
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(rsInitCard!�����, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(IIf(IsNull(rsInitCard!�����), 0, rsInitCard!�����), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(rsInitCard!������, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������), intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(IIf(IsNull(rsInitCard!ʵ������), "0", rsInitCard!ʵ������), intNumberDigit, , True)
                    If mintҵ��ģʽ = 1 Then
                        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                        .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                        If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                            '����Ϊ��Ч��
                            .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                        End If
                        If Not IsNull(rsInitCard!�³ɱ���) Then
                            .TextMatrix(intRow, mconintCol�³ɱ���) = zlStr.FormatEx(rsInitCard!�³ɱ��� * rsInitCard!����ϵ��, intCostDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx((rsInitCard!����� - rsInitCard!�����) / rsInitCard!��������, intCostDigit, , True)
                        If Not IsNull(rsInitCard!�³ɱ���) Then
                            .TextMatrix(intRow, mconintCol�³ɱ���) = zlStr.FormatEx(rsInitCard!�³ɱ��� * rsInitCard!����ϵ��, intCostDigit, , True)
                        End If
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconintCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconintCol�³ɱ���) = "�³ɱ���"
        .TextMatrix(0, mconIntColʵ������) = "ʵ������"
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
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntColҩ��) = 2500
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol���) = 1000
        .ColWidth(mconIntCol����) = 1000
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol�����) = 1200
        .ColWidth(mconIntCol�����) = 1200
        .ColWidth(mconintCol������) = 1200
        .ColWidth(mconintCol�³ɱ���) = 1200
        .ColWidth(mconIntColʵ������) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        
        If mintҵ��ģʽ = 1 Then
            .ColWidth(mconIntCol��������) = 0
            .ColWidth(mconIntCol����) = 1000
            .ColWidth(mconIntColЧ��) = 1000
            .ColWidth(mconintCol�ɱ���) = 1200
        Else
            .ColWidth(mconIntCol��������) = 1000
            .ColWidth(mconIntCol����) = 0
            .ColWidth(mconIntColЧ��) = 0
            .ColWidth(mconintCol�ɱ���) = 1200
        End If
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol�����) = 5
        .ColData(mconIntCol�����) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntColʵ������) = 5
        .ColData(mconintCol�ɱ���) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mconIntColҩ��) = 1
            .ColData(mconintCol�³ɱ���) = 4
            If mintҵ��ģʽ = 1 Then
                .ColData(mconintCol������) = 4
            Else
                .ColData(mconintCol������) = 5
            End If
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            .ColData(mconintCol������) = 5
            .ColData(mconintCol�³ɱ���) = 5
        End If
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�����) = flexAlignRightCenter
        .ColAlignment(mconintCol������) = flexAlignRightCenter
        .ColAlignment(mconintCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconintCol�³ɱ���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
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
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - cmdCancel.Height - 200
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
        LblNO.Left = .Left - LblNO.Width - 100
        .Top = LblTitle.Top
        LblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
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
        '.Width = .Left - .Left
        Debug.Print .Width
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
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
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With cmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = cmdCancel.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = cmdCancel.Top
    End With
        
    With cmdFind
        .Top = cmdCancel.Top
    End With
    
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\����۵�������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelLength = Len(txtProvider.Text)
        txtProvider.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
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

Private Function SaveCheck() As Boolean
    Dim strNo As String
    Dim str����� As String
    Dim n As Integer
    
    mblnSave = False
    SaveCheck = False
    
    str����� = UserInfo.�û�����
    strNo = txtNo.Tag
    On Error GoTo errHandle
    
    '���Ӧ����¼������Ѿ�������ܵ����ɱ���
    If mintҵ��ģʽ = 2 Then
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
                If Not CheckӦ����¼(Val(mshBill.TextMatrix(n, 0)), mlng��ҩ��λID) Then
                    MsgBox mshBill.TextMatrix(n, mconIntColҩ��) & " ��ȫ������ܵ����ɱ��ۣ�", vbInformation + vbOKOnly, gstrSysName
                    mshBill.SetFocus
                    mshBill.Col = mconIntColҩ��
                    Exit Function
                End If
            End If
        Next
    End If
                            
    gstrSQL = "zl_ҩƷ����۵���_Verify('" & strNo & "','" & str����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
   
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "���ʧ�ܣ�", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
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
End Sub
Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntColҩ��) = 0 Then
        'Cancel = True    '�ȴ���CANCEL����
        Exit Sub
    End If
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
    Dim RecReturn As Recordset
    Dim strҩ�� As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, IIf(mintҵ��ģʽ = 1, 2, 1), IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), , , False)
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(IIf(mintҵ��ģʽ = 1, 2, 1), MStrCaption, IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0))
    End If
    Set RecReturn = frmSelector.showMe(Me, 0, IIf(mintҵ��ģʽ = 1, 2, 1), , , , IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), , , False, , , , , , mstrPrivs & ";�鿴�ɱ���;")
    mshBill.CmdEnable = True
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
'        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
'            strҩ�� = RecReturn!ͨ����
'        Else
'            strҩ�� = IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
'        End If
            
        If mintҵ��ģʽ = 2 Then
            '���ҩƷ�ظ�
'            If Not CheckͬһҩƷ(RecReturn!ҩƷID, mshBill.Row) Then
'                MsgBox "ҩƷ" & strҩ�� & "�Ѵ��ڣ����������룡", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntColҩ��
'                Exit Sub
'            End If
'
'            '���ҩƷ���
'            If Not Check���(RecReturn!ҩƷID) Then
'                MsgBox "ҩƷ" & strҩ�� & " �����пⷿ���޿�棬���ܵ����ɱ��ۣ�", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntColҩ��
'                Exit Sub
'            End If
'
'            '���ҩƷ��Ӧ�̹�ϵ
'            If Not CheckҩƷ��Ӧ��(RecReturn!ҩƷID, mlng��ҩ��λID) Then
'                MsgBox txtProvider.Text & "����ҩƷ" & strҩ�� & " �Ĺ�ҩ��λ��������ѡ��ҩƷ���߹�ҩ��λ��", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntColҩ��
'                Exit Sub
'            End If
'
'            '���Ӧ����¼������Ѿ�������ܵ����ɱ���
'            If Not CheckӦ����¼(RecReturn!ҩƷID, mlng��ҩ��λID) Then
'                MsgBox strҩ�� & " ��ȫ������ܵ����ɱ��ۣ�", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntColҩ��
'                Exit Sub
'            End If
            Set RecReturn = CheckData(RecReturn)
        End If
        If RecReturn.RecordCount > 0 Then
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                intRow = mshBill.Row
                With mshBill
                    .TextMatrix(intRow, mconIntCol�к�) = .Row
                    SetColValue .Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", _
                        RecReturn!��Ʒ��), Nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                        IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                        Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                        IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                        IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                        IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
                        IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
                        Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                        IIf(IsNull(RecReturn!�������), "0", RecReturn!�������)
                    
                    .Col = mconintCol�³ɱ���
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                    RecReturn.MoveNext
                End With
            Next
            mshBill.Row = intOldRow
            RecReturn.Close
        End If
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
        If strkey = "" Then
            strkey = .TextMatrix(.Row, .Col)
        End If
        Select Case .Col
            Case mconintCol�³ɱ���
               intDigit = mintCostDigit
            Case mconintCol������
                intDigit = mintMoneyDigit
        End Select
        
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If .SelLength = Len(strkey) Then Exit Sub
            If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        Select Case .Col
            Case mconIntColҩ��
                .TxtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
            Case mconintCol������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890-"
            Case mconintCol�ɱ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconintCol�³ɱ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strҩ�� As String
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntColҩ��
                If strkey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    Dim i As Integer
                    Dim intCurRow As Integer
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, IIf(mintҵ��ģʽ = 1, 2, 1), IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), , , strkey, sngLeft, sngTop, False)
                    
                    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                        Call SetSelectorRS(IIf(mintҵ��ģʽ = 1, 2, 1), MStrCaption, IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0))
                    End If
                    Set RecReturn = frmSelector.showMe(Me, 1, IIf(mintҵ��ģʽ = 1, 2, 1), strkey, sngLeft, sngTop, IIf(mintҵ��ģʽ = 1, txtStock.Tag, 0), , , False, , , , , , mstrPrivs & ";�鿴�ɱ���;")
'                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
'                        strҩ�� = RecReturn!ͨ����
'                    Else
'                        strҩ�� = IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
'                    End If
            
                    If mintҵ��ģʽ = 2 Then
                    '���ҩƷ�ظ�
'                        If Not CheckͬһҩƷ(RecReturn!ҩƷID, mshBill.Row) Then
'                            MsgBox "ҩƷ" & strҩ�� & "�Ѵ��ڣ����������룡", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntColҩ��
'                            Cancel = True
'                            Exit Sub
'                        End If
'
'                        '���ҩƷ���
'                        If Not Check���(RecReturn!ҩƷID) Then
'                            MsgBox "ҩƷ" & strҩ�� & " �����пⷿ���޿�棬���ܵ����ɱ��ۣ�", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            mshBill.Col = mconIntColҩ��
'                            Exit Sub
'                        End If
'
'                        '���ҩƷ��Ӧ�̹�ϵ
'                        If Not CheckҩƷ��Ӧ��(RecReturn!ҩƷID, mlng��ҩ��λID) Then
'                            MsgBox txtProvider.Text & "����ҩƷ" & strҩ�� & " �Ĺ�ҩ��λ��������ѡ��ҩƷ���߹�ҩ��λ��", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntColҩ��
'                            Exit Sub
'                        End If
'
'                        '���Ӧ����¼������Ѿ�������ܵ����ɱ���
'                        If Not CheckӦ����¼(RecReturn!ҩƷID, mlng��ҩ��λID) Then
'                            MsgBox strҩ�� & " ��ȫ������ܵ����ɱ��ۣ�", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntColҩ��
'                            Exit Sub
'                        End If
                        If RecReturn.RecordCount > 0 Then
                            Set RecReturn = CheckData(RecReturn)
                        End If
                    End If
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                    Nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                    IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                    IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
                                    IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
                                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), IIf(IsNull(RecReturn!�������), "0", RecReturn!�������)) = False Then
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            
                            Call ��ʾ�����
                        
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        Cancel = True
                    End If
                End If
            Case mconintCol�³ɱ���
                If strkey = "" And mintҵ��ģʽ = 1 Then
                    .Col = mconintCol������
                    Cancel = True
                    Exit Sub
                End If
                
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬳ɱ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) < 0.001 Then
                        MsgBox "�Բ��𣬳ɱ��۱������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "�ɱ��۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strkey, mintCostDigit, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
      
                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintCostDigit, , True)
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol�³ɱ���) = .Text
                End If
                                
                '�����۵�����(�����������������*�ɱ���-�����)
                If strkey <> "" Then
                    .TextMatrix(.Row, mconintCol������) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�����) = "", 0, .TextMatrix(.Row, mconIntCol�����)) - Val(IIf(.TextMatrix(.Row, mconIntCol��������) = "", 0, .TextMatrix(.Row, mconIntCol��������))) * Val(IIf(.TextMatrix(.Row, mconintCol�³ɱ���) = "", 0, .TextMatrix(.Row, mconintCol�³ɱ���))) _
                        - Val(IIf(.TextMatrix(.Row, mconIntCol�����) = "", 0, .TextMatrix(.Row, mconIntCol�����))), mintMoneyDigit, , True)
                End If
                
            Case mconintCol������
                If .TextMatrix(.Row, .Col) = "" And strkey = "" Then
                    MsgBox "�Բ��𣬵�����������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬵��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) = 0 Then
                        MsgBox "�Բ��𣬵������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strkey)) < 0.00001 Then
                        MsgBox "�Բ��𣬵�����ľ���ֵ���벻С��0.00001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "���������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = zlStr.FormatEx(strkey, mintMoneyDigit, , True)
                    .Text = strkey
                    
                    '����ɱ���(�ɱ���=(�����-�����-������)/��������)
                    If strkey <> "" And Val(.TextMatrix(.Row, mconIntCol��������)) <> 0 Then
                        .TextMatrix(.Row, mconintCol�³ɱ���) = zlStr.FormatEx((IIf(.TextMatrix(.Row, mconIntCol�����) = "", 0, .TextMatrix(.Row, mconIntCol�����)) - Val(IIf(.TextMatrix(.Row, mconIntCol�����) = "", 0, .TextMatrix(.Row, mconIntCol�����))) - Val(strkey)) / Val(IIf(.TextMatrix(.Row, mconIntCol��������) = "", 0, .TextMatrix(.Row, mconIntCol��������))), mintCostDigit, , True)
                    End If
                End If
                Call ��ʾ�ϼƽ��
        End Select
    End With
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal intҩƷid As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, _
    ByVal str����ҩ�� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal str���� As String, ByVal strЧ�� As String, _
    ByVal num����� As Double, ByVal lng���� As Long, ByVal num�������� As Double, _
    ByVal num����ϵ�� As Double, ByVal num����� As Double, ByVal numʵ������ As Double) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim strҩ�� As String
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
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
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(num�����, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(num�����, mintMoneyDigit, , True)
        
        If mintҵ��ģʽ = 1 Then
            If lng���� > 0 Then
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num�������� / num����ϵ��, mintNumberDigit, , True)
            End If
            .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Get�ɱ���(intҩƷid, txtStock.Tag, Val(lng����)) * num����ϵ��, mintCostDigit, , True)
        Else
            .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num�������� / num����ϵ��, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx((((num����� - num�����)) / num��������) / num����ϵ��, mintCostDigit, , True)
        End If
        .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��������)
    End With
    SetColValue = True
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

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub


Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelStart = 0
        txtProvider.SelLength = Len(txtProvider.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtProvider.Text = mshProvider.TextMatrix(mshProvider.Row, 2)
        txtProvider.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        mshBill.SetFocus
    End If

    If Val(txtProvider.Tag) <> mlng��ҩ��λID Then
        mshBill.ClearBill
        mlng��ҩ��λID = Val(txtProvider.Tag)
        mshBill.TextMatrix(1, mconIntCol�к�) = "1"
    End If
End Sub


Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
    End If
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
            If mintҵ��ģʽ = 2 Then
                If mlng��ҩ��λID = 0 Then
                    MsgBox "�Բ��𣬹�ҩ��λ����Ϊ�գ�", vbOKOnly + vbInformation, gstrSysName
                    txtProvider.SetFocus
                    Exit Function
                End If
            End If
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If mintҵ��ģʽ = 2 Then
                    If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                        If Trim(.TextMatrix(intLop, mconintCol�³ɱ���)) = "" Then
                            MsgBox "�Բ����³ɱ��۲���Ϊ�գ�", vbOKOnly + vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconintCol�³ɱ���
                            Exit Function
                        End If
                    End If
                End If
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol������))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol������)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĵ�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol������
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
    Dim lng������id As Long
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim str���� As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dbl�������� As Double
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim dbl������ As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim dat�������� As String
    Dim rs������ As New Recordset
    Dim dbl�³ɱ��� As Double
    
    Dim intRow As Integer
    Dim n As Integer
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveCard = False
    arrSql = Array()
    On Error GoTo errHandle
    '����������������ID����Ҫ������ҩƷ��Ҫ����
    gstrSQL = "SELECT B.Id " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 5 "
    Set rs������ = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
    
    If rs������.EOF Then
        MsgBox "û������ҩƷ����۵���������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rs������.Close
        Exit Function
    End If
    lng������id = rs������.Fields(0)
    rs������.Close
   
    With mshBill
        chrNo = Trim(txtNo)
        lng�ⷿID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(25, lng�ⷿID)
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        dat�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_ҩƷ����۵���_Delete('" & mstr���ݺ� & "')"
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
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng����ID = Val(.TextMatrix(intRow, mconIntCol����))
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                
                dbl�������� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColʵ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_����, , True)
                dbl����� = .TextMatrix(intRow, mconIntCol�����)
                dbl����� = .TextMatrix(intRow, mconIntCol�����)
                dbl������ = .TextMatrix(intRow, mconintCol������)
                dbl�³ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�³ɱ���)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
                lng��� = intRow
                
                'zl_ҩƷ����۵���_INSERT( /*������ID_IN*/, /*NO_IN*/, /*���_IN*/,
                    '/*�ⷿID_IN*/, /*ҩƷID_IN*/, /*����_IN*/, /*��������_IN*/,
                    '/*�����_IN*/, /*������_IN*/, /*������_IN*/, /*��������_IN*/,
                    '/*����_IN*/, /*����_IN*/, /*Ч��_IN*/, /*ժҪ_IN*/ );
                    
                gstrSQL = "zl_ҩƷ����۵���_INSERT(" & lng������id & ",'" & chrNo & "'," & lng��� & "," _
                    & lng�ⷿID & "," & lngҩƷID & "," & lng����ID & "," & dbl�������� & "," _
                    & dbl����� & "," & dbl����� & "," & dbl������ & ",'" & str������ & "',to_date('" & dat�������� & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & str���� & "','" & str���� & "'," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & strժҪ & "'," & mlng��ҩ��λID & "," & dbl�³ɱ��� & "," & IIf(mintҵ��ģʽ = 1, 0, 1) & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim dbl����� As Double
    Dim dbl������ As Double
    Dim dbl����� As Double
    
    Dim intLop As Integer
    
    dbl����� = 0
    dbl������ = 0
    
    With mshBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl����� = dbl����� + Val(.TextMatrix(intLop, mconIntCol�����))
                dbl����� = dbl����� + Val(.TextMatrix(intLop, mconIntCol�����))
                dbl������ = dbl������ + Val(.TextMatrix(intLop, mconintCol������))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "�����ϼƣ�" & zlStr.FormatEx(dbl�����, mintMoneyDigit, , True)
    lblSalePrice.Caption = "����ۺϼƣ�" & zlStr.FormatEx(dbl�����, mintMoneyDigit, , True)
    lblDifference.Caption = "������ϼƣ�" & zlStr.FormatEx(dbl������, mintMoneyDigit, , True)
    
End Sub

Private Sub ��ʾ�����()
    
    If mint�༭״̬ = 4 Then Exit Sub
    With mshBill
        If .TextMatrix(.Row, mconIntColҩ��) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol��������), mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ)
    End With
End Sub

Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    mblnChange = True
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub


Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [2] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                  "  And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] Or ���� like [1] or ���� like [1] )"
        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If adoProvider.EOF Then
            MsgBox "û��������Ĺ�ҩ��λ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1000
                .ColWidth(2) = 2700
                .ColWidth(3) = 1200
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Top = txtProvider.Top + txtProvider.Height
                .Left = cmdProvider.Left + cmdProvider.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!����
            .Tag = adoProvider!id
        End If
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If Val(.Tag) <> mlng��ҩ��λID Then
            mlng��ҩ��λID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mconIntCol�к�) = "1"
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub


Private Sub txtProvider_Validate(Cancel As Boolean)
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng��ҩ��λID Then
        mlng��ҩ��λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol�к�) = "1"
    End If
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
    FrmBillPrint.showMe Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1303", "zl8_bill_1303"), mint��¼״̬, int��λϵ��, 1303, "ҩƷ��۵�����", strNo
End Sub

Private Sub SearchData(ByVal str��;ID, ByVal str���ͱ��� As String, _
    ByVal lng�ⷿID As Long, ByVal intRate As Integer)
    
    Dim rsData As New Recordset  'ҩƷ����¼��
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim intRecordCount As Integer
    Dim strUnitQuantity As String
    Dim strҩ�� As String
    Dim strUseID As String, strClassID As String
    
    On Error GoTo errHandle:
    '���ý�����ʾ����
    staThis.Panels(2).Text = "���ڶ�" & txtStock & "��ҩƷ�����Զ���ۼ���"
    '����ҩƷ��ѯ����(ҩƷĿ¼ A)
    strPhysic = " And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)"
    If str���ͱ��� = "" Then str���ͱ��� = "'ZYB'"
    
    If str��;ID <> "" Then
        If InStr(1, "'�г�ҩ','�в�ҩ','����ҩ'", str��;ID) <> 0 Then
            Select Case str��;ID
            Case "'����ҩ'"
                strClassID = "1"
            Case "'�г�ҩ'"
                strClassID = "2"
            Case Else
                strClassID = "3"
            End Select
            strPhysic = strPhysic & " And F.���� = [5] "
        Else
            strUseID = str��;ID
            strPhysic = strPhysic & " And M.����ID in (select * from Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And F.���� In ('1','2','3') "     '���������� In δ���Ż�����
        End If
    End If
    
    DoEvents    ': Me.Refresh

    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnitQuantity = "C.���㵥λ AS ��λ, nvl(b.ʵ������,0) AS ��������, '1' as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���,b.ƽ���ɱ���) �ɱ���,"
        Case mconint���ﵥλ
            strUnitQuantity = "a.���ﵥλ AS ��λ,(nvl(b.ʵ������,0)/a.�����װ) AS ��������,a.�����װ as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���*a.�����װ,b.ƽ���ɱ���*a.�����װ) �ɱ���,"
        Case mconintסԺ��λ
            strUnitQuantity = "a.סԺ��λ AS ��λ, (nvl(b.ʵ������,0)/a.סԺ��װ) AS ��������, a.סԺ��װ as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���*a.סԺ��װ,b.ƽ���ɱ���*a.סԺ��װ) �ɱ���,"
        Case mconintҩ�ⵥλ
            strUnitQuantity = "a.ҩ�ⵥλ AS ��λ, (nvl(b.ʵ������,0)/a.ҩ���װ) AS ��������,a.ҩ���װ as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���*a.ҩ���װ,b.ƽ���ɱ���*a.ҩ���װ) �ɱ���,"
    End Select

    gstrSQL = "SELECT DISTINCT B.ҩƷID,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, D.���� As ��Ʒ��," & _
        " A.ҩƷ��Դ,A.����ҩ��,C.���,NVL(B.�ϴβ���,C.����) AS ����,B.����,B.�ϴ����� AS ����, B.Ч��," & _
        " B.ʵ�ʽ��, B.ʵ�ʲ��," & strUnitQuantity & _
        " DECODE(SIGN (B.ʵ�ʲ��/B.ʵ�ʽ��*100-(A.ָ�������+[3])),1,-(ʵ�ʲ��-B.ʵ�ʽ��*A.ָ�������/100)," & _
        " DECODE(SIGN(B.ʵ�ʲ��/B.ʵ�ʽ��*100-(A.ָ�������-[3])),-1,B.ʵ�ʽ��*A.ָ�������/100-ʵ�ʲ��)) AS ��۵�����,NVL(b.ʵ������,0) ʵ������ " & _
        " FROM ҩƷ��� A,(SELECT �ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�, ���ۼ�, �ϴο���,ƽ���ɱ��� FROM ҩƷ��� WHERE NVL(ʵ�ʽ��,0)<>0) B," & _
        " �շ���ĿĿ¼ C,�շ���Ŀ���� D,ҩƷ���� T,���Ʒ���Ŀ¼ F,������ĿĿ¼ M"
    
    gstrSQL = gstrSQL & " WHERE A.ҩƷID = C.ID and A.ҩ��ID=T.ҩ��ID " & _
        " And T.ҩ��ID=M.ID And M.����ID=F.ID " & _
        " AND A.ҩƷID=D.�շ�ϸĿID(+) AND D.����(+)=3 AND D.����(+)=1 " & _
        " AND B.����=1 AND B.�ⷿID=[1] AND A.ҩƷID=B.ҩƷID " & _
        " AND (B.ʵ�ʲ��/NVL(B.ʵ�ʽ��,1)*100>(A.ָ�������+[3]) OR B.ʵ�ʲ��/NVL(B.ʵ�ʽ��,1)*100<A.ָ�������-[3])" & strPhysic
        
    If str���ͱ��� <> "" Then
        gstrSQL = gstrSQL & " And T.ҩƷ���� in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) "
    End If
    
    gstrSQL = gstrSQL & " ORDER BY ҩƷ����"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���ڼ���ҩƷ�������]", lng�ⷿID, str���ͱ���, intRate, strUseID, strClassID)
    
    intRecordCount = rsData.RecordCount
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    If intRecordCount = 0 Then
        MsgBox "δ����ȷ��ȡҩƷ�������,�����Ի��ֹ�����ҩƷ��", vbInformation, gstrSysName: Exit Sub
    End If
    
    DoEvents: 'Me.Refresh
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .rows = .rows + 1
            .TextMatrix(i, 0) = rsData!ҩƷid
           
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                strҩ�� = rsData!ͨ����
            Else
                strҩ�� = IIf(IsNull(rsData!��Ʒ��), rsData!ͨ����, rsData!��Ʒ��)
            End If
            
            .TextMatrix(i, mconIntColҩƷ���������) = rsData!ҩƷ���� & strҩ��
            .TextMatrix(i, mconIntColҩƷ����) = rsData!ҩƷ����
            .TextMatrix(i, mconIntColҩƷ����) = strҩ��
            
            If mintDrugNameShow = 1 Then
                .TextMatrix(i, mconIntColҩ��) = .TextMatrix(i, mconIntColҩƷ����)
            ElseIf mintDrugNameShow = 2 Then
                .TextMatrix(i, mconIntColҩ��) = .TextMatrix(i, mconIntColҩƷ����)
            Else
                .TextMatrix(i, mconIntColҩ��) = .TextMatrix(i, mconIntColҩƷ���������)
            End If
            
            .TextMatrix(i, mconIntCol��Ʒ��) = IIf(IsNull(rsData!��Ʒ��), "", rsData!��Ʒ��)
            .TextMatrix(i, mconIntCol��Դ) = IIf(IsNull(rsData!ҩƷ��Դ), "", rsData!ҩƷ��Դ)
            .TextMatrix(i, mconIntCol����ҩ��) = IIf(IsNull(rsData!����ҩ��), "", rsData!����ҩ��)

            .TextMatrix(i, mconIntCol���) = IIf(IsNull(rsData!���), "", rsData!���)
            .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(i, mconIntCol��λ) = IIf(IsNull(rsData!��λ), "", rsData!��λ)
            .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsData!����), "0", rsData!����)
            .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(i, mconIntColЧ��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(i, mconIntColЧ��) <> "" Then
                '����Ϊ��Ч��
                .TextMatrix(i, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(i, mconIntColЧ��)), "yyyy-mm-dd")
            End If
           
            .TextMatrix(i, mconIntCol��������) = rsData!��������
            .TextMatrix(i, mconIntColʵ������) = rsData!ʵ������ / rsData!����ϵ��
            .TextMatrix(i, mconIntCol�����) = zlStr.FormatEx(rsData!ʵ�ʽ��, mintMoneyDigit, , True)
            .TextMatrix(i, mconIntCol�����) = zlStr.FormatEx(rsData!ʵ�ʲ��, mintMoneyDigit, , True)
            .TextMatrix(i, mconintCol������) = zlStr.FormatEx(rsData!��۵�����, mintMoneyDigit, , True)
            .TextMatrix(i, mconIntCol����ϵ��) = rsData!����ϵ��
            .TextMatrix(i, mconintCol�ɱ���) = zlStr.FormatEx(rsData!�ɱ���, mintCostDigit, , True)
                
            Call zlControl.StaShowPercent(i / intRecordCount, staThis.Panels(2), frmDiffPriceAdjustCard)
            i = i + 1
            rsData.MoveNext
        Loop
        .Redraw = True
    End With
    rsData.Close
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    
    staThis.Panels(2).Text = ""
    mshBill.Row = 1
    mshBill.Col = mconintCol������
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    mshBill.Redraw = True
    Call SaveErrLog
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str��� As String
    Dim strsql As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim strNotNum As String  '�޿��ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strNotҩ�� As String    '������¼��ЩҩƷ��ʱ�۵��޿��
    Dim bln��Ӧ�� As Boolean    '��֤��ҩƷ�Ƿ��ѡ��Ĺ�Ӧ����ͬ
    Dim str��Ӧ�� As String
    Dim strPro As String
    Dim strProvider As String
    Dim bln�Ƿ񸶿� As Boolean
    Dim str�Ƿ񸶿� As String
    Dim strPay As String
    Dim strToP As String
    Dim strmsg��Ӧ�� As String
    Dim strmsg�Ƿ��ظ� As String
    Dim strmsg��� As String
    Dim strmsg�Ƿ񸶿� As String
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str���� = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
        
        If rsTemp!ʱ�� = 1 Then '��ʱ���޿��ļ�¼�ҳ���
            gstrSQL = "select Decode(Nvl(����,0),0,ʵ�ʽ��/ʵ������,Nvl(���ۼ�,ʵ�ʽ��/ʵ������))*" & Choose(mintUnit, 1, rsTemp!�����װ, rsTemp!סԺ��װ, rsTemp!ҩ���װ) & " as  �ۼ� " _
                & "  from ҩƷ��� " _
                & " where �ⷿid=[1] " _
                & " and ҩƷid=[2] " _
                & " and ����=1 and ʵ������>0 and " _
                & " nvl(����,0)=[3]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!ҩƷid, IIf(IsNull(rsTemp!����), 0, rsTemp!����))
            If rsPrice.EOF Then
                str��� = str��� & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
            End If
        End If
        
        bln��Ӧ�� = CheckҩƷ��Ӧ��(rsTemp!ҩƷid, mlng��ҩ��λID)  '���ҩƷ�Ĺ�Ӧ��
        If bln��Ӧ�� = False Then
            str��Ӧ�� = str��Ӧ�� & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
        End If
        
        bln�Ƿ񸶿� = CheckӦ����¼(rsTemp!ҩƷid, mlng��ҩ��λID)  '����Ƿ񸶿�
        If bln�Ƿ񸶿� = False Then
            str�Ƿ񸶿� = str�Ƿ񸶿� & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
        End If
        
        rsTemp.MoveNext
    Loop
        
    With mshBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
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
        If str��� <> "" Then
            strNotNum = ""
            For i = 0 To UBound(Split(str���, "|")) - 1
                strNotNum = strNotNum & "ҩƷid<>" & Split(Split(str���, "|")(i), ",")(0) & " and "
                If UBound(Split(strNotҩ��, ",")) <= 2 Then
                    strNotҩ�� = strNotҩ�� & Split(Split(str���, "|")(i), ",")(1) & ","
                End If
            Next
            If strNotNum <> "" Then
                strNotNum = Mid(strNotNum, 1, Len(strNotNum) - 4)
            End If
        End If
        If str��Ӧ�� <> "" Then
            strProvider = ""
            For i = 0 To UBound(Split(str��Ӧ��, "|")) - 1
                strProvider = strProvider & "ҩƷid<>" & Split(Split(str��Ӧ��, "|")(i), ",")(0) & " and "
                If UBound(Split(strPro, ",")) <= 2 Then
                    strPro = strPro & Split(Split(str��Ӧ��, "|")(i), ",")(1) & ","
                End If
            Next
            If strProvider <> "" Then
                strProvider = Mid(strProvider, 1, Len(strProvider) - 4)
            End If
        End If
        If str�Ƿ񸶿� <> "" Then
            strProvider = ""
            For i = 0 To UBound(Split(str�Ƿ񸶿�, "|")) - 1
                strPay = strPay & "ҩƷid<>" & Split(Split(str�Ƿ񸶿�, "|")(i), ",")(0) & " and "
                If UBound(Split(strToP, ",")) <= 2 Then
                    strToP = strToP & Split(Split(str�Ƿ񸶿�, "|")(i), ",")(1) & ","
                End If
            Next
            If strPay <> "" Then
                strPay = Mid(strPay, 1, Len(strPay) - 4)
            End If
        End If
        
        
        '�ж���ʲô��ʽƴ��sql
        strsql = strDub & " " & strNotNum & " " & strProvider & " " & strPay
        If str�ظ�ҩ�� <> "" Then
            strmsg�Ƿ��ظ� = str�ظ�ҩ�� & "�б����Ѿ������ˣ�"
            strsql = strDub
        End If
        If strNotҩ�� <> "" Then
            strmsg��� = vbCrLf & strNotҩ�� & "��ʱ��ҩƷ��û�п�治������⣡"
            If strsql = "" Then
                strsql = strNotNum
            Else
                strsql = strsql & " and " & strNotNum
            End If
        End If
        If strPro <> "" Then
            strmsg��Ӧ�� = vbCrLf & strPro & "��ʱ��ҩƷ��û�п�治������⣡"
            If strsql = "" Then
                strsql = strProvider
            Else
                strsql = strsql & " and " & strProvider
            End If
        End If
        If strToP <> "" Then
            strmsg�Ƿ񸶿� = vbCrLf & strToP & "��ʱ��ҩƷ��û�п�治������⣡"
            If strsql = "" Then
                strsql = strPay
            Else
                strsql = strsql & " and " & strPay
            End If
        End If
        If strmsg�Ƿ��ظ� <> "" Or strmsg��� <> "" Or strmsg��Ӧ�� <> "" Or strmsg�Ƿ񸶿� <> "" Then
            MsgBox strmsg�Ƿ��ظ� & strmsg��� & strmsg��Ӧ�� & strmsg�Ƿ񸶿� & "...����ҩƷ��������ӣ�", vbInformation, gstrSysName
        End If
        
        If strsql <> "" Then
            rsTemp.Filter = strsql
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

