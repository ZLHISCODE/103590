VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷƷ�ֱ༭"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmMediItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk������ҩ 
      Caption         =   "������ҩ"
      Height          =   210
      Left            =   6750
      TabIndex        =   30
      Top             =   4560
      Width           =   1050
   End
   Begin VB.ComboBox cbo������ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox chk������ 
      Caption         =   "����ҩ��(&Q)"
      Height          =   270
      Left            =   5265
      TabIndex        =   32
      Top             =   5175
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6870
      TabIndex        =   38
      Top             =   6225
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmMediItem.frx":058A
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6225
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�����˳�(&O)"
      Height          =   350
      Left            =   5385
      TabIndex        =   35
      Top             =   6225
      Width           =   1215
   End
   Begin VB.CheckBox chkԭ��ҩ 
      Caption         =   "ԭ��ҩ(&M)"
      Height          =   210
      Left            =   5265
      TabIndex        =   22
      Top             =   3690
      Width           =   1155
   End
   Begin VB.CheckBox chk��ҩ 
      Caption         =   "��ҩ(&W)"
      Height          =   210
      Left            =   6750
      TabIndex        =   26
      Top             =   3390
      Width           =   1155
   End
   Begin VB.CheckBox chkƤ�� 
      Caption         =   "Ƥ��(&Y)"
      Height          =   210
      Left            =   6750
      TabIndex        =   27
      Top             =   3690
      Width           =   1155
   End
   Begin VB.CheckBox chk����ҩ 
      Caption         =   "����ҩ(&J)"
      Height          =   210
      Left            =   5265
      TabIndex        =   21
      Top             =   3390
      Width           =   1155
   End
   Begin VB.TextBox txtӢ�� 
      Height          =   300
      Left            =   1230
      TabIndex        =   6
      Top             =   1575
      Width           =   3675
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   3135
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1200
      Width           =   1170
   End
   Begin VB.ComboBox cboҩƷ���� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1575
      Width           =   1455
   End
   Begin VB.ComboBox cboҽ��ְ�� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2295
      Width           =   1455
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   120
      Width           =   3360
   End
   Begin VB.TextBox txtƴ�� 
      Height          =   300
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1215
      Width           =   1170
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   3450
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1935
      Width           =   1470
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cbo��Դ 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   855
      Width           =   1455
   End
   Begin VB.ComboBox cbo��ֵ 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   495
      Width           =   1455
   End
   Begin VB.ComboBox cbo�ݴ� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1215
      Width           =   1455
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      Left            =   6510
      MaxLength       =   16
      TabIndex        =   19
      Text            =   "0"
      Top             =   2655
      Width           =   1455
   End
   Begin VB.ComboBox cbo��λ 
      Height          =   300
      Left            =   1230
      TabIndex        =   7
      Top             =   1950
      Width           =   1155
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1230
      MaxLength       =   40
      TabIndex        =   3
      Top             =   855
      Width           =   3675
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1230
      MaxLength       =   13
      TabIndex        =   2
      Top             =   495
      Width           =   1935
   End
   Begin VB.ComboBox cbo����ְ�� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1935
      Width           =   1455
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -285
      TabIndex        =   40
      Top             =   5895
      Width           =   8490
   End
   Begin VB.TextBox txt�ο� 
      Height          =   300
      Left            =   1230
      TabIndex        =   9
      Top             =   2295
      Width           =   3135
   End
   Begin VB.CommandButton cmd�ο� 
      Caption         =   "��"
      Height          =   285
      Left            =   4350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2295
      Width           =   285
   End
   Begin VB.CheckBox chkƷ��ҽ�� 
      Caption         =   "ҩƷ��Ʒ���³���ҽ��"
      Height          =   210
      Left            =   5265
      TabIndex        =   31
      Top             =   4890
      Width           =   2115
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3015
      Width           =   1455
   End
   Begin VB.CommandButton cmdDel�ο� 
      Height          =   285
      Left            =   4650
      Picture         =   "frmMediItem.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   4545
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   135
      Width           =   285
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "���������Ʒ��(&A)"
      Height          =   350
      Left            =   1500
      TabIndex        =   37
      Top             =   6225
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "������������(&B)"
      Height          =   350
      Left            =   3450
      TabIndex        =   36
      Top             =   6225
      Width           =   1695
   End
   Begin VB.TextBox txtAtccode 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6510
      MaxLength       =   50
      TabIndex        =   34
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CheckBox chkTumour 
      Caption         =   "����ҩ(&T)"
      Height          =   210
      Left            =   5265
      TabIndex        =   23
      Top             =   3990
      Width           =   1155
   End
   Begin VB.CheckBox chkSolvent 
      Caption         =   "��ý(&M)"
      Height          =   210
      Left            =   6750
      TabIndex        =   28
      Top             =   3990
      Width           =   1155
   End
   Begin VB.CheckBox chkԭ��ҩ 
      Caption         =   "ԭ��ҩ(&P)"
      Height          =   210
      Left            =   5265
      TabIndex        =   24
      Top             =   4275
      Width           =   1155
   End
   Begin VB.CheckBox chkר��ҩ 
      Caption         =   "ר��ҩ"
      Height          =   210
      Left            =   6750
      TabIndex        =   29
      Top             =   4275
      Width           =   1110
   End
   Begin VB.CheckBox chk�������� 
      Caption         =   "��������"
      Height          =   210
      Left            =   5265
      TabIndex        =   25
      Top             =   4575
      Width           =   1140
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4305
      Top             =   6555
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
            Picture         =   "frmMediItem.frx":0A97
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":1031
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":15CB
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":1B65
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   660
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6495
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ZL9BillEdit.BillEdit msf���� 
      Height          =   2805
      Left            =   135
      TabIndex        =   11
      Top             =   3015
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4948
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
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "ע����Ʒ�ֽ�����2003-09-01"
      Height          =   180
      Left            =   135
      TabIndex        =   62
      Top             =   5970
      Width           =   2340
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   61
      Top             =   2715
      Width           =   720
   End
   Begin VB.Label lblӢ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   60
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label LblҩƷ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   59
      Top             =   1635
      Width           =   990
   End
   Begin VB.Label Lblҽ��ְ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��ְ��(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   58
      Top             =   2370
      Width           =   990
   End
   Begin VB.Label Lbl����ְ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ְ��(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   57
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   56
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ƽ���(&S)              (ƴ��)               (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   55
      Top             =   1275
      Width           =   4680
   End
   Begin VB.Label Lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2760
      TabIndex        =   54
      Top             =   1995
      Width           =   630
   End
   Begin VB.Label Lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   53
      Top             =   195
      Width           =   990
   End
   Begin VB.Label Lbl��Դ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Դ���(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   52
      Top             =   915
      Width           =   990
   End
   Begin VB.Label Lbl��ֵ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ֵ����(&V)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   51
      Top             =   555
      Width           =   990
   End
   Begin VB.Label Lbl�ݴ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ҩ�ݴ�(&G)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   50
      Top             =   1275
      Width           =   990
   End
   Begin VB.Label Lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   49
      Top             =   2715
      Width           =   990
   End
   Begin VB.Label Lbl��λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������λ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   48
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͨ������(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   47
      Top             =   915
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   46
      Top             =   555
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "�ο���Ŀ(&F)"
      Height          =   255
      Left            =   165
      TabIndex        =   45
      Top             =   2355
      Width           =   1095
   End
   Begin VB.Label lbl�����Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ա�(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5265
      TabIndex        =   44
      Top             =   3075
      Width           =   990
   End
   Begin VB.Label lblAtccode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ATCCODE(&H)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5310
      TabIndex        =   43
      Top             =   5580
      Width           =   900
   End
End
Attribute VB_Name = "frmMediItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ���ʣ���Me.tag��ţ��ֱ�Ϊ1-����ҩ��2-�г�ҩ�����ϼ�������
'   2���༭״̬����Me.cmdCancel.Tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�������
'---------------------------------------------------
Public lng����id As Long        '���༭�ķ���ID���ϼ����򴫵ݽ���
Public lngҩ��id As Long        '���༭��ҩ��ID���޸ġ�����ʱ���ϼ����򴫵ݽ���
Public strPrivs As String       '��ǰ�û��Ա������Ȩ�ޣ����ϼ�����򴫵ݽ���
Public lng������ As Long         '��ǰ�Ŀ����ؼ���
Private mint������� As Integer     'ҩƷƷ�ֱ����������
Private mblnOK As Boolean       '��¼ȷ����ť�Ƿ񱻵����
Private mblnCancel As Boolean   '��¼ȡ����ť�Ƿ񱻵����

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Dim mstrMatch As String, strRefer As String '�ο�����
Private mstrID As String, mstr���� As String, mstr���� As String
Private mblnLoad As Boolean      '��¼������صĴ���

Private mlng���볤�� As Long
Private mlng���볤�� As Long
Private mint���Ƴ��� As Integer
Private mintӢ�ĳ��� As Integer
Private mstr���м�¼ As String
Private mbln�Թ�ҩ As Boolean

Public Sub ShowMe(ByVal bln�Թ�ҩ As Boolean, ByVal frmPar As Form)
    mbln�Թ�ҩ = bln�Թ�ҩ
    Me.Show vbModal, frmPar
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    
    gstrSql = "Select ���� From �շ���ĿĿ¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
        
    txt����.MaxLength = mlng���볤��
    
    gstrSql = "Select ����,���� From ������Ŀ���� Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mintӢ�ĳ��� = rsTmp.Fields("����").DefinedSize
    
    txtƴ��.MaxLength = mlng���볤��
    txt���.MaxLength = mlng���볤��
    txtӢ��.MaxLength = mintӢ�ĳ���
    
    gstrSql = "Select ���� From ������ĿĿ¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mint���Ƴ��� = rsTmp.Fields("����").DefinedSize
    txt����.MaxLength = mint���Ƴ���
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cbo����ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��λ_GotFocus()
    Me.cbo��λ.SelStart = 0: Me.cbo��λ.SelLength = 100
End Sub

Private Sub cbo��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Or (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Then KeyAscii = 0
    
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkTumour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chkSolvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chk������ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo�ݴ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cboҩƷ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo�����Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cboҽ��ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk����ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk������_Click()
    If Me.chk������.Value = 1 Then
        Me.cbo������.Enabled = True
        txtAtccode.Enabled = True
    Else
        Me.cbo������.Enabled = False
        txtAtccode.Enabled = False
        txtAtccode.Text = ""
    End If
End Sub

Private Sub chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkƤ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkƷ��ҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkԭ��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkԭ��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkר��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Dim strTemp As String
    Dim str���� As String
    Dim i As Integer
    
    With msf����
        For i = 1 To .Rows - 1
            str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    strTemp = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                cbo����.Text & "|" & txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk����ҩ.Value & "|" & chk��ҩ.Value & "|" & chkԭ��ҩ.Value & "|" & _
                chkԭ��ҩ.Value & "|" & chkר��ҩ.Value & "|" & chk��������.Value & "|" & chk������ҩ.Value & "|" & chkƤ��.Value & "|" & chkƷ��ҽ��.Value & "|" & chk������.Value & "|" & cbo������.Text & "|" & str���� & "|" & txtAtccode.Text
        If strTemp <> mstr���м�¼ Then
        mblnCancel = True
        If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
            gblnCancel = True
            Unload Me
        Else
            mblnCancel = False
        End If
    Else
        gblnCancel = True
        Unload Me
    End If
    Exit Sub
End Sub

Private Sub cmdDel�ο�_Click()
    Me.txt�ο�.Text = ""
    Me.txt�ο�.Tag = ""
    strRefer = ""
    Me.txt�ο�.SetFocus
End Sub

Private Sub cmdOK_Click()

    '�༭���ݼ��
    mblnOK = True
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "������ҩƷ���룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mlng���볤�� Then
        MsgBox "ҩƷ����ĳ��ȳ��������" & mlng���볤�� & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "������ͨ�����ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mint���Ƴ��� Then
        MsgBox "ͨ�����Ƴ��ȳ��������" & mint���Ƴ��� & "���ַ���" & Int(mint���Ƴ��� / 2) & "�����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If Trim(Me.cbo��λ.Text) = "" Then
        MsgBox "�����������λ��", vbInformation, gstrSysName
        Me.cbo��λ.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.cbo��λ.Text), vbFromUnicode)) > 10 Then
        MsgBox "������λ�ĳ��ȳ��������10���ַ���5�����֣���", vbInformation, gstrSysName
        Me.cbo��λ.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtAtccode.Text), vbFromUnicode)) > 50 Then
        MsgBox "ATCCODE�ĳ��ȳ��������50���ַ���25�����֣���", vbInformation, gstrSysName
        Me.txtAtccode.SetFocus: Exit Sub
    End If
    
    '�������
    strTemp = ";" & Trim(Me.txt����.Text) & ";" & Trim(Me.txtӢ��.Text)
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "���������ظ�������ͨ�����ƺ�Ӣ�����ƣ���", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    
    '���ݱ���
    If Me.cmdCancel.Tag = "����" Then
        lngҩ��id = Sys.NextId("������ĿĿ¼")
        If zlClinicCodeRepeat(Trim(Me.txt����.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(Trim(Me.txt����.Text), lngҩ��id) = True Then Exit Sub
    End If
    gstrSql = Me.txt����.Tag & "," & lngҩ��id & ",'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.txtƴ��.Text)) & "','" & MoveSpecialChar(Trim(Me.txt���.Text)) & "','" & MoveSpecialChar(Trim(Me.txtӢ��.Text)) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.cbo��λ.Text)) & "','" & Mid(Me.cbo����.Text, InStr(1, Me.cbo����.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo����.Text, InStr(1, Me.cbo����.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��ֵ.Text, InStr(1, Me.cbo��ֵ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��Դ.Text, InStr(1, Me.cbo��Դ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo�ݴ�.Text, InStr(1, Me.cbo�ݴ�.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Left(Me.cboҩƷ����.Text, 1) & ",'" & Left(Me.cbo����ְ��.Text, 1) & Left(Me.cboҽ��ְ��.Text, 1) & "'"
    gstrSql = gstrSql & "," & Val(Trim(Me.txt��������.Text))
    gstrSql = gstrSql & "," & Me.chk����ҩ.Value & "," & Me.chk��ҩ.Value & "," & Me.chkԭ��ҩ.Value & "," & Me.chkƤ��.Value & "," & IIf(Me.chk������.Value = 0, Me.chk������.Value, Me.cbo������.ListIndex + 1)
    gstrSql = gstrSql & "," & ZVal(Me.txt�ο�.Tag) & "," & Me.chkƷ��ҽ��.Value & "," & Left(Me.cbo�����Ա�.Text, 1)
    strTemp = ""
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    '������������
    If LenB(strTemp) > 4000 Then
        msf����.SetFocus
        MsgBox "�����ַ���̫��������ٱ����������߱������ȡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    err = 0: On Error GoTo ErrHand
    If Me.cmdCancel.Tag = "����" Then
        gstrSql = "zl_��ҩƷ��_INSERT('" & IIf(Me.Tag = 1, "5", "6") & "'," & gstrSql & ",'" & strTemp & "'," & IIf(mbln�Թ�ҩ = True, "1", "Null") & "," & IIf(Trim(txtAtccode.Text) = "", "NULL,", "'" & txtAtccode.Text & "',") & Me.chkTumour.Value & "," & Me.chkSolvent.Value & "," & Me.chkԭ��ҩ.Value & "," & Me.chkר��ҩ.Value & "," & Me.chk��������.Value & "," & Me.chk������ҩ.Value & ")"
    Else
        gstrSql = "zl_��ҩƷ��_UPDATE(" & gstrSql & ",'" & strTemp & "'," & IIf(mbln�Թ�ҩ = True, "1", "Null") & "," & IIf(Trim(txtAtccode.Text) = "", "NULL,", "'" & txtAtccode.Text & "',") & Me.chkTumour.Value & "," & Me.chkSolvent.Value & "," & Me.chkԭ��ҩ.Value & "," & Me.chkר��ҩ.Value & "," & Me.chk��������.Value & "," & Me.chk������ҩ.Value & ")"
    End If
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    
    If Me.cmdCancel.Tag = "����" Then
        'Val(zldatabase.GetPara("Ʒ������ģʽ", glngSys, 1023, 0))
        Select Case ActiveControl
        Case cmdSaveAddItem  'Ʒ����������
            Call frmMediLists.zlRefRecords(lngҩ��id)
            lngҩ��id = 0
            Call Form_Activate
            Me.txt����.SetFocus
            mblnOK = False
        Case cmdSaveAddSpec  'Ʒ�����Ӻ����ӹ��
            Unload Me
            mblnOK = False
            Call frmMediLists.zlRefRecords(lngҩ��id)
            With frmMediSpec
                .mlng����id = lng����id
                .stbSpec.Tag = "����"
                .lngҩ��id = lngҩ��id
                .lngҩƷID = 0
                .strPrivs = Me.strPrivs
                .Show 1, frmMediLists
            End With
        Case Else
            Unload Me
        End Select
    Else
        Unload Me
    End If
    
    If lng������ <> 0 And mblnOK = True Then Call frmMediLists.ZlRefBut(3)
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    
    Call cmdOK_Click
End Sub

Private Sub cmd����_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmd�ο�_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo errHandle
    strSql = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����id)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = " Select ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=" & iAttr & " Order By ����"
    Else
        strSQLItem = " From ���Ʋο�Ŀ¼ A,���Ʋο����� B" & _
            " Where A.ID=B.�ο�Ŀ¼ID And A.����=" & iAttr & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = " Select DISTINCT A.ID,A.����ID,A.����,A.����,A.˵�� " & strSQLItem & " Order By ����"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 0, "�ο�", , , , , True)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmd����_Click()
    Dim blnRe As Boolean
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = " & Me.Tag & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    blnRe = frmTreeSel.ShowTree(gstrSql, mstrID, mstr����, mstr����, "", "ҩƷ����", "���з���", False)
    If blnRe Then
        txt����.Text = "[" & mstr���� & "]" & mstr����
        txt����.Tag = mstrID
        Me.txt����.SetFocus
        lng����id = mstrID
    End If
    mblnLoad = True
End Sub

Private Sub Command2_Click()
    Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    Dim strCode As String
    Dim str���� As String
    Dim i As Integer
    
    gblnCancel = False
    If cmdCancel.Tag <> "����" Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    
    If mblnLoad = False Then
        If Me.Tag = "1" Then
            Me.Caption = "����ҩƷ��" & Me.cmdCancel.Tag
        Else
            Me.Caption = "�г�ҩƷ��" & Me.cmdCancel.Tag
        End If
        
        '�������ݼ��
        If Me.cbo����.ListCount = 0 Then MsgBox "δ����ҩƷ���ͣ������ֵ�����н�������", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo����.ListCount = 0 Then MsgBox "�޶���������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo��ֵ.ListCount = 0 Then MsgBox "�޼�ֵ�������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo��Դ.ListCount = 0 Then MsgBox "�޻�Դ�������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo�ݴ�.ListCount = 0 Then MsgBox "����ҩ�ݴ����ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo������.ListCount <> 0 Then
            Me.cbo������.ListIndex = IIf(lng������ <> 0, lng������ - 1, 0)
            If lng������ <> 0 Then
                Me.chk������.Value = 1
                Me.cbo������.Enabled = True
                txtAtccode.Enabled = True
            End If
        End If
        
        '�������޸�ʱ������Ȩ������
        If InStr(1, strPrivs, "ҽ����ҩĿ¼") = 0 Then
            Me.cboҽ��ְ��.Enabled = False
        End If
        
        '�������޸�ʱ��װ�뱾���͵ķ���
        err = 0: On Error GoTo ErrHand
        
        '����ѡ����װ��
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ���Ʒ���Ŀ¼" & _
                " Where ���� = [1] and id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, lng����id)
        
        With rsTemp
            Me.txt����.Text = "[" & !���� & "]" & !����
            Me.txt����.Tag = lng����id
        End With
        
        '���ݱ༭״̬��������������ʾ
        If Me.cmdCancel.Tag = "����" Then
            lngҩ��id = 0
    
            If mint������� = 0 Then
'                gstrSql = "select nvl(max(����),'0000000') as ����" & _
'                        " From ������ĿĿ¼" & _
'                        " Where ��� = [1]"

                gstrSql = "Select Nvl(Max(����), '0000000') As ����" & vbNewLine & _
                                "From (Select ���� From ������ĿĿ¼ Where ��� = [1] Order By Length(����) Desc, ���� Desc, ����ʱ�� Desc)" & vbNewLine & _
                                "Where Rownum = 1 "
                                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = 1, "5", "6"))
                
'                Me.txt����.Text = Right(String(13, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����))
                Me.txt����.Text = zlCommFun.IncStr(rsTemp!����)

            Else
                strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
'                gstrSql = "select nvl(max(����),'') as ����" & _
'                        " From ������ĿĿ¼" & _
'                        " Where ��� = [1] and ���� like [2] and length(����)>=[3]"

                gstrSql = "Select Nvl(Max(����), '') As ����" & vbNewLine & _
                                "From (Select ����" & vbNewLine & _
                                "       From ������ĿĿ¼" & vbNewLine & _
                                "       Where ��� = [1] And ���� Like [2] And Length(����) >=[3] " & vbNewLine & _
                                "       Order By Length(����) Desc, ���� Desc, ����ʱ�� Desc)" & vbNewLine & _
                                "Where Rownum = 1"

                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = 1, "5", "6"), IIf(Me.Tag = 1, "5", "6") & strTemp & "%", Len("*" & strTemp & "**"))
                
                err = 0: On Error Resume Next
    '            Me.txt����.Text = IIf(Me.Tag = 1, "5", "6") & strTemp & Right(String(13, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����) - 1 - Len(strTemp))
                
                strTemp = IIf(Me.Tag = 1, "5", "6") & strTemp
                If Nvl(rsTemp!����) = "" Then
                    Me.txt����.Text = strTemp & "01"
                Else
                    strCode = rsTemp!����
                    strCode = Mid(strCode, Len(strTemp) + 1)
                    strCode = zlCommFun.IncStr(strCode)
                    Me.txt����.Text = strTemp & strCode
                End If
            End If
    
            Me.txt����.Text = "": Me.txtӢ��.Text = ""
            Me.lblNote.Visible = False
            Me.txt�ο� = "": Me.txt�ο�.Tag = "": strRefer = ""
        Else
            '������Ϣ��Ŀ
            gstrSql = "select I.����ID,I.����,I.����,I.���㵥λ,T.ҩƷ����," & _
                    "        T.�������,T.��Դ���,T.��ֵ����,T.��ҩ�ݴ�," & _
                    "        nvl(T.ҩƷ����,0) as ҩƷ����,nvl(T.����ְ��,'00') as ����ְ��,nvl(T.��������,0) as ��������," & _
                    "        nvl(T.����ҩ��,0) as ����ҩ��,nvl(T.�Ƿ�ԭ��,0) as �Ƿ�ԭ��,nvl(t.������,0) as ������,nvl(T.�Ƿ���ҩ,0) as �Ƿ���ҩ,nvl(T.�Ƿ�Ƥ��,0) as �Ƿ�Ƥ��,nvl(T.�Ƿ�ԭ��ҩ,0) as �Ƿ�ԭ��ҩ,nvl(T.�Ƿ�ר��ҩ,0) as �Ƿ�ר��ҩ,nvl(T.�Ƿ񵥶�����,0) as �Ƿ񵥶�����,Nvl(t.�Ƿ�����ҩ, 0) As �Ƿ�����ҩ," & _
                    "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,B.���� as �ο�����,I.�ο�Ŀ¼id,nvl(T.Ʒ��ҽ��,0) as Ʒ��ҽ��,Nvl(I.�����Ա�,0) AS �����Ա�,t.ATCCODE,nvl(T.�Ƿ�����ҩ,0) as ����ҩ,T.��ý" & _
                    " from ������ĿĿ¼ I,ҩƷ���� T,���Ʋο�Ŀ¼ B" & _
                    " where I.ID=T.ҩ��ID and I.ID=[1] and I.�ο�Ŀ¼id=B.id(+) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                If Not .EOF Then
                    Me.lblNote.Caption = "ע����ҩƷ������" & Format(!����ʱ��, "YYYY-MM-DD")
                    If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        Me.lblNote.Caption = Me.lblNote.Caption & "����" & Format(!����ʱ��, "YYYY-MM-DD") & "ͣ�á�"
                    End If
                    Me.txt����.Text = !����
                    Me.txt����.Text = !����
                    Me.cbo��λ.Text = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                    Me.txt�ο�.Text = Nvl(!�ο�����)
                    Me.txt�ο�.Tag = Nvl(!�ο�Ŀ¼ID)
                    strRefer = Me.txt�ο�.Text
                    For intCount = 0 To Me.cbo����.ListCount - 1
                        If Mid(Me.cbo����.List(intCount), InStr(1, Me.cbo����.List(intCount), "-") + 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����) Then
                            Me.cbo����.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo����.ListCount - 1
                        If Mid(Me.cbo����.List(intCount), InStr(1, Me.cbo����.List(intCount), "-") + 1) = IIf(IsNull(!�������), "", !�������) Then
                            Me.cbo����.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo��ֵ.ListCount - 1
                        If Mid(Me.cbo��ֵ.List(intCount), InStr(1, Me.cbo��ֵ.List(intCount), "-") + 1) = IIf(IsNull(!��ֵ����), "", !��ֵ����) Then
                            Me.cbo��ֵ.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo��Դ.ListCount - 1
                        If Mid(Me.cbo��Դ.List(intCount), InStr(1, Me.cbo��Դ.List(intCount), "-") + 1) = IIf(IsNull(!��Դ���), "", !��Դ���) Then
                            Me.cbo��Դ.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo�ݴ�.ListCount - 1
                        If Mid(Me.cbo�ݴ�.List(intCount), InStr(1, Me.cbo�ݴ�.List(intCount), "-") + 1) = IIf(IsNull(!��ҩ�ݴ�), "", !��ҩ�ݴ�) Then
                            Me.cbo�ݴ�.ListIndex = intCount: Exit For
                        End If
                    Next
                    Me.cboҩƷ����.ListIndex = !ҩƷ����
                    Me.cbo����ְ��.ListIndex = IIf(CInt(Left(Format(!����ְ��, "00"), 1)) <> 9, CInt(Left(Format(!����ְ��, "00"), 1)), Me.cbo����ְ��.ListCount - 1)
                    Me.cboҽ��ְ��.ListIndex = IIf(CInt(Right(Format(!����ְ��, "00"), 1)) <> 9, CInt(Right(Format(!����ְ��, "00"), 1)), Me.cboҽ��ְ��.ListCount - 1)
                    Me.txt��������.Text = !��������
                    Me.txtAtccode.Text = IIf(IsNull(!ATCCODE), "", !ATCCODE)
                    Me.chk����ҩ.Value = IIf(!����ҩ�� = 0, 0, 1)
                    Me.chkԭ��ҩ.Value = IIf(!�Ƿ�ԭ�� = 0, 0, 1)
                    Me.chk��ҩ.Value = IIf(!�Ƿ���ҩ = 0, 0, 1)
                    Me.chkƤ��.Value = IIf(!�Ƿ�Ƥ�� = 0, 0, 1)
                    Me.chkTumour.Value = IIf(!����ҩ = 0, 0, 1)
                    Me.chkԭ��ҩ.Value = IIf(!�Ƿ�ԭ��ҩ = 0, 0, 1)
                    Me.chkר��ҩ.Value = IIf(!�Ƿ�ר��ҩ = 0, 0, 1)
                    Me.chk��������.Value = IIf(!�Ƿ񵥶����� = 0, 0, 1)
                    Me.chk������ҩ.Value = IIf(!�Ƿ�����ҩ = 0, 0, 1)
                    
                    '���˺�:2008/03/17���뿹����,��ҪӦ����Ժ��ϵͳ�Ŀ���ҩ��ļ��:12753
                    Me.chk������.Value = IIf(!������ <> 0, 1, 0)
                    If !������ <> 0 Then
                        Me.cbo������.ListIndex = !������ - 1
                    End If
                    Me.chkƷ��ҽ��.Value = IIf(!Ʒ��ҽ�� = 0, 0, 1)
                    Me.cbo�����Ա�.ListIndex = !�����Ա�
                    Me.chkSolvent.Value = IIf(Nvl(!��ý, 0) = 0, 0, 1)
                End If
            End With
            
            '����������Ӣ����
            gstrSql = "select ����,����,����,���� from ������Ŀ���� where ���� in (1,2) and ������ĿID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                Do While Not .EOF
                    If !���� = 1 And !���� = 1 Then Me.txtƴ��.Text = !����
                    If !���� = 1 And !���� = 2 Then Me.txt���.Text = !����
                    If !���� = 2 Then Me.txtӢ��.Text = !����
                    .MoveNext
                Loop
            End With
                
            '��������
            gstrSql = "select N.����,P.���� as ƴ��,W.���� as ���" & _
                    " from (select distinct ���� from ������Ŀ���� where ������ĿID=[1] and ����=9) N," & _
                    "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=1) P," & _
                    "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=2) W" & _
                    " where N.����=P.����(+) and N.����=W.����(+)"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                Do While Not .EOF
                    If Me.msf����.Rows - 1 < .AbsolutePosition Then Me.msf����.Rows = Me.msf����.Rows + 1
                    Me.msf����.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                    Me.msf����.TextMatrix(.AbsolutePosition, 1) = !����
                    Me.msf����.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!ƴ��), "", !ƴ��)
                    Me.msf����.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!���), "", !���)
                    .MoveNext
                Loop
            End With
        End If
        
        If Me.cmdCancel.Tag = "����" Then
            '����ʱ�����ؼ��ı༭״̬
            Me.cmdOK.Visible = False
            Me.cmdCancel.Caption = "�ر�(&C)"
            Me.txt����.Enabled = False: Me.cmd����.Enabled = False
            Me.txt����.Enabled = False
            Me.txt����.Enabled = False
            Me.txtƴ��.Enabled = False: Me.txt���.Enabled = False
            Me.txtӢ��.Enabled = False
            Me.cbo��λ.Enabled = False: Me.cbo����.Enabled = False
            Me.cbo����.Enabled = False: Me.cbo��ֵ.Enabled = False: Me.cbo��Դ.Enabled = False: Me.cbo�ݴ�.Enabled = False
            Me.cboҩƷ����.Enabled = False: Me.cbo����ְ��.Enabled = False: Me.cboҽ��ְ��.Enabled = False: Me.txt��������.Enabled = False: txtAtccode.Enabled = False
            Me.chk����ҩ.Enabled = False: Me.chkԭ��ҩ.Enabled = False: Me.chkƤ��.Enabled = False: Me.chk��ҩ.Enabled = False
            Me.chkƷ��ҽ��.Enabled = False
            Me.chk������.Enabled = False
            Me.msf����.Active = False
            Me.txt�ο�.Enabled = False
            Me.cmd�ο�.Enabled = False
            Me.cmdDel�ο�.Enabled = False
            Me.cbo�����Ա�.Enabled = False
            Me.chkԭ��ҩ.Enabled = False
            Me.chkר��ҩ.Enabled = False
            Me.chk��������.Enabled = False
            Me.chk������ҩ.Enabled = False
        End If
    End If
    mstr���м�¼ = ""
    str���� = ""
    With msf����
        For i = 1 To .Rows - 1
            str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    mstr���м�¼ = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                cbo����.Text & "|" & txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk����ҩ.Value & "|" & chk��ҩ.Value & "|" & chkԭ��ҩ.Value & "|" & _
                chkԭ��ҩ.Value & "|" & chkר��ҩ.Value & "|" & chk��������.Value & "|" & chk������ҩ.Value & "|" & chkƤ��.Value & "|" & chkƷ��ҽ��.Value & "|" & chk������.Value & "|" & cbo������.Text & "|" & str���� & "|" & txtAtccode.Text
    If txt����.Enabled = True Then
        txt����.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    'ȡҩƷƷ�ֱ���Ĳ�������
    mint������� = Val(zlDatabase.GetPara(87, glngSys))
    
    Call GetDefineSize
    
    '-------------����ѡ������װ��-----------------------
    On Error GoTo errHandle
    With rsTemp
        gstrSql = "select ����||'-'||���� from ҩƷ���� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo����.Clear
        Do While Not rsTemp.EOF
            Me.cbo����.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    
        gstrSql = "select distinct ���㵥λ from ������ĿĿ¼ where ��� in ('5','6') and ���㵥λ is not null"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Do While Not rsTemp.EOF
            Me.cbo��λ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        
        gstrSql = "select ����||'-'||���� from ҩƷ������� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo����.Clear
        Do While Not rsTemp.EOF
            Me.cbo����.AddItem rsTemp.Fields(0).Value
            If InStr(1, rsTemp.Fields(0).Value, "��ͨ") > 0 Then
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If Me.cbo����.ListIndex = -1 And Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ֵ���� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo��ֵ.Clear
        Do While Not rsTemp.EOF
            Me.cbo��ֵ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo��ֵ.ListCount > 0 Then Me.cbo��ֵ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��Դ��� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo��Դ.Clear
        Do While Not rsTemp.EOF
            Me.cbo��Դ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo��Դ.ListCount > 0 Then Me.cbo��Դ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ҩ�ݴ� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo�ݴ�.Clear
        Do While Not rsTemp.EOF
            Me.cbo�ݴ�.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo�ݴ�.ListCount > 0 Then Me.cbo�ݴ�.ListIndex = 0
    End With
    
    aryTemp = Split("0-δ�趨;1-����ҩ;2-����Ǵ���ҩ;3-����Ǵ���ҩ;4-�Ǵ���ҩ;5-������ҩ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cboҩƷ����.AddItem aryTemp(intCount)
    Next
    Me.cboҩƷ����.ListIndex = 0
    
    aryTemp = Split("0-����;1-����;2-����;3-�м�;4-����/ʦ��;5-Ա/ʿ;9-��Ƹ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����ְ��.AddItem aryTemp(intCount): Me.cboҽ��ְ��.AddItem aryTemp(intCount)
    Next
    Me.cbo����ְ��.ListIndex = 0: Me.cboҽ��ְ��.ListIndex = 0
    
    aryTemp = Split("0-���Ա�����;1-����;2-Ů��", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�����Ա�.AddItem aryTemp(intCount)
    Next
    Me.cbo�����Ա�.ListIndex = 0
    
    aryTemp = Split("1-������ʹ��;2-����ʹ��;3-����ʹ��", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo������.AddItem aryTemp(intCount)
    Next
    Me.cbo������.ListIndex = 0
    
    '��ʼ�����ñ��༭
    With Me.msf����
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "ҩƷ����": .TextMatrix(0, 2) = "ƴ����": .TextMatrix(0, 3) = "�����"
        .colData(0) = 5: .colData(1) = 4: .colData(2) = 4: .colData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 2500: .ColWidth(2) = 950: .ColWidth(3) = 950
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    strRefer = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim str���� As String
    Dim i As Integer
    
    If mblnOK = False And mblnCancel = False Then
        With msf����
            For i = 1 To .Rows - 1
                str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
            Next
        End With
        strTemp = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                    cbo����.Text & "|" & txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                    cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk����ҩ.Value & "|" & chk��ҩ.Value & "|" & chkԭ��ҩ.Value & "|" & _
                    chkԭ��ҩ.Value & "|" & chkר��ҩ.Value & "|" & chk��������.Value & "|" & chk������ҩ.Value & "|" & chkƤ��.Value & "|" & chkƷ��ҽ��.Value & "|" & chk������.Value & "|" & cbo������.Text & "|" & str���� & "|" & txtAtccode.Text
        If strTemp <> mstr���м�¼ Then
            If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
                mblnLoad = False
                mblnOK = False
                mblnCancel = False
                If mblnOK = True Then
                    gblnCancel = True
                End If
            Else
                Cancel = 1
            End If
        Else
            mblnLoad = False
            mblnOK = False
            mblnCancel = False
            If mblnOK = True Then
                gblnCancel = True
            End If
        End If
    End If
    mblnLoad = False
    mblnOK = False
    mblnCancel = False
End Sub

Private Sub msf����_AfterAddRow(Row As Long)
    With Me.msf����
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_AfterDeleteRow()
    With Me.msf����
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_EditKeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Call OS.PressKey(vbKeyTab): Exit Sub
            strTemp = Trim(.Text)
            If strTemp <> "" Then
                .TextMatrix(.Row, 1) = strTemp
                .TextMatrix(.Row, 2) = zlStr.GetCodeByORCL(strTemp, False, mlng���볤��)
                .TextMatrix(.Row, 3) = zlStr.GetCodeByORCL(strTemp, True, mlng���볤��)
            Else
                Call OS.PressKey(vbKeyTab): Exit Sub
            End If
        End If
    End With
End Sub

Private Sub msf����_KeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtAtccode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Asc("-")
        If InStr(1, txt����.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ο�_GotFocus()
    Me.txt�ο�.SelStart = 0: Me.txt�ο�.SelLength = 100
End Sub


Private Sub txt�ο�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt�ο� <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt�ο�))
            If rsTmp Is Nothing Then
                Me.txt�ο� = strRefer
                Me.SetFocus
                Exit Sub
            Else
                Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
            End If
        End If
        Call OS.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt�ο�_LostFocus()
    If Me.txt�ο� <> strRefer Then
        Me.txt�ο� = strRefer
    End If
End Sub


Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 100
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    Dim strTmp As String
    '���¼�����ƣ���ȥ �������ַ�
    strTmp = MoveSpecialChar(txt����.Text)
    If txt����.Text <> strTmp Then
        txt����.Text = strTmp
    End If
    Me.txtƴ��.Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
    Me.txt���.Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
    End Select
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    Else
        If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Me.txtƴ��.Text = zlStr.GetCodeByORCL(Me.txt����.Text, False, mlng���볤��)
        Me.txt���.Text = zlStr.GetCodeByORCL(Me.txt����.Text, True, mlng���볤��)
    End If
    
        
End Sub

Private Sub txt����_LostFocus()
'    Me.txtƴ��.Text = zlGetSymbol(Me.txt����.Text, 0, mlng���볤��)
'    Me.txt���.Text = zlGetSymbol(Me.txt����.Text, 1, mlng���볤��)
    Call OS.OpenIme(False)
End Sub

Private Sub txtƴ��_GotFocus()
    Me.txtƴ��.SelStart = 0: Me.txtƴ��.SelLength = 100
End Sub

Private Sub txtƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 100
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
        Case vbKeyReturn
            Call OS.PressKey(vbKeyTab)
    End Select
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
