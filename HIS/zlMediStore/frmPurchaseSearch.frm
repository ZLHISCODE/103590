VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmPurchaseSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   5010
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmPurchaseSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1560
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   4815
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmPurchaseSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmPurchaseSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   5505
         Begin MSComctlLib.ListView lvw���� 
            Height          =   2835
            Left            =   1200
            TabIndex        =   39
            Top             =   3360
            Visible         =   0   'False
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   5001
            View            =   1
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "imgsDrug"
            SmallIcons      =   "imgsDrug"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TreeView tvw��� 
            Height          =   4245
            Left            =   120
            TabIndex        =   38
            Top             =   3240
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   7488
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imgsDrug"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.CheckBox chk��Ʊ���� 
            Caption         =   "��Ʊ�������"
            Height          =   405
            Left            =   600
            TabIndex        =   55
            Top             =   2340
            Width           =   1035
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ"
            Height          =   300
            Left            =   600
            TabIndex        =   54
            Top             =   1140
            Width           =   990
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1140
            Width           =   3255
         End
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   52
            Top             =   1140
            Width           =   255
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   51
            Top             =   750
            Width           =   3255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   50
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Cmd������ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   49
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox txt������ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   48
            Top             =   1920
            Width           =   3255
         End
         Begin VB.CheckBox Chk������ 
            Caption         =   "������"
            Height          =   300
            Left            =   600
            TabIndex        =   47
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   46
            Top             =   1560
            Width           =   255
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1530
            Width           =   3255
         End
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��"
            Height          =   300
            Left            =   600
            TabIndex        =   44
            Top             =   1530
            Width           =   1110
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   43
            Top             =   750
            Width           =   255
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   600
            TabIndex        =   42
            Top             =   750
            Width           =   1095
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   41
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox chkClass 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   600
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   9
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   10
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt��ʼ��Ʊ�� 
            Height          =   300
            Left            =   1530
            TabIndex        =   11
            Top             =   3330
            Width           =   1365
         End
         Begin VB.TextBox Txt������Ʊ�� 
            Height          =   300
            Left            =   3780
            TabIndex        =   12
            Top             =   3330
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpStart��Ʊ 
            Height          =   315
            Left            =   1650
            TabIndex        =   35
            Top             =   2340
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpEnd��Ʊ 
            Height          =   315
            Left            =   3600
            TabIndex        =   36
            Top             =   2340
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   3360
            TabIndex        =   37
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   975
            TabIndex        =   28
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   3120
            TabIndex        =   27
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl��Ʊ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊ��"
            Height          =   180
            Left            =   975
            TabIndex        =   26
            Top             =   3390
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   25
            Top             =   3390
            Width           =   180
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   4050
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "δ����˿�"
            Height          =   180
            Left            =   720
            TabIndex        =   57
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "������˿�"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2400
            TabIndex        =   56
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CheckBox chkAcc 
            Caption         =   "δ�������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   34
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CheckBox chk�޷�Ʊ 
            Caption         =   "�޷�Ʊ"
            Height          =   180
            Left            =   2400
            TabIndex        =   33
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CheckBox chk�з�Ʊ 
            Caption         =   "�з�Ʊ"
            Height          =   180
            Left            =   720
            TabIndex        =   32
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CheckBox chkδ��� 
            Caption         =   "δ��������"
            Height          =   255
            Left            =   2400
            TabIndex        =   31
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CheckBox chk�ѱ�� 
            Caption         =   "����������"
            Height          =   255
            Left            =   720
            TabIndex        =   30
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CheckBox chkAccStrike 
            Caption         =   "�Ѳ������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   29
            Top             =   2640
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox txt��ʼNo 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt����NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246284291
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246284291
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   23
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   22
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   21
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   20
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   19
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   18
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   14
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   13
      Top             =   435
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   6480
      Top             =   1200
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
            Picture         =   "frmPurchaseSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '�����ַ���
Private BlnAdvance As Boolean '�Ƿ�չ��
Private mdatStart As Date   '��ʼʱ��
Private mdatEnd As Date     '����ʱ��
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '������
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Public lngҩƷid As Long
Private mstrMatch As String 'ƥ�䷽ʽ 0-˫��ƥ�� 1-�������ҵ���ƥ��

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    str������ As String
    str����� As String
    lng������ As Long
    str���� As String
    str��Ʊ�ſ�ʼ As String
    str��Ʊ�Ž��� As String
    int�������һ����ѯ As Integer
    intδ��� As Integer
    int�ѱ�� As Integer
    int�з�Ʊ As Integer
    int�޷�Ʊ As Integer
    lngҩƷ���� As Long
    str���� As String
    date��Ʊ������ڿ�ʼ As Date
    date��Ʊ������ڽ��� As Date
End Type

Private SQLCondition As Type_SQLCondition

Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO��ʼ As String, _
        ByRef strNO���� As String, _
        ByRef date����ʱ�俪ʼ As Date, _
        ByRef date����ʱ����� As Date, _
        ByRef date���ʱ�俪ʼ As Date, _
        ByRef date���ʱ����� As Date, _
        ByRef lngҩƷ As Long, _
        ByRef str������ As String, _
        ByRef str����� As String, _
        ByRef lng������ As Long, _
        ByRef str���� As String, _
        ByRef str��Ʊ�ſ�ʼ As String, _
        ByRef str��Ʊ�Ž��� As String, _
        ByRef lngҩƷ���� As Long, _
        ByRef str���� As String, _
        ByRef date��Ʊ������ڿ�ʼ As Date, _
        ByRef date��Ʊ������ڽ��� As Date, _
        ByRef intNo��� As Integer, _
        ByRef intYes��� As Integer, _
        ByRef intNo��Ʊ As Integer, _
        ByRef intYes��Ʊ As Integer, _
        Optional ByRef intTmp As Integer = 0) As String
    mstrFind = ""
    mstrSelectTag = ""
    Set mfrmMain = FrmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
    strNO��ʼ = SQLCondition.strNO��ʼ
    strNO���� = SQLCondition.strNO����
    date����ʱ�俪ʼ = SQLCondition.date����ʱ�俪ʼ
    date����ʱ����� = SQLCondition.date����ʱ�����
    date���ʱ�俪ʼ = SQLCondition.date���ʱ�俪ʼ
    date���ʱ����� = SQLCondition.date���ʱ�����
    lngҩƷ = SQLCondition.lngҩƷ
    str����� = SQLCondition.str�����
    str������ = SQLCondition.str������
    lng������ = SQLCondition.lng������
    str���� = SQLCondition.str����
    str��Ʊ�ſ�ʼ = SQLCondition.str��Ʊ�ſ�ʼ
    str��Ʊ�Ž��� = SQLCondition.str��Ʊ�Ž���
    lngҩƷ���� = SQLCondition.lngҩƷ����
    str���� = SQLCondition.str����
    date��Ʊ������ڿ�ʼ = SQLCondition.date��Ʊ������ڿ�ʼ
    date��Ʊ������ڽ��� = SQLCondition.date��Ʊ������ڽ���
    intNo��� = SQLCondition.intδ���
    intYes��� = SQLCondition.int�ѱ��
    intNo��Ʊ = SQLCondition.int�޷�Ʊ
    intYes��Ʊ = SQLCondition.int�з�Ʊ
    intTmp = SQLCondition.int�������һ����ѯ
End Function


Private Sub chkClass_Click()
    If chkClass.Value = 1 Then
        txtClass.Enabled = True
        cmdClass.Enabled = True
    Else
        txtClass.Enabled = False
        cmdClass.Enabled = False
    End If
End Sub

Private Sub chkJiXin_Click()
    If chkJiXin.Value = 1 Then
        txtJiXing.Enabled = True
        cmdJiXin.Enabled = True
    Else
        txtJiXing.Enabled = False
        cmdJiXin.Enabled = False
    End If
End Sub

Private Sub chkStrike_Click()
    chkAccStrike.Enabled = IIf(chkStrike.Value = 1, True, False)
End Sub
Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        cmdȷ��.SetFocus
    End If
    
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        cmdȷ��.SetFocus
    End If
End Sub

Private Sub chk��Ʊ����_Click()
    If chk��Ʊ����.Value = 1 Then
        dtpStart��Ʊ.Enabled = True
        dtpEnd��Ʊ.Enabled = True
    Else
        dtpStart��Ʊ.Enabled = False
        dtpEnd��Ʊ.Enabled = False
    End If
End Sub

Private Sub Chk��Ӧ��_Click()
    txt��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    Cmd��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    
End Sub

Private Sub Chk��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk��Ӧ��.Value = 1 Then
        txt��Ӧ��.SetFocus
    Else
        Chk������.SetFocus
    End If
End Sub


Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk���.Value = 0 Then
            cmdȷ��.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub Chk������_Click()
    Me.txt������.Enabled = IIf(Chk������.Value = 1, True, False)
    Cmd������.Enabled = IIf(Chk������.Value = 1, True, False)
End Sub

Private Sub Chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk������.Value = 1 Then
            txt������.SetFocus
        
        Else
            Txt������.SetFocus
        End If
    End If
End Sub

Private Sub chk����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    chkNOVerifyBack.Enabled = IIf(chk����.Value = 1, True, False)
    If chk����.Value = 0 Then chkNOVerifyBack.Value = 0
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
    chk�ѱ��.Enabled = IIf(chk���.Value = 1, True, False)
    chkδ���.Enabled = IIf(chk���.Value = 1, True, False)
    chkAcc.Enabled = IIf(chk���.Value = 1, True, False)
    chkYesVerifyBack.Enabled = IIf(chk���.Value = 1, True, False)
    If chk���.Value = 0 Then chkYesVerifyBack.Value = 0
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    CmdҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
End Sub

Private Sub ChkҩƷ_GotFocus()
    sstFilter.Tab = 1
    ChkҩƷ.SetFocus
End Sub

Private Sub ChkҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    If ChkҩƷ.Value = 1 Then
        TxtҩƷ.SetFocus
    ElseIf Chk��Ӧ��.Visible = True Then
        Chk��Ӧ��.SetFocus
    End If
End Sub



Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿid As Long
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    tvw���.Left = txtClass.Left
    tvw���.Top = txtClass.Top + txtClass.Height
    tvw���.Visible = True
    tvw���.SetFocus
        
    gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
              "Where Instr([1], ����, 1) > 0 " & _
              "Order by ���� "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw���
        .Nodes.Clear
'        Set nodTmp = .Nodes.Add(, , "Root", "����", 2, 2)
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!����, rsTmp!����, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!����
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    gstrSQL = "Select ID, �ϼ�ID, ����, 1 as ĩ��, decode(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') as ����, ���� " & _
                  "From ���Ʒ���Ŀ¼ " & _
                  "Where ���� in (1,2,3) " & _
                  "Start With �ϼ�ID IS NULL Connect By Prior ID=�ϼ�ID Order by level,ID "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��;����")
    
    With rsTmp
        If .EOF Then
            Exit Sub
        End If
        
        '��ҩƷ��;��������װ��
        Do While Not .EOF
            Intĩ�� = IIf(!ĩ�� = 1, 3, 2)
            If IsNull(!�ϼ�ID) Then
                Set nodTmp = tvw���.Nodes.Add("Root" & !����, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
            Else
                Set nodTmp = tvw���.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
            End If
            nodTmp.Tag = !����   '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With

    With tvw���
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Intĩ�� = 1
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Intĩ�� = 2
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Intĩ�� = 3
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Intĩ�� = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdJiXin_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿid As Long
    
    lvw����.Left = txtJiXing.Left
    lvw����.Top = txtJiXing.Top + txtJiXing.Height
    lvw����.Visible = True
    lvw����.SetFocus
    
    On Error GoTo errHandle
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng�ⷿid <> 0 Then
        '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
        gstrSQL = "Select Distinct J.����,J.���� " & _
                  "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                  "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] " & _
                  "Order by J.���� "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿid)
    Else
        gstrSQL = "Select ����,���� From ҩƷ���� order by ���� "
        Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����")
    End If
    
    With rsTmp
        lvw����.ListItems.Clear
        Do While Not .EOF
            lvw����.ListItems.Add , "K" & !����, !����, 1, 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "����", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt��Ӧ��.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt��Ӧ��.SetFocus
    txt��Ӧ��.Tag = rsProvider!id
    txt��Ӧ��.Text = rsProvider!����
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    Dim δ��������� As String
    Dim ����������� As String
    
    '��ʼ׼��
    intNO = 21
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '�������
    If chkClass.Value = 1 Then
        If txtClass.Tag = "" Then
            MsgBox "��ѡ��Ҫ��ѯ�ķ�����Ϣ��", vbInformation, gstrSysName
            Me.txtClass.SetFocus
            Exit Sub
        End If
    End If
    If chkJiXin.Value = 1 Then
        If txtJiXing.Tag = "" Then
            MsgBox "��ѡ��Ҫ��ѯ�ļ�����Ϣ��", vbInformation, gstrSysName
            Me.txtJiXing.SetFocus
            Exit Sub
        End If
    End If
    If ChkҩƷ.Value = 1 Then
        If TxtҩƷ.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
            Me.TxtҩƷ.SetFocus
            Exit Sub
        End If
    End If
    If Chk��Ӧ��.Value = 1 Then
        If txt��Ӧ��.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
            Me.txt��Ӧ��.SetFocus
            Exit Sub
        End If
    End If
    If Chk������.Value = 1 Then
        If txt������.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��������Ϣ��", vbInformation, gstrSysName
            Me.txt������.SetFocus
            Exit Sub
        End If
    End If
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "�Բ��𣬱���ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '������ѯ����
    Dim i As Integer
    
    'δ���������
    If chkNOVerifyBack = 0 Then '����ѡֻ��ʾ���ģ���ѡ���˿�Ҳ��ʾ
       δ��������� = " and nvl(a.��ҩ��ʽ,0)=0 "
    End If
    '�����������
    If chkStrike.Value = 1 Then
        If chkAccStrike.Value = 0 And chkAcc.Value = 1 Then 'δ�������
            ����������� = " And Nvl(A.����ID,0)<>1 "
        ElseIf chkAccStrike.Value = 1 And chkAcc.Value = 0 Then  '�Ѳ������
            ����������� = " And Nvl(A.����ID,0)<>0  "
        End If
    Else
        If chkAcc.Value = 1 Then    'δ�������
            ����������� = " And Nvl(A.����ID,0)=0 "
        End If
        ����������� = ����������� & " And a.��¼״̬ =1 "
    End If
    If chk�ѱ��.Value = 1 And chkδ���.Value = 0 Then
        ����������� = ����������� & " And d.�����־ =1"
        SQLCondition.intδ��� = 0
        SQLCondition.int�ѱ�� = 1
    ElseIf chkδ���.Value = 1 And chk�ѱ��.Value = 0 Then
        ����������� = ����������� & " And d.�����־ <>1"
        SQLCondition.intδ��� = 1
        SQLCondition.int�ѱ�� = 0
    End If
    If chkYesVerifyBack.Value = 0 Then
        ����������� = ����������� & " and nvl(a.��ҩ��ʽ,0)=0 "
    End If
    
    SQLCondition.int�������һ����ѯ = 0
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        SQLCondition.int�������һ����ѯ = 1

        mstrFind = "  and ((A.�������� between [3] and [4] and A.������� is null " & δ��������� & ") or (a.������� between [5] and [6] " & ����������� & "))"
        
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        mstrFind = " And A.������� Between [5] And [6] " & �����������
            
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [3] And [4]) and A.������� is null " & δ���������
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    If chk�з�Ʊ.Value = 1 And chk�޷�Ʊ.Value = 0 Then
        mstrFind = mstrFind & " And d.��Ʊ�� is not null"
        SQLCondition.int�з�Ʊ = 1
        SQLCondition.int�޷�Ʊ = 0
    ElseIf chk�޷�Ʊ.Value = 1 And chk�з�Ʊ.Value = 0 Then
        mstrFind = mstrFind & " And d.��Ʊ�� is null"
        SQLCondition.int�з�Ʊ = 0
        SQLCondition.int�޷�Ʊ = 1
    End If
        
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿid)
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿid)
    End If
    
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2]"
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
    
    SQLCondition.strNO��ʼ = Me.txt��ʼNo
    SQLCondition.strNO���� = Me.txt����NO
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date���ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date���ʱ����� = CDate(Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59")
    
    
    '��չ��ѯ����
    SQLCondition.lngҩƷ���� = 0
    SQLCondition.str���� = ""
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If ChkҩƷ.Value = 1 Then
        lngҩƷid = TxtҩƷ.Tag
        mstrFind = mstrFind & " And A.ҩƷID + 0 =[7] "
    End If
    If Chk��Ӧ��.Value = 1 Then mstrFind = mstrFind & " And A.��ҩ��λID + 0 =[11] "
    If Chk������.Value = 1 Then mstrFind = mstrFind & " And A.����=[12] "
    
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like [10] "
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like [9] "
    If Me.Txt��ʼ��Ʊ�� <> "" And Me.Txt������Ʊ�� <> "" Then mstrFind = mstrFind & " And d.��Ʊ�� >= [13] And d.��Ʊ�� <=[14] "
    If Me.Txt��ʼ��Ʊ�� <> "" And Me.Txt������Ʊ�� = "" Then mstrFind = mstrFind & " And d.��Ʊ�� >= [13] "
    If Me.Txt��ʼ��Ʊ�� = "" And Me.Txt������Ʊ�� <> "" Then mstrFind = mstrFind & " And d.��Ʊ�� <= [14] "
        
    If chkClass.Value = 1 Then
        SQLCondition.lngҩƷ���� = Val(txtClass.Tag)
    End If
        
    If chkJiXin.Value = 1 Then
        SQLCondition.str���� = txtJiXing.Tag
    End If
    If chk��Ʊ����.Value = 1 Then
        SQLCondition.date��Ʊ������ڿ�ʼ = CDate(Format(dtpStart��Ʊ.Value, "yyyy-mm-dd") & " 00:00:00")
        SQLCondition.date��Ʊ������ڽ��� = CDate(Format(dtpEnd��Ʊ.Value, "yyyy-mm-dd") & " 23:59:59")
        mstrFind = mstrFind + " and d.������� between [19] and [20]"
    End If
    
    SQLCondition.lngҩƷ = Val(TxtҩƷ.Tag)
    SQLCondition.lng������ = txt��Ӧ��.Tag
    SQLCondition.str���� = txt������
    SQLCondition.str����� = Me.Txt����� & "%"
    SQLCondition.str������ = Me.Txt������ & "%"
    SQLCondition.str��Ʊ�ſ�ʼ = Me.Txt��ʼ��Ʊ��
    SQLCondition.str��Ʊ�Ž��� = Me.Txt������Ʊ��
    
    Unload Me
End Sub

Private Sub Cmd������_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt������.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� as id ,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ���� "
'    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "ҩƷ������", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt������.SetFocus: Exit Sub '��ѡ����ʱ����Esc�������´���
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt������.SetFocus
    txt������.Tag = 1
    txt������.Text = rsProvider!����

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "ҩƷ�⹺������", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)

'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷID
        
    If Chk��Ӧ��.Visible = True Then
        Chk��Ӧ��.SetFocus
    End If
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp����ʱ��(Index).SetFocus
End Sub


Private Sub Form_Activate()
    SQLCondition.intδ��� = 0
    SQLCondition.int�ѱ�� = 0
    SQLCondition.int�޷�Ʊ = 0
    SQLCondition.int�з�Ʊ = 0
    
    If gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
        chk�ѱ��.Visible = True
        chkδ���.Visible = True
        chk�з�Ʊ.Top = chk�ѱ��.Height + chk�ѱ��.Top + 70
        chk�޷�Ʊ.Top = chk�з�Ʊ.Top
    Else
        chk�ѱ��.Visible = False
        chkδ���.Visible = False
        chk�з�Ʊ.Top = chk�ѱ��.Top
        chk�޷�Ʊ.Top = chk�з�Ʊ.Top
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    
    Me.dtp����ʱ��(0) = Sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    Me.txt��Ӧ��.Tag = 0
    Me.TxtҩƷ.Tag = 0
    Me.txt������.Tag = 0
    lngҩƷid = 0
    
    sstFilter.Tab = 0
    BlnAdvance = False
    chk�ѱ��.Enabled = False
    chkδ���.Enabled = False
    mstrMatch = IIf(zlDataBase.GetPara("����ƥ��", , , 0) = "0", "%", "")
    
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    dtpStart��Ʊ.Value = DateAdd("m", -1, CDate(StrToday))
    dtpEnd��Ʊ.Value = CDate(StrToday)
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    gstrSQL = "Select id,�ϼ�ID,����,����,ĩ��,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And ���� is Not NULL " & _
              "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is NULL Connect by prior id=�ϼ�id"
    Set rsCompete = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ӧ��", gstrNodeNo)
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "ҩƷ��Ӧ����Ϣ��ȫ�����ڹ�ҩ��λ����������ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    gstrSQL = "Select ����,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null "
    Set rsCompete = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩƷ������", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "ҩƷ��������Ϣ��ȫ,�����ֵ����������ҩƷ��������Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            
            Case "Booker"
                Txt������.SetFocus
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
            Case "Verify"
                Txt�����.SetFocus
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
        End Select
        Cancel = True
    End If
    Call ReleaseSelectorRS
End Sub

Private Sub lvw����_DblClick()
    Dim i As Integer
    Dim strName As String
    
    With lvw����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked = True Then
                strName = strName & .ListItems(i).Text & ","
            End If
        Next
        lvw����.Visible = False
        txtJiXing.Tag = strName
        txtJiXing.Text = strName
    End With
End Sub

Private Sub lvw����_LostFocus()
    lvw����.Visible = False
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Booker"
                    Txt������ = .TextMatrix(.Row, 2)
                    Txt�����.SetFocus
                Case "Verify"
                    Txt����� = .TextMatrix(.Row, 2)
                    Txt��ʼ��Ʊ��.SetFocus
                
            End Select
            .Visible = False
            
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            ChkҩƷ.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            ChkҩƷ.SetFocus
        End If
    End If
    
End Sub

Private Sub tvw���_DblClick()
    With tvw���
        If .SelectedItem.Text <> "" Then
            If .SelectedItem.Key Like "Root*" Then Exit Sub
            txtClass.Tag = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "_") + 1)
            txtClass.Text = .SelectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub tvw���_LostFocus()
    tvw���.Visible = False
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = 100
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿid As Long
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(Trim(txtClass.Text))
        If strTemp <> "" Then
            tvw���.Left = txtClass.Left
            tvw���.Top = txtClass.Top + txtClass.Height
            tvw���.Visible = True
            tvw���.SetFocus
            
            gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
                      "Where Instr([1], ����, 1) > 0 " & _
                      "Order by ���� "
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
            With tvw���
                .Nodes.Clear
                Do While Not rsTmp.EOF
                    Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!����, rsTmp!����, 2, 2)
                    nodTmp.Tag = "Root" & rsTmp!����
                    rsTmp.MoveNext
                Loop
                rsTmp.Close
            End With
            
            gstrSQL = "Select ID, �ϼ�id, ����, 1 As ĩ��, ����, ����" & _
                        " From (Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, ����" & _
                               " From ���Ʒ���Ŀ¼" & _
                               " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                     " (���� Like [1] Or ���� Like [1] Or ���� Like [1])" & _
                               " Start With �ϼ�id Is Null" & _
                               " Connect By Prior ID = �ϼ�id" & _
                               " Union " & _
                               " Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, ����" & _
                               " From ���Ʒ���Ŀ¼" & _
                               " Where ID In (Select �ϼ�id" & _
                                            " From ���Ʒ���Ŀ¼" & _
                                            " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                                  " (���� Like [1] Or ���� Like [1] Or ���� Like [1])))" & _
                        " Start With �ϼ�id Is Null" & _
                        " Connect By Prior ID = �ϼ�id" & _
                        " Order By Level, ID"
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & strTemp & mstrMatch)
            
            With rsTmp
                If .EOF Then
                    Exit Sub
                End If
                
                '��ҩƷ��;��������װ��
                Do While Not .EOF
                    Intĩ�� = IIf(!ĩ�� = 1, 3, 2)
                    If IsNull(!�ϼ�ID) Then
                        Set nodTmp = tvw���.Nodes.Add("Root" & !����, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
                    Else
                        Set nodTmp = tvw���.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !id, !����, Intĩ��, Intĩ��)
                    End If
                    nodTmp.Tag = !����   '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
                    .MoveNext
                Loop
            End With
        
            With tvw���
                .Nodes(1).Selected = True
                If .Nodes(1).Children <> 0 Then
                    Intĩ�� = 1
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(2).Children <> 0 Then
                    Intĩ�� = 2
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(3).Children <> 0 Then
                    Intĩ�� = 3
                    .Nodes(Intĩ��).Child.Selected = True
                    .SelectedItem.Selected = True
                Else
                    Intĩ�� = 0
                    .Nodes(1).Selected = True
                    .SelectedItem.Selected = True
                End If
                If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        txtClass.Tag = 0
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtJiXing_GotFocus()
    txtJiXing.SelStart = 0
    txtJiXing.SelLength = 100
End Sub

Private Sub txtJiXing_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿid As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind = "" Then Exit Sub
        
        lvw����.Left = txtJiXing.Left
        lvw����.Top = txtJiXing.Top + txtJiXing.Height
        lvw����.Visible = True
        lvw����.SetFocus
        
        On Error GoTo errHandle
        lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If lng�ⷿid <> 0 Then
            '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
            gstrSQL = "Select Distinct J.����,J.���� " & _
                      "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                      "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] and (j.���� like [2] or j.���� like [2] or j.���� like [2]) " & _
                      "Order by J.���� "
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿid, "%" & strFind & mstrMatch)
        Else
            gstrSQL = "Select ����,���� From ҩƷ���� where ���� like [1] or ���� like [1] or ���� like [1] order by ���� "
            Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����", "%" & strFind & mstrMatch)
        End If
        
        With rsTmp
            lvw����.ListItems.Clear
            Do While Not .EOF
                lvw����.ListItems.Add , "K" & !����, !����, 1, 1
                .MoveNext
            Loop
        End With
    ElseIf KeyCode = vbKeyDelete Then
        txtJiXing.Tag = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt��Ӧ��_GotFocus()
'    Tvw.Visible = False
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt��Ӧ��)) <> "" Then
        txt��Ӧ�� = UCase(txt��Ӧ��)
        vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd)

        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [2] Or վ�� is Null) " & _
                  "  And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] or ���� like [1] or ���� like [1]) "
'        Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ҩƷ��Ӧ��]", IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%", gstrNodeNo)
        Set RecTmp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%", , gstrNodeNo)
        
        If blnCancel Then txt��Ӧ��.SetFocus: Exit Sub
        
        If RecTmp.State = 0 Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            KeyCode = 0
            txt��Ӧ��.Tag = 0
            txt��Ӧ��.SelStart = 0
            txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
            Exit Sub
        End If
        
        txt��Ӧ�� = RecTmp!����
        txt��Ӧ��.Tag = RecTmp!id
        
    End If
    
    If Chk������.Value = 1 Then
        txt������.SetFocus
    Else
        Chk������.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = 21
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿid)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.cmdȷ��.SetFocus
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = 21
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿid)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt��ʼ��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt������Ʊ��.SetFocus
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Txt��ʼ��Ʊ��.SetFocus
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            Txt��ʼ��Ʊ��.SetFocus
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)
        
        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", _
                        Me.Txt����� & "%", gstrNodeNo)
        
        With rstemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rstemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt�����.Top + Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra��������.Top - Txt�����.Top - Txt�����.Height - 50
                    .Width = Me.ScaleWidth - sstFilter.Left - fra��������.Left - Txt�����.Left - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt����� = IIf(IsNull(!����), "", !����)
                Txt��ʼ��Ʊ��.SetFocus
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

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt������ = "" Then Exit Sub
        If Trim(txt������) = "" Then Exit Sub
        txt������ = UCase(txt������)
        vRect = zlControl.GetControlRect(txt������.hWnd)
    
        Dim rstemp As New ADODB.Recordset

        gstrSQL = "Select ���� as id,����,���� From ҩƷ������ " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) Order By ����"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ҩƷ������]", _
'                        IIf(gstrMatchMethod = "0", "%", "") & Me.txt������ & "%", _
'                        Me.txt������ & "%", gstrNodeNo)
        Set rstemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & Me.txt������ & "%", Me.txt������ & "%", gstrNodeNo)
        
        If blnCancel Then txt������.SetFocus: Exit Sub
        
        If rstemp.State = 0 Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            KeyCode = 0
            txt������.Tag = 0
            txt������.SelStart = 0
            txt������.SelLength = Len(txt������.Text)
            Exit Sub
        End If
        
        txt������ = IIf(IsNull(rstemp!����), "", rstemp!����)
        
        txt������.Tag = 1
        Txt������.SetFocus

    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As New ADODB.Recordset
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            Txt�����.SetFocus
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", _
                        Me.Txt������ & "%", gstrNodeNo)
        
        With rstemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rstemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt������.Top + Txt������.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt������.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra��������.Top - Txt������.Top - Txt������.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt������ = IIf(IsNull(!����), "", !����)
                Me.Txt�����.SetFocus
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

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra��������.Left + TxtҩƷ.Left
    sngTop = Me.Top + sstFilter.Top + fra��������.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - TxtҩƷ.Height - 3630
    End If
    
    strkey = Trim(TxtҩƷ.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "ҩƷ�⹺������", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷID
    
    If Chk��Ӧ��.Visible = True Then
        If Chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            Chk��Ӧ��.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

