VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTransferSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4260
   ClientLeft      =   3156
   ClientTop       =   3168
   ClientWidth     =   7560
   Icon            =   "frmTransferSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1605
      Left            =   1200
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7853
      _ExtentY        =   2836
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
      Height          =   3975
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   7006
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmTransferSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmTransferSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   5520
         Begin MSComctlLib.ListView lvw���� 
            Height          =   1755
            Left            =   1560
            TabIndex        =   35
            Top             =   2640
            Visible         =   0   'False
            Width           =   3885
            _ExtentX        =   6858
            _ExtentY        =   3090
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
            Height          =   2205
            Left            =   0
            TabIndex        =   36
            Top             =   2640
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6435
            _ExtentY        =   3895
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
         Begin VB.CheckBox chkClass 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "ҩƷ����"
            Height          =   300
            Left            =   360
            TabIndex        =   33
            Top             =   680
            Width           =   1095
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ"
            Height          =   300
            Left            =   360
            TabIndex        =   32
            Top             =   1120
            Width           =   990
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            ScrollBars      =   3  'Both
            TabIndex        =   31
            Top             =   1120
            Width           =   3255
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   30
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   29
            Top             =   680
            Width           =   255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   28
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   27
            Top             =   680
            Width           =   3255
         End
         Begin VB.CheckBox Chk����ⷿ 
            Caption         =   "����ⷿ"
            Height          =   420
            Left            =   360
            TabIndex        =   9
            Top             =   1500
            Width           =   1110
         End
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   8
            Top             =   1120
            Width           =   255
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   11
            Top             =   2100
            Width           =   1845
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   12
            Top             =   2460
            Width           =   1845
         End
         Begin VB.ComboBox Cbo�ⷿ 
            Enabled         =   0   'False
            Height          =   276
            Left            =   1530
            TabIndex        =   10
            Text            =   "Cbo�ⷿ"
            Top             =   1560
            Width           =   3550
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   255
            Left            =   570
            TabIndex        =   18
            Top             =   2123
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   570
            TabIndex        =   19
            Top             =   2520
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   39
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "����˳���"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   38
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "δ��˳���"
            Height          =   300
            Left            =   720
            TabIndex        =   37
            Top             =   1400
            Visible         =   0   'False
            Width           =   1335
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
            Top             =   1680
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   312
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   312
            Index           =   1
            Left            =   3588
            TabIndex        =   7
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   15
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
            TabIndex        =   24
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
            TabIndex        =   17
            Top             =   2028
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   3348
            TabIndex        =   23
            Top             =   2028
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   16
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
            TabIndex        =   22
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
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
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTransferSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '�����ַ���
Private BlnAdvance As Boolean '�Ƿ�չ��
Private mlngMode As Long    '��������
Private mdatStart As Date   '��ʼʱ��
Private mdatEnd As Date     '����ʱ��
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '������
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mblnStock As Boolean            '��ǰ����Ա�Ƿ���ҩ����Ա���������õ�����Ч
Private mint������� As Integer
Private mlngStoreId As Long     '��ǰ�ⷿid
Private mstrMatch As String 'ƥ�䷽ʽ 0-˫��ƥ�� 1-�������ҵ���ƥ��
Private mint�������� As Integer '0-����Ҫ����;1-��Ҫ����

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng�ⷿ As Long
    str������ As String
    str����� As String
    int�������һ����ѯ As Integer
    lngҩƷ���� As Long
    str���� As String
End Type

Private SQLCondition As Type_SQLCondition

Public Property Get In_�������() As Integer
    In_������� = mint�������
End Property

Public Property Let In_�������(ByVal vNewValue As Integer)
    mint������� = vNewValue
End Property
Private Function Check�Ƿ���ҩ����Ա() As Boolean
    Dim rsDepend As ADODB.Recordset
    
    On Error GoTo errHandle
    '�ж��ǲ���ҩ����Աʹ�ñ�ģ��
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [2] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr('HIJKLMN', b.����, 1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " _
            & "  And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) "
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.�û�ID, gstrNodeNo)
                  
    Check�Ƿ���ҩ����Ա = (rsDepend.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lngStoreid As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO��ʼ As String, _
        ByRef strNO���� As String, _
        ByRef date����ʱ�俪ʼ As Date, _
        ByRef date����ʱ����� As Date, _
        ByRef date���ʱ�俪ʼ As Date, _
        ByRef date���ʱ����� As Date, _
        ByRef lngҩƷ As Long, _
        ByRef lng�ⷿ As Long, _
        ByRef str������ As String, _
        ByRef str����� As String, _
        ByRef lngҩƷ���� As Long, _
        ByRef str���� As String, _
        Optional ByRef intTmp As Integer = 0) As String
    mstrFind = ""
    mlngMode = lngMode
    mlngStoreId = lngStoreid
    Set mfrmMain = FrmMain
    
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
    lng�ⷿ = SQLCondition.lng�ⷿ
    str����� = SQLCondition.str�����
    str������ = SQLCondition.str������
    lngҩƷ���� = SQLCondition.lngҩƷ����
    str���� = SQLCondition.str����
    intTmp = SQLCondition.int�������һ����ѯ
    
End Function

Private Sub Cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    '��ȡ�ɲ����Ŀⷿ
    Select Case mlngMode
        Case ģ���.ҩƷ�ƿ�
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.ҩƷ����
            str�������� = "O"
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo�ⷿ.ListCount = 0 Then Exit Sub
    
    If Cbo�ⷿ.ListIndex >= 0 Then
        If Val(Cbo�ⷿ.Tag) = Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, Cbo�ⷿ, Trim(Cbo�ⷿ.Text), str��������) = False Then
        Exit Sub
    End If
    If Cbo�ⷿ.ListIndex >= 0 Then
        Cbo�ⷿ.Tag = Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex)
    End If
End Sub

Private Sub Cbo�ⷿ_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0

    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Cbo�ⷿ_Validate(Cancel As Boolean)
    If Cbo�ⷿ.ListCount > 0 Then
        If Cbo�ⷿ.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

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

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdȷ��.SetFocus
    End If
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chk���.Value = 1 Then
            SendKeys vbTab
        Else
            cmdȷ��.SetFocus
        End If
    End If
    
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub ChkҩƷ_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        ChkҩƷ.SetFocus
    End If
    
End Sub

Private Sub ChkҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk����ⷿ_click()
    Cbo�ⷿ.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
End Sub

Private Sub Chk����ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk����ⷿ.Value = 1 Then
        Cbo�ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
End Sub
Private Sub chk����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    chkNoStrike.Enabled = IIf(chk����.Value = 1, True, False)
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
    chkYesStrike.Enabled = IIf(chk���.Value = 1, True, False)
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    CmdҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
End Sub



Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿID As Long
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    tvw���.Left = txtClass.Left
    tvw���.Top = txtClass.Top + txtClass.Height
    tvw���.Visible = True
    tvw���.SetFocus
        
    gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
              "Where Instr([1], ����, 1) > 0 " & _
              "Order by ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
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
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��;����")
    
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
    Dim lng�ⷿID As Long
    
    lvw����.Left = txtJiXing.Left
    lvw����.Top = txtJiXing.Top + txtJiXing.Height
    lvw����.Visible = True
    lvw����.SetFocus
    
    On Error GoTo errHandle
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng�ⷿID <> 0 Then
        '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
        gstrSQL = "Select Distinct J.����,J.���� " & _
                  "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                  "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] " & _
                  "Order by J.���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿID)
    Else
        gstrSQL = "Select ����,���� From ҩƷ���� order by ���� "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����")
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

Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '�������
    If chkClass.Value = 1 Then
        If txtClass.Tag = 0 Then
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
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "�Բ��𣬱���ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '������ѯ����
    Dim i As Integer
    
    SQLCondition.int�������һ����ѯ = 0
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        SQLCondition.int�������һ����ѯ = 1
        If mlngMode <> 1304 Then 'ҩƷ�ƿ�
            If chkStrike.Value = 1 Then
                If mlngMode <> 1306 Then
                    mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                            & " or (A.������� Between [5] And [6]))"
                Else
                    mstrFind = " And 1=1 "
                End If
            Else
                If mlngMode <> 1306 Then
                    mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                            & " or (A.������� Between [5] And [6] and a.��¼״̬ =1))  "
                Else
                    mstrFind = " And a.��¼״̬ =1 "
                End If
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                        & " or (A.������� Between [5] And [6]))"
            Else
                If chkNoStrike.Value = 1 And chkYesStrike.Value = 1 Then
                    mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                                & " or (A.������� Between [5] And [6]))"
                ElseIf chkNoStrike.Value = 1 And chkYesStrike.Value = 0 Then
                    mstrFind = "And ((((Mod(a.��¼״̬, 3) = 0 And ������� Is not Null) or (Mod(a.��¼״̬, 3) = 2  And ������� Is Null)) and �������� Between [3] and [4] " _
                               & "And Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2 And b.������� Is Null)) " _
                               & " or (A.������� Between [5] And [6] and (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0)" _
                               & "And Not Exists (Select 1 From ҩƷ�շ���¼ Y Where a.���� = y.���� and a.�ⷿid =y.�ⷿid and a.No = y.No And Mod(y.��¼״̬, 3) = 2)))"
                ElseIf chkNoStrike.Value = 0 And chkYesStrike.Value = 1 Then
                    mstrFind = " and ((A.��¼״̬=2 or mod(A.��¼״̬,3)=2 or mod(A.��¼״̬,3)=0) And A.������� Between [5] And [6] " _
                               & "And Not Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2 And b.������� Is Null) " _
                               & " or ((a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) and (A.�������� Between [3] And [4]) and a.������� is null)) "
                Else
                    mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                                & " or (A.������� Between [5] And [6])) and (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) " _
                                & "And Not Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2)"
                End If
            End If
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If mlngMode <> 1304 Then 'ҩƷ�ƿ�
            If chkStrike.Value = 1 Then
                mstrFind = " And A.������� Between [5] And [6] "
            Else
                mstrFind = " And A.������� Between [5] And [6] and a.��¼״̬ =1 "
                
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And A.������� Between [5] And [6] "
            Else
                If chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.��¼״̬=2 or mod(A.��¼״̬,3)=2 or mod(A.��¼״̬,3)=0) And A.������� Between [5] And [6] " _
                               & "And Not Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2 And b.������� Is Null)"
                Else
                    mstrFind = " And A.������� Between [5] And [6] and (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) " _
                               & "And Not Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2)"
                End If
            End If
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        If mlngMode <> 1304 Then 'ҩƷ�ƿ�
            mstrFind = " And (A.�������� Between [3] And [4]) and ������� is null "
        Else
            If chkNoStrike.Value = 1 Then
                mstrFind = "And ((Mod(a.��¼״̬, 3) = 0 And ������� Is not Null) or (Mod(a.��¼״̬, 3) = 2  And ������� Is Null)) and �������� Between [3] and [4] " _
                           & "And Exists (Select 1 From ҩƷ�շ���¼ B Where a.���� = b.���� and a.�ⷿid =b.�ⷿid and a.No = b.No And Mod(b.��¼״̬, 3) = 2 And b.������� Is Null)"
            Else
                mstrFind = " And (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) and (A.�������� Between [3] And [4]) and ������� is null "
            End If
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿID)
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿID)
    End If
    
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2] "
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
        mstrFind = mstrFind & " And A.ҩƷID + 0 =[7] "
    End If
    
    If mlngMode = ģ���.�������� Then
        If Chk����ⷿ.Value = 1 Then mstrFind = mstrFind & " And A.������ID=[8]"
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Then
        If mint������� = -1 Then
            If Chk����ⷿ.Value = 1 Then mstrFind = mstrFind & " And A.�Է�����ID + 0 =[8]"
        Else
            If Chk����ⷿ.Value = 1 Then mstrFind = mstrFind & " And A.�ⷿID+0=[8]"
        End If
    Else
        If Chk����ⷿ.Value = 1 Then mstrFind = mstrFind & " And A.�Է�����ID + 0 =[8]"
    End If
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like [10] "
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like [9] "
    
    If chkClass.Value = 1 Then
        SQLCondition.lngҩƷ���� = Val(txtClass.Tag)
    End If
    If chkJiXin.Value = 1 Then
        SQLCondition.str���� = txtJiXing.Tag
    End If
    SQLCondition.lngҩƷ = Val(TxtҩƷ.Tag)
    If Cbo�ⷿ.Visible Then
        SQLCondition.lng�ⷿ = Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex)
    End If
    SQLCondition.str����� = Me.Txt����� & "%"
    SQLCondition.str������ = Me.Txt������ & "%"
    
    Unload Me
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "ҩƷ�ƿ����", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    If Chk����ⷿ.Visible = True Then
        Chk����ⷿ.SetFocus
    Else
        Txt������.SetFocus
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

Private Sub Form_Load()
    Me.dtp����ʱ��(0) = Sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    mblnStock = Check�Ƿ���ҩ����Ա
    mstrMatch = IIf(zlDatabase.GetPara("����ƥ��", , , 0) = "0", "%", "")
    
    Me.TxtҩƷ.Tag = 0
    sstFilter.Tab = 0
    Select Case mlngMode
        Case ģ���.ҩƷ�ƿ�
            mint�������� = Val(zlDatabase.GetPara("��������", glngSys, ģ���.ҩƷ�ƿ�))
            If mint������� = -1 Then
                Chk����ⷿ.Caption = "����ⷿ"
            Else
                Chk����ⷿ.Caption = "�Ƴ��ⷿ"
            End If
            If mint�������� = 0 Then    '����Ҫ����
                chkStrike.Visible = True
                chkNoStrike.Visible = False
                chkYesStrike.Visible = False
            Else
                chkStrike.Visible = False
                chkNoStrike.Visible = True
                chkYesStrike.Visible = True
            End If
        Case ģ���.ҩƷ����
            Chk����ⷿ.Caption = "���ò���"
        Case ģ���.��������
            Chk����ⷿ.Caption = "������"
    End Select
    BlnAdvance = False
End Sub


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

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsDepartment As New Recordset
    Dim strStock As String
    Dim strվ������ As String
    
    On Error GoTo errHandle
    strվ������ = GetDeptStationNode(mlngStoreId)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            If Cbo�ⷿ.ListCount < 1 Then
                Select Case mlngMode
                    Case 1304
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                                & "Where " & IIf(strվ������ <> "", " (a.վ�� = [3] or a.վ�� is null) AND ", "") & " c.�������� = b.���� " _
                                & "  AND Instr([1],b.����,1) > 0 " _
                                & "  AND a.id = c.����id " _
                                & "  AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1305
                        strStock = "O"
                        gstrSQL = " Select C.ID " & _
                            " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                            " Where " & IIf(strվ������ <> "", " (c.վ�� = [3] or c.վ�� is null) AND ", "") & " A.��������=B.���� And A.����ID=C.ID " & _
                            "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                            "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])"
                        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                            & "Where " & IIf(strվ������ <> "", " (a.վ�� = [3] or a.վ�� is null) AND ", "") & " c.�������� = b.���� " _
                            & "  AND Instr([1],b.����,1) > 0 " _
                            & "  AND a.id = c.����id " _
                            & "  AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')" _
                            & IIf(mblnStock, "", " And a.ID IN (Select Distinct ���ò���ID From ҩƷ���ÿ��� Where ���ò���ID IN (" & gstrSQL & "))")
                    Case 1306
                       gstrSQL = "SELECT b.Id,b.���� " _
                               & "FROM ҩƷ�������� A, ҩƷ������ B " _
                               & "Where A.���id = B.ID AND A.���� = 11 "
                    Case 1303, 1307
                        If Chk����ⷿ.Visible = True Then
                            Chk����ⷿ.Visible = False
                            Cbo�ⷿ.Visible = False
                        End If
                        Exit Sub
                End Select
                Set rsDepartment = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock, UserInfo.�û�ID, gstrNodeNo)
            
                With Cbo�ⷿ
                    Do While Not rsDepartment.EOF
                        .AddItem rsDepartment.Fields(1)
                        .ItemData(.NewIndex) = rsDepartment.Fields(0)
                        rsDepartment.MoveNext
                    Loop
                    If .ListCount > 0 Then .ListIndex = 0
                End With
                rsDepartment.Close
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
    Dim lng�ⷿID As Long
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
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
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
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & strTemp & mstrMatch)
            
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
    Dim lng�ⷿID As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind = "" Then Exit Sub
        
        lvw����.Left = txtJiXing.Left
        lvw����.Top = txtJiXing.Top + txtJiXing.Height
        lvw����.Visible = True
        lvw����.SetFocus
        
        On Error GoTo errHandle
        lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If lng�ⷿID <> 0 Then
            '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
            gstrSQL = "Select Distinct J.����,J.���� " & _
                      "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                      "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] and (j.���� like [2] or j.���� like [2] or j.���� like [2]) " & _
                      "Order by J.���� "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿID, "%" & strFind & mstrMatch)
        Else
            gstrSQL = "Select ����,���� From ҩƷ���� where ���� like [1] or ���� like [1] or ���� like [1] order by ���� "
            Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����", "%" & strFind & mstrMatch)
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

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If mfrmMain.TabShow.Tab = 1 Then
            '�̵��
            intNO = 29
        Else
            '�̵��¼��
            intNO = 62
        End If
    End If
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If mfrmMain.TabShow.Tab = 1 Then
            '�̵��
            intNO = 29
        Else
            '�̵��¼��
            intNO = 62
        End If
    End If
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿID)
        End If
        Me.txt����NO.SetFocus
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then cmdȷ��.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            cmdȷ��.SetFocus
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", _
                        Me.Txt����� & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt�����.Top + Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra��������.Top - Txt�����.Top - Txt�����.Height - 50
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
                cmdȷ��.SetFocus
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


Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt�����.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", _
                        Me.Txt������ & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
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

Private Sub Txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra��������.Left + TxtҩƷ.Left
    sngTop = Me.Top + fra��������.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
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
    
    Call SetSelectorRS(1, "ҩƷ�ƿ����", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    If Chk����ⷿ.Visible = True Then
        Chk����ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
    
End Sub

Private Sub TxtҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
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
                    cmdȷ��.SetFocus
                
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

