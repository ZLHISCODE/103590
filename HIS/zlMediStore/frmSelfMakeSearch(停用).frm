VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSelfMakeSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmSelfMakeSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1935
      Left            =   2040
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
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
      TabIndex        =   17
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmSelfMakeSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmSelfMakeSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   5505
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   10
            Top             =   480
            Width           =   3375
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   480
            Width           =   870
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1140
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   13
            Top             =   1740
            Width           =   1365
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   810
            TabIndex        =   27
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   810
            TabIndex        =   26
            Top             =   1800
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   5520
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
            Format          =   151519235
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
            Format          =   151519235
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
            Format          =   151519235
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
            Format          =   151519235
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
      TabIndex        =   15
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   14
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmSelfMakeSearch"
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
Public lngҩƷID As Long
Private mstrSelectTag As String     '��ǰѡ��Ķ���

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
End Type

Private SQLCondition As Type_SQLCondition

Private Type Type_TemporaryInquiries
    intδ��˵��� As Integer
    int����˵��� As Integer
    int�������� As Integer
End Type

Private TemporaryInquiries As Type_TemporaryInquiries   '��ʱ��ѯ���������ڻָ��ϴ����õĹ�����������������رպ�������������ã�

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
        ByRef intδ��˵��� As Integer, _
        ByRef int����˵��� As Integer, _
        ByRef int�������� As Integer) As String
        
    mstrFind = ""
    mstrSelectTag = ""
        
    '��ʱ��ѯ��ʼ��
    '---------------------
    SQLCondition.date����ʱ�俪ʼ = date����ʱ�俪ʼ
    SQLCondition.date����ʱ����� = date����ʱ�����
    SQLCondition.date���ʱ�俪ʼ = date���ʱ�俪ʼ
    SQLCondition.date���ʱ����� = date���ʱ�����
    
    TemporaryInquiries.intδ��˵��� = intδ��˵���
    TemporaryInquiries.int����˵��� = int����˵���
    TemporaryInquiries.int�������� = int��������
    '---------------------
        
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
    
    '��ʱ��ѯ����
    '---------------------
    intδ��˵��� = TemporaryInquiries.intδ��˵���
    int����˵��� = TemporaryInquiries.int����˵���
    int�������� = TemporaryInquiries.int��������
    '---------------------
    
End Function


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

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk���.Value = 0 Then
            cmdȷ��.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub



Private Sub chk����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
    
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
    Else
        Txt������.SetFocus
    End If
End Sub



Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '��ʼ׼��
    intNO = 22
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '�������
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
    If chk����.Value = 1 And chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                    & " or (A.������� Between [5] And [6]))"
        Else
            mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                    & " or (A.������� Between [5] And [6] and a.��¼״̬ =1))  "
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.������� Between [5] And [6] "
        Else
            mstrFind = " And A.������� Between [5] And [6] and a.��¼״̬ =1 "
            
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [3] And [4]) and ������� is null "
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
    
    TemporaryInquiries.intδ��˵��� = chk����.Value
    TemporaryInquiries.int����˵��� = chk���.Value
    TemporaryInquiries.int�������� = chkStrike.Value
    
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If ChkҩƷ.Value = 1 Then
        lngҩƷID = TxtҩƷ.Tag
        mstrFind = mstrFind & " And A.ҩƷID + 0 =[7]"
    End If
    
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like [10] "
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like [9] "
    
    SQLCondition.lngҩƷ = Val(TxtҩƷ.Tag)
    SQLCondition.str������ = Me.Txt������ & "%"
    SQLCondition.str����� = Me.Txt����� & "%"
    
    
    Unload Me
End Sub


Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "ҩƷ����������", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    Txt������.SetFocus
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
    '�ָ���һ�ε�����
    '--------------------------------
    Me.dtp����ʱ��(0) = SQLCondition.date����ʱ�����
    Me.dtp����ʱ��(1) = SQLCondition.date���ʱ�����
    Me.dtp��ʼʱ��(0) = SQLCondition.date����ʱ�俪ʼ
    Me.dtp��ʼʱ��(1) = SQLCondition.date���ʱ�俪ʼ
    
    Me.chk����.Value = TemporaryInquiries.intδ��˵���
    Me.chk���.Value = TemporaryInquiries.int����˵���
    Me.chkStrike.Value = TemporaryInquiries.int��������
    '--------------------------------
    
    Me.TxtҩƷ.Tag = 0
    lngҩƷID = 0
    
    sstFilter.Tab = 0
    BlnAdvance = False
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    
    gstrSQL = "Select ����,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ���� "
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩƷ������", gstrNodeNo)
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
        Exit Sub
    End If
    Call ReleaseSelectorRS
    
    Set mfrmMain = Nothing
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

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '��ʼ׼��
    intNO = 22
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer
    
    '��ʼ׼��
    intNO = 22
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If
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

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra��������.Left + TxtҩƷ.Left
    sngTop = Me.Top + sstFilter.Top + fra��������.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 4530 > Screen.Height Then
        sngTop = sngTop - TxtҩƷ.Height - 4530
    End If
    
    strkey = Trim(TxtҩƷ.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "ҩƷ����������", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    Txt������.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
