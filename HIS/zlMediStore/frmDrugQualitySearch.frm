VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDrugQualitySearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4260
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7410
   Icon            =   "frmDrugQualitySearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2175
      Left            =   720
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
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
      TabIndex        =   21
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmDrugQualitySearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmDrugQualitySearch.frx":0028
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
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��ҩ��λ(&S)"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   960
            Width           =   1350
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   10
            Top             =   960
            Width           =   3255
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   11
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   8
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   7
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ(&P)"
            Height          =   300
            Left            =   480
            TabIndex        =   6
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1740
            Width           =   1845
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   13
            Top             =   2220
            Width           =   1845
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ǽ���"
            Height          =   180
            Left            =   1170
            TabIndex        =   18
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   1170
            TabIndex        =   19
            Top             =   2280
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chk���� 
            Caption         =   "δ������"
            Height          =   300
            Left            =   480
            TabIndex        =   0
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "�Ѵ�����"
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   1560
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   960
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   2
            Top             =   960
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   4
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   5
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   17
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
            TabIndex        =   24
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ǽ�����"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   16
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   23
            Top             =   1020
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6210
      TabIndex        =   15
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6210
      TabIndex        =   14
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmDrugQualitySearch"
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
Private mlng�ⷿID As Long  '�ⷿid

Private Type Type_SQLCondition
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng��Ӧ�� As Long
    str������ As String
    str����� As String
End Type

Private SQLCondition As Type_SQLCondition
Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef date����ʱ�俪ʼ As Date, _
        ByRef date����ʱ����� As Date, _
        ByRef date���ʱ�俪ʼ As Date, _
        ByRef date���ʱ����� As Date, _
        ByRef lngҩƷ As Long, _
        ByRef lng��Ӧ�� As Long, _
        ByRef str������ As String, _
        ByRef str����� As String, _
        ByVal lng�ⷿID As Long) As String
    mstrFind = ""
    mlng�ⷿID = lng�ⷿID
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd

    date����ʱ�俪ʼ = SQLCondition.date����ʱ�俪ʼ
    date����ʱ����� = SQLCondition.date����ʱ�����
    date���ʱ�俪ʼ = SQLCondition.date���ʱ�俪ʼ
    date���ʱ����� = SQLCondition.date���ʱ�����
    lngҩƷ = SQLCondition.lngҩƷ
    lng��Ӧ�� = SQLCondition.lng��Ӧ��
    str����� = SQLCondition.str�����
    str������ = SQLCondition.str������
    
End Function



Private Sub Chk��Ӧ��_Click()
    txt��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    Cmd��Ӧ��.Enabled = txt��Ӧ��.Enabled
End Sub

Private Sub Chk��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
End Sub


Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    cmdҩƷ.Enabled = TxtҩƷ.Enabled
    
End Sub

Private Sub ChkҩƷ_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        ChkҩƷ.SetFocus
    End If
    
    
End Sub

Private Sub ChkҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� " & _
              "From ��Ӧ�� Where (վ�� = [1] Or վ�� is Null) " & _
              " And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID Order by level,ID"
'    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ӧ��", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "��Ӧ��", True, "", "", False, False, _
            True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
 
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
    '�������
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "�Բ��𣬱���ѡ��һ���Ǽ����ڻ��ߴ�������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
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
    
    
    mstrFind = ""
    '������ѯ����
    Dim i As Integer
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        mstrFind = " And ((A.�Ǽ�ʱ�� Between [1] And [2] and ����ʱ�� is null) " _
                    & " or (A.����ʱ�� Between [3] And [4]))"
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        mstrFind = " And A.����ʱ�� Between [3] And [4] "
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�Ǽ�ʱ�� Between [1] And [2]) and ����ʱ�� is null "
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date���ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date���ʱ����� = CDate(Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59")
    
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If ChkҩƷ.Value = 1 Then mstrFind = mstrFind & " And A.ҩƷid=[5] "
    If Chk��Ӧ��.Value = 1 Then mstrFind = mstrFind & " And A.��ҩ��λid +0 =[6]"
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.������ like [8] "
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.�Ǽ��� like [7] "
    
    SQLCondition.lngҩƷ = Val(TxtҩƷ.Tag)
    SQLCondition.lng��Ӧ�� = Val(txt��Ӧ��.Tag)
    SQLCondition.str����� = Me.Txt����� & "%"
    SQLCondition.str������ = Me.Txt������ & "%"
        
    Unload Me
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
    Dim intLop As Integer
    
    Me.dtp����ʱ��(0) = Sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    TxtҩƷ.Tag = 0
    txt��Ӧ��.Tag = 0
    
    sstFilter.Tab = 0
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
            chk����.SetFocus
        Else
            ChkҩƷ.SetFocus
        End If
    End If
    
End Sub

Private Sub txt��Ӧ��_GotFocus()
    txt��Ӧ��.SelStart = 0
    txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
End Sub

Private Sub Txt�����_GotFocus()
    Txt�����.SelStart = 0
    Txt�����.SelLength = Len(Txt�����.Text)
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmdȷ��.SetFocus
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
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", _
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

Private Sub Txt������_GotFocus()
    Txt������.SelStart = 0
    Txt������.SelLength = Len(Txt������.Text)
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Me.Txt�����.SetFocus
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
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtҩƷ_GotFocus()
    TxtҩƷ.SelStart = 0
    TxtҩƷ.SelLength = Len(TxtҩƷ.Text)
End Sub

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra��������.Left + TxtҩƷ.Left
    sngTop = Me.Top + sstFilter.Top + fra��������.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight  '  50
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
    
    Call SetSelectorRS(1, "ҩƷ��������", mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , True)
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , , UserInfo.����ID, strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If

    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    If Chk��Ӧ��.Visible = True Then
        If Chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            Chk��Ӧ��.SetFocus
        End If
    End If
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "ҩƷ��������", mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , True)

'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, , , UserInfo.����ID)
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    If Chk��Ӧ��.Visible = True Then
        If Chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            Chk��Ӧ��.SetFocus
        End If
    End If
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
                  "  And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0)" & _
                  "  And (���� like [1] or ���� like [1] or ���� like [1] )"
'        Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%", gstrNodeNo)
        Set RecTmp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%", gstrNodeNo)
        
        
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
                  
'        With RecTmp
'            If .EOF Then
'                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
'                KeyCode = 0
'                txt��Ӧ��.Tag = 0
'                txt��Ӧ��.SelStart = 0
'                txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
'
'                Exit Sub
'            End If
'            If .RecordCount > 1 Then
'                mstrSelectTag = "Provider"
'                Set mshSelect.Recordset = RecTmp
'                With mshSelect
'                    .Top = sstFilter.Top + fra��������.Top + txt��Ӧ��.Top + txt��Ӧ��.Height
'                    .Left = sstFilter.Left + fra��������.Left + txt��Ӧ��.Left
'                    .Visible = True
'                    .SetFocus
'                    .ColWidth(0) = 0
'                    .ColWidth(1) = 800
'                    .ColWidth(2) = 800
'
'                    .ColWidth(3) = .Width - .ColWidth(1) - .ColWidth(2)
'                    .Row = 1
'                    .Col = 0
'                    .ColSel = .Cols - 1
'                    Exit Sub
'
'                End With
'            Else
'                txt��Ӧ�� = !����
'                txt��Ӧ��.Tag = !id
'
'            End If
'        End With
    End If
    
    Txt������.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


