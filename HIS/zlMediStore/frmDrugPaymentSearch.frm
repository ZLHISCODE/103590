VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDrugPaymentSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmDrugPaymentSearch.frx":0000
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
      Height          =   2535
      Left            =   2760
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
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
      TabIndex        =   18
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmDrugPaymentSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmDrugPaymentSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   5505
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   540
            Width           =   870
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   10
            Top             =   540
            Width           =   3375
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   540
            Width           =   255
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   13
            Top             =   1860
            Width           =   1365
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   750
            TabIndex        =   28
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   750
            TabIndex        =   27
            Top             =   1920
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   19
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
            Format          =   108396547
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
            Format          =   108396547
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
            Format          =   108396547
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
            Format          =   108396547
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   25
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6330
      TabIndex        =   16
      Top             =   3720
      Width           =   1100
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
Attribute VB_Name = "FrmDrugPaymentSearch"
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
Private mstrSelectTag As String     '��ǰѡ��Ķ���

Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date) As String
    mstrFind = ""
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
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

Private Sub Chk��Ӧ��_Click()
    txt��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    Cmd��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    
End Sub

Private Sub Chk��Ӧ��_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
    End If
    Chk��Ӧ��.SetFocus
End Sub

Private Sub Chk��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
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


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ҩƷ��Ӧ�� Where " & _
                " To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01'  start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "Cmd��Ӧ��_Click")
    Call SQLTest
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "����ҩƷ��Ӧ��"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = True Then
            txt��Ӧ��.Tag = .CurrentID
            txt��Ӧ��.Text = .CurrentName
        End If
    End With
    Unload FrmSelect
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
    If Chk��Ӧ��.Value = 1 Then
        If txt��Ӧ��.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
            Me.txt��Ӧ��.SetFocus
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
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
                    & " or (A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')))"
        Else
            mstrFind = " And ((A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
                    & " or (A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS'))) and a.��¼״̬ =1 "
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') "
        Else
            mstrFind = " And A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') and a.��¼״̬ =1 "
            
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and ������� is null "
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        Me.txt����NO = UCase(LTrim(Me.txt����NO))
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
    End If
    
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt��ʼNo & "' And A.No <='" & Me.txt����NO & "'"
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt��ʼNo & "'"
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= '" & Me.txt����NO & "'"
    
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk��Ӧ��.Value = 1 Then
        mstrFind = mstrFind & " and a.��λid=" & txt��Ӧ��.Tag
    End If
    
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like '" & Me.Txt����� & "%'"
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like '" & Me.Txt������ & "%'"
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
    Me.dtp����ʱ��(0) = zldatabase.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    
    Me.dtp��ʼʱ��(0) = DateAdd("m", -1, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    Me.txt��Ӧ��.Tag = 0
    '�򿪼�¼��
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    With rsCompete
        If .State = 1 Then .Close
        gstrSQL = "Select id,�ϼ�ID,����,����,ĩ��,���� From ҩƷ��Ӧ�� Where ���� is Not NULL And " & _
              " To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' Start with �ϼ�ID is NULL Connect by prior id=�ϼ�id"
        Call SQLTest(App.Title, "����", gstrSQL)
        Set rsCompete = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
        Call SQLTest
        
        If .EOF Then
            .Close
            MsgBox "ҩƷ��Ӧ����Ϣ��ȫ�����ڹ�ҩ��λ����������ҩƷ��Ӧ����Ϣ��", vbInformation, gstrSysName
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
            Case "Provider"
                txt��Ӧ��.SetFocus
                txt��Ӧ��.SelStart = 0
                txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
            
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
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Provider"
                    txt��Ӧ��.Text = .TextMatrix(.Row, 3)
                    txt��Ӧ��.Tag = .TextMatrix(.Row, 0)
                    Txt������.SetFocus
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
            Chk��Ӧ��.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            Chk��Ӧ��.SetFocus
        End If
    End If
    
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt��Ӧ��)) <> "" Then
        txt��Ӧ�� = UCase(txt��Ӧ��)
        With RecTmp
            gstrSQL = "Select id,����,����,���� From ҩƷ��Ӧ�� Where " & _
                  " To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' And ĩ��=1" & _
                  " And (���� like '" & IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%' or ���� like '" & IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%' or ���� like '" & IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ�� & "%')"
            Call OpenRecordset(RecTmp, "ҩƷ��Ӧ��")
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                txt��Ӧ��.Tag = 0
                txt��Ӧ��.SelStart = 0
                txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Provider"
                Set mshSelect.Recordset = RecTmp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + txt��Ӧ��.Top + txt��Ӧ��.Height
                    .Left = sstFilter.Left + fra��������.Left + txt��Ӧ��.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 800
                    .ColWidth(3) = .Width - .ColWidth(1) - .ColWidth(2)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt��Ӧ�� = !����
                txt��Ӧ��.Tag = !Id
                
            End If
        End With
    End If
    
    Txt������.SetFocus
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
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt����NO = UCase(LTrim(Me.txt����NO))
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
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
        With rsTemp
            gstrSQL = "Select ���,����,���� From ��Ա�� Where upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%' or Upper(���) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%' or Upper(����) like '" & Me.Txt����� & "%'"
            Call OpenRecordset(rsTemp, "ȡ�����")
            
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

Private Sub Txt������_GotFocus()
    Txt������.SelStart = 0
    Txt������.SelLength = Len(Txt������.Text)
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
        With rsTemp
            gstrSQL = "Select ���,����,���� From ��Ա�� Where upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%' or Upper(���) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%' or Upper(����) like '" & Me.Txt������ & "%'"
            Call OpenRecordset(rsTemp, "ȡ������")
            
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
