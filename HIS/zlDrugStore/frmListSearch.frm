VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4260
   ClientLeft      =   3156
   ClientTop       =   3168
   ClientWidth     =   7692
   Icon            =   "frmListSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2535
      Left            =   960
      TabIndex        =   29
      Top             =   3090
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7853
      _ExtentY        =   4466
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
      TabIndex        =   23
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
      TabPicture(0)   =   "frmListSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmListSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   28
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox Chk����ⷿ 
            Caption         =   "�Ƴ��ⷿ"
            Height          =   420
            Left            =   360
            TabIndex        =   16
            Top             =   900
            Width           =   1110
         End
         Begin VB.CommandButton CmdҩƷ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   22
            Top             =   420
            Width           =   255
         End
         Begin VB.TextBox TxtҩƷ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   15
            Top             =   420
            Width           =   3375
         End
         Begin VB.CheckBox ChkҩƷ 
            Caption         =   "ҩƷ"
            Height          =   300
            Left            =   360
            TabIndex        =   14
            Top             =   420
            Width           =   990
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   19
            Top             =   1500
            Width           =   1845
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   21
            Top             =   1980
            Width           =   1845
         End
         Begin VB.ComboBox Cbo����ⷿ 
            Enabled         =   0   'False
            Height          =   276
            Left            =   1530
            TabIndex        =   17
            Text            =   "Cbo����ⷿ"
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   570
            TabIndex        =   18
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   570
            TabIndex        =   20
            Top             =   2040
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   5520
         Begin VB.TextBox txt��ʼNo 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt����NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   2
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   11
            Top             =   2280
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   5
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
            TabIndex        =   6
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
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   10
            Top             =   1845
            Width           =   1605
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
            TabIndex        =   0
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
            TabIndex        =   27
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
            TabIndex        =   8
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
            TabIndex        =   26
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
            TabIndex        =   4
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
            TabIndex        =   25
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6450
      TabIndex        =   13
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6450
      TabIndex        =   12
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmListSearch"
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
Private mlng�ⷿid As Long  '�ⷿid

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng�Ƴ��ⷿ As Long
    str������ As String
    str����� As String
End Type

Private SQLCondition As Type_SQLCondition
Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lng�ⷿid As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO��ʼ As String, _
        ByRef strNO���� As String, _
        ByRef date����ʱ�俪ʼ As Date, _
        ByRef date����ʱ����� As Date, _
        ByRef date���ʱ�俪ʼ As Date, _
        ByRef date���ʱ����� As Date, _
        ByRef lngҩƷ As Long, _
        ByRef lng�Ƴ��ⷿ As Long, _
        ByRef str������ As String, _
        ByRef str����� As String) As String
    mstrFind = ""
    mlngMode = lngMode
    mlng�ⷿid = lng�ⷿid
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
    lng�Ƴ��ⷿ = SQLCondition.lng�Ƴ��ⷿ
    str����� = SQLCondition.str�����
    str������ = SQLCondition.str������

End Function

Private Sub Cbo����ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "H,I,J,K,L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo����ⷿ.ListCount = 0 Then Exit Sub
    
    If Cbo����ⷿ.ListIndex >= 0 Then
        If Val(Cbo����ⷿ.Tag) = Cbo����ⷿ.ItemData(Cbo����ⷿ.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, Cbo����ⷿ, Trim(Cbo����ⷿ.Text), str��������, , "0,1,2,3") = False Then
        Exit Sub
    End If
    If Cbo����ⷿ.ListIndex >= 0 Then
        Cbo����ⷿ.Tag = Cbo����ⷿ.ItemData(Cbo����ⷿ.ListIndex)
    End If
End Sub

Private Sub Cbo����ⷿ_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cbo����ⷿ_Validate(Cancel As Boolean)
    If Cbo����ⷿ.ListCount > 0 Then
        If Cbo����ⷿ.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Chk����ⷿ_click()
    Cbo����ⷿ.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
End Sub

Private Sub Chk����ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk����ⷿ.Value = 1 Then
        Cbo����ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
End Sub
Private Sub chk����_Click()
    Dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    Dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    
End Sub

Private Sub chk���_Click()
    Dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    Dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
    
End Sub

Private Sub ChkҩƷ_Click()
    txtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    cmdҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
End Sub



Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    '�������
    If ChkҩƷ.Value = 1 Then
        If txtҩƷ.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
            Me.txtҩƷ.SetFocus
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
'            mstrFind = " And ((A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'                    & " or (A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                    & " or (A.������� Between [5] And [6]))"
        Else
'            mstrFind = " And ((A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'                    & " or (A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS'))) and a.��¼״̬ =1 "
            mstrFind = " And ((A.�������� Between [3] And [4] and ������� is null) " _
                    & " or (A.������� Between [5] And [6])) and (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) "
        End If
        
        mdatStart = Format(Dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(Dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(Dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(Dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
'            mstrFind = " And A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') "
            mstrFind = " and (A.��¼״̬=2 or A.��¼״̬=1 or mod(A.��¼״̬,3)=2 or mod(A.��¼״̬,3)=0) And A.������� Between [5] And [6] "
        Else
'            mstrFind = " And A.������� Between To_Date('" & Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') and a.��¼״̬ =1 "
            mstrFind = " and (a.��¼״̬ =1 or mod(A.��¼״̬,3)=0) And A.������� Between [5] And [6] "
        End If
        
        mdatVerifyStart = Format(Dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(Dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
'        mstrFind = " And (A.�������� Between To_Date('" & Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp����ʱ��(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and ������� is null "
        mstrFind = " And (A.�������� Between [3] And [4]) and ������� is null "
            
        mdatStart = Format(Dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(Dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNO) < 8 And Len(txt��ʼNO) > 0 Then
        txt��ʼNO.Text = GetFullNO(txt��ʼNO.Text, intNO, lng�ⷿid)
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        txt����NO.Text = GetFullNO(txt����NO.Text, intNO, lng�ⷿid)
    End If
    
'    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt��ʼNo & "' And A.No <='" & Me.txt����NO & "'"
    If Me.txt��ʼNO <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2] "
    
    
'    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt��ʼNo & "'"
    If Me.txt��ʼNO <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
       
'    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= '" & Me.txt����NO & "'"
    If Me.txt��ʼNO = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
        
    SQLCondition.strNO��ʼ = Me.txt��ʼNO
    SQLCondition.strNO���� = Me.txt����NO
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(Dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(Dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date���ʱ�俪ʼ = CDate(Format(Dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date���ʱ����� = CDate(Format(Dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59")
        
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    
    If ChkҩƷ.Value = 1 Then
'        mstrFind = mstrFind & " And A.ҩƷID=" & TxtҩƷ.Tag
        mstrFind = mstrFind & " And A.ҩƷID + 0=[7] "
    End If
    
    If mlngMode = 1343 Then
'        If Chk����ⷿ.Value = 1 Then mstrFind = mstrFind & " And A.�Է�����ID=" & Cbo����ⷿ.ItemData(Cbo����ⷿ.ListIndex)
        If Chk����ⷿ.Value = 1 Then
            mstrFind = mstrFind & " And A.�Է�����ID + 0=[8] "
        End If
    End If
'    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like '" & Me.Txt����� & "%'"
    If Me.Txt����� <> "" Then
        mstrFind = mstrFind & " And A.����� like [10] "
    End If
        
'    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like '" & Me.Txt������ & "%'"
    If Me.Txt������ <> "" Then
        mstrFind = mstrFind & " And A.������ like [9] "
    End If
    
    SQLCondition.lngҩƷ = Val(txtҩƷ.Tag)
    If Cbo����ⷿ.Visible Then
        SQLCondition.lng�Ƴ��ⷿ = Cbo����ⷿ.ItemData(Cbo����ⷿ.ListIndex)
    End If
    SQLCondition.str����� = Me.Txt����� & "%"
    SQLCondition.str������ = Me.Txt������ & "%"
    
    
    Unload Me
    
End Sub


Private Sub cmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ�������", mlng�ⷿid, mlng�ⷿid, mlng�ⷿid, , , True)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , mlng�ⷿid, mlng�ⷿid, mlng�ⷿid, , , , , 2, False)
    
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mlng�ⷿid, mlng�ⷿid)
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        txtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        txtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    txtҩƷ.Tag = RecReturn!ҩƷID
    
    If Chk����ⷿ.Visible = True Then
        Chk����ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
End Sub

Private Sub Dtp����ʱ��_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If BlnAdvance Then
            ChkҩƷ.SetFocus
        Else
            cmdȷ��.SetFocus
        End If
    End If
End Sub

Private Sub Dtp��ʼʱ��_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.Dtp����ʱ��(index).SetFocus
End Sub

Private Sub Form_Load()
    Dim intLop As Integer
    
    Me.Dtp����ʱ��(0) = zldatabase.Currentdate
    Me.Dtp����ʱ��(1) = Me.Dtp����ʱ��(0)
    Me.Dtp��ʼʱ��(0) = DateAdd("d", -7, Me.Dtp����ʱ��(0))
    Me.Dtp��ʼʱ��(1) = Me.Dtp��ʼʱ��(0)
    
    Me.txtҩƷ.Tag = 0
    sstFilter.Tab = 0
    
    Select Case mlngMode
        Case 1304
            Chk����ⷿ.Caption = "����ⷿ"
        Case 1305
            Chk����ⷿ.Caption = "���ò���"
        Case 1306
            Chk����ⷿ.Caption = "������"
    End Select
    
    '�򿪼�¼��
    
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

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsDepartment As New Recordset
    Dim strStock As String
    
    On Error GoTo errHandle
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            If Cbo����ⷿ.ListCount < 1 Then
                Select Case mlngMode
                    Case 1343
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                             & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                            & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
                              & "AND Instr([1],b.����,1) > 0 " _
                             & " AND a.id = c.����id " _
                              & "AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1344
                        If Chk����ⷿ.Visible = True Then
                            Chk����ⷿ.Visible = False
                            Cbo����ⷿ.Visible = False
                            
                        End If
                        Exit Sub
                        
                End Select

                Set rsDepartment = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
                
                With Cbo����ⷿ
                    Do While Not rsDepartment.EOF
                        .AddItem rsDepartment.Fields(1)
                        .ItemData(.NewIndex) = rsDepartment.Fields(0)
                        rsDepartment.MoveNext
                    Loop
                    .ListIndex = 0
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

Private Sub Txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = GetFullNO(txt����NO.Text, intNO, lng�ⷿid)
        End If
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt��ʼNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNO) < 8 And Len(txt��ʼNO) > 0 Then
            txt��ʼNO.Text = GetFullNO(txt��ʼNO.Text, intNO, lng�ⷿid)
        End If
        Me.txt����NO.SetFocus
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
    
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            cmdȷ��.SetFocus
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)
        
        gstrSQL = "Select ���,����,���� From ��Ա�� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", Me.Txt����� & "%")
            
        With rstemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rstemp
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


Private Sub Txt������_GotFocus()
    Txt������.SelStart = 0
    Txt������.SelLength = Len(Txt������.Text)
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt�����.SetFocus
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            Txt�����.SetFocus
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)
        
        gstrSQL = "Select ���,����,���� From ��Ա�� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ������]", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", Me.Txt������ & "%")
        
        With rstemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
                
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

Private Sub Txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra��������.Left + txtҩƷ.Left
    sngTop = Me.Top + fra��������.Top + txtҩƷ.Top + txtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - txtҩƷ.Height - 3630
    End If
    
    strKey = Trim(txtҩƷ.Text)
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mlng�ⷿid, mlng�ⷿid, strkey, sngLeft, sngTop)
    
    
    Call SetSelectorRS(1, "ҩƷ�������", mlng�ⷿid, mlng�ⷿid, mlng�ⷿid, , , True)

    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, strKey, sngLeft, sngTop, mlng�ⷿid, mlng�ⷿid, mlng�ⷿid, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        txtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        txtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    txtҩƷ.Tag = RecReturn!ҩƷID
    
    If Chk����ⷿ.Visible = True Then
        Chk����ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
    
End Sub

Private Sub TxtҩƷ_KeyPress(KeyAscii As Integer)
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

