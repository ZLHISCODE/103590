VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTransferSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4260
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7560
   Icon            =   "frmTransferSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2190
      Left            =   6075
      TabIndex        =   31
      Top             =   2250
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3863
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
      Left            =   105
      TabIndex        =   25
      Top             =   135
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
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
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   5520
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   1530
            TabIndex        =   33
            Top             =   2400
            Width           =   3765
         End
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4900
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   960
            Width           =   270
         End
         Begin VB.TextBox txtDept 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            TabIndex        =   13
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox Chk����ⷿ 
            Caption         =   "����ⷿ"
            Height          =   420
            Left            =   420
            TabIndex        =   12
            Top             =   945
            Width           =   1110
         End
         Begin VB.CommandButton Cmd���� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4900
            TabIndex        =   11
            Top             =   420
            Width           =   255
         End
         Begin VB.TextBox Txt���� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            ScrollBars      =   3  'Both
            TabIndex        =   10
            Top             =   420
            Width           =   3375
         End
         Begin VB.CheckBox Chk���� 
            Caption         =   "��������"
            Height          =   300
            Left            =   420
            TabIndex        =   9
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   16
            Top             =   1500
            Width           =   1845
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   17
            Top             =   1980
            Width           =   1845
         End
         Begin VB.ComboBox Cbo����ⷿ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   980
            Width           =   3615
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��  ��"
            Height          =   180
            Left            =   570
            TabIndex        =   34
            Top             =   2460
            Width           =   540
         End
         Begin VB.Label LblEnterStock 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ϲ���(&D)"
            Height          =   180
            Left            =   480
            TabIndex        =   32
            Top             =   1005
            Width           =   990
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   570
            TabIndex        =   23
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   570
            TabIndex        =   24
            Top             =   2040
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "δ��˳���"
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "����˳���"
            Enabled         =   0   'False
            Height          =   300
            Left            =   3585
            TabIndex        =   35
            Top             =   2280
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
            Format          =   169738243
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
            Format          =   169738243
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
            Format          =   169738243
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
            Format          =   169738243
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   20
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
      TabIndex        =   19
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   18
      Top             =   420
      Width           =   1100
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
Private mstrOthers(0 To 13) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
Private mint�������� As Integer '0-����Ҫ���� 1-��Ҫ����
Private mstrPrivs As String

Public Function GetSearch(ByVal frmMain As Form, ByVal lngMode As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strPrivs As String, _
        ByRef strOthers() As String) As String
        'strOthers():������ز���ֵ:0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����)
        
    mstrFind = ""
    mlngMode = lngMode
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    Me.Show vbModal, mfrmMain
    
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
    
End Function

Private Sub Cbo����ⷿ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
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

Private Sub Chk����_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk����.SetFocus
    End If
End Sub

Private Sub Chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk����_Click()
    Txt����.Enabled = IIf(Chk����.Value = 1, True, False)
    Cmd����.Enabled = IIf(Chk����.Value = 1, True, False)
End Sub

Private Sub Chk����ⷿ_click()
    If mlngMode = 1718 Then
        Cbo����ⷿ.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
    Else
        txtDept.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
        cmdDept.Enabled = IIf(Chk����ⷿ.Value = 1, True, False)
    End If
End Sub

Private Sub Chk����ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey vbKeyTab
'    If Chk����ⷿ.Value = 1 Then
'        Cbo����ⷿ.SetFocus
'    Else
'        Txt������.SetFocus
'    End If
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

Private Sub cmdDept_Click()
    If getDept("") = False Then
        Exit Sub
    End If
    If Txt������.Enabled Then Txt������.SetFocus
End Sub

Private Sub Cmdȡ��_Click()
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    mstrFind = ""
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    '�������
    If Chk����.Value = 1 Then
        If Txt����.Tag = 0 Then
            MsgBox "��ѡ�����ѯ�Ĳ�����Ϣ��", vbInformation, gstrSysName
            Me.Txt����.SetFocus
            Exit Sub
        End If
    End If
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "����ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
    End If
    mstrFind = ""
    '������ѯ����
    
    mdatStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
    
    '2-��ʼ��������,3-������������
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        If mlngMode <> 1716 Then '�����ƿ�
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.�������� Between [2] And [3] and A.����� is null) " _
                        & " or (A.������� Between [4] And [5]))"
            Else
                mstrFind = " And ((A.�������� Between [2] And [3] and A.����� is null) " _
                        & " or (A.������� Between [4] And [5] and a.��¼״̬ =[6]))  "
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.�������� Between [2] And [3] and A.����� is null) " _
                    & " or (A.������� Between [4] And [5]))"
            Else
                If chkNoStrike.Value = 1 And chkYesStrike.Value = 1 Then
                    mstrFind = " And ((A.�������� Between [2] And [3] and A.����� is null) " _
                                & " or (A.������� Between [4] And [5]))"
                ElseIf chkNoStrike.Value = 1 And chkYesStrike.Value = 0 Then
                    mstrFind = " and (A.��¼״̬=2 or mod(A.��¼״̬,3)=2) And A.�������� Between [2] And [3] and A.����� is null  "
                ElseIf chkNoStrike.Value = 0 And chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.��¼״̬=2 or mod(A.��¼״̬,3)=2) And ((A.�������� Between [2] And [3] and A.����� is null) " _
                                & " or (A.������� Between [4] And [5])) and A.����� is not null "
                Else
                    mstrFind = " And ((A.�������� Between [2] And [3] and A.����� is null) " _
                                & " or (A.������� Between [4] And [5])) and a.��¼״̬ =1 "
                End If
            End If
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
    ElseIf chk���.Value = 1 Then
        If mlngMode <> 1716 Then
            If chkStrike.Value = 1 Then
                mstrFind = " And A.������� Between [4] And [5] "
            Else
                mstrFind = " And A.������� Between [4] And [5] and a.��¼״̬ =[6] "
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And A.������� Between [4] And [5] "
            Else
                If chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.��¼״̬=2 or mod(A.��¼״̬,3)=2) And A.������� Between [4] And [5] "
                Else
                    mstrFind = " And A.������� Between [4] And [5] and a.��¼״̬ =1"
                End If
            End If
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        If mlngMode <> 1716 Then
            mstrFind = " And (A.�������� Between [2] And [3] and ������� is null ) "
        Else
            If chkNoStrike.Value = 1 Then
                mstrFind = " and (A.��¼״̬=2 or mod(A.��¼״̬,3)=2) and (A.�������� Between [2] And [3]) and ������� is null "
            Else
                mstrFind = " And (A.�������� Between [2] And [3]) and ������� is null "
            End If
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        Me.txt����NO = UCase(LTrim(Me.txt����NO))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
    End If
    
    
    mstrOthers(1) = Trim(Me.txt��ʼNo.Text)
    mstrOthers(2) = Trim(Me.txt����NO.Text)
       
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [7] And A.No <=[8] "
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >=[7] "
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <=[8] "
    
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    
    If Chk����.Value = 1 Then
        mstrFind = mstrFind & " And A.ҩƷID=[9]"
        mstrOthers(3) = Val(Txt����.Tag)
    End If
    
    If Chk����ⷿ.Value = 1 Then
        If txtDept.Visible = True Then
            mstrOthers(4) = txtDept.Tag
        Else
            mstrOthers(4) = Cbo����ⷿ.ItemData(Cbo����ⷿ.ListIndex)
        End If
    End If
    If mlngMode = 1718 Then
        If Chk����ⷿ.Value = 1 Then
            mstrFind = mstrFind & " And A.������ID=[10]"
        End If
    Else
        If Chk����ⷿ.Value = 1 Then
            mstrFind = mstrFind & " And A.�Է�����ID=[10]"
        End If
        
    End If
    If Me.Txt������ <> "" Then
        mstrFind = mstrFind & " And A.������ like [11] "
        mstrOthers(5) = Txt������.Text & "%"
    End If
    
    If Me.Txt����� <> "" Then
        mstrFind = mstrFind & " And A.����� like [12]"
        mstrOthers(6) = Me.Txt����� & "%"
    End If
    
    If gblnCode = True And Trim(txt����.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt����.Text))
        mstrFind = mstrFind & " And (A.��Ʒ���� Like [19] Or A.�ڲ����� Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd����_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, 0, _
        mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), _
        True, True, False, False, True, 0, True, False, "", False, 0, False, mstrPrivs)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
    
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
    
    Dim intLop As Integer
    
    Me.dtp����ʱ��(0) = sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    Me.Txt����.Tag = 0
    sstFilter.Tab = 0
    
    lbl����.Visible = gblnCode
    txt����.Visible = gblnCode
    
    Select Case mlngMode
        Case 1715   '��۵���
            Chk����ⷿ.Caption = "�ⷿ"
            lbl����.Visible = False
            txt����.Visible = False
        Case 1716 '�ƿ����
            mint�������� = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngMode, "0")) = 1, 1, 0)
            Chk����ⷿ.Caption = "�����ⷿ"
            If mint�������� = 1 Then
                chkStrike.Visible = False
                chkNoStrike.Visible = True
                chkYesStrike.Visible = True
            Else
                chkStrike.Visible = True
                chkNoStrike.Visible = False
                chkYesStrike.Visible = False
            End If
        Case 1717   '��������
            Chk����ⷿ.Caption = "���ò���"
        Case 1718   '�����������
            Chk����ⷿ.Caption = "������"
        Case 1719   '�̵����
            lbl����.Visible = False
            txt����.Visible = False
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
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsTemp As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            txtDept.Visible = False
            cmdDept.Visible = False
            
            Cbo����ⷿ.Visible = True
            If Cbo����ⷿ.ListCount < 1 Then
                Select Case mlngMode
                    Case 1716
                        txtDept.Visible = True
                        cmdDept.Visible = True
                        Cbo����ⷿ.Visible = False
                        Exit Sub
'                        strStock = "V,W,K,12"
'                        gstrSQL = "" & _
'                        "   SELECT /*+ Rule*/ DISTINCT a.id, a.���� " & _
'                        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a, Table(cast(f_Str2List([2]) as zlTools.t_StrList)) D " & _
'                        "   Where c.�������� = b.���� " & _
'                        "       AND b.����=D.Column_Value " & _
'                        "       AND a.id = c.����id " & _
'                        "       AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1717
                        txtDept.Visible = True
                        cmdDept.Visible = True
                        Cbo����ⷿ.Visible = False
                        Exit Sub
                        'strStock = "O"
'                        If Check��ͨ���� = True Then
'                            gstrSQL = "" & _
'                            "   SELECT DISTINCT a.id, a.���� " & _
'                            "   FROM  ���ű� a " & _
'                            "   Where a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd') " & _
'                            "       and a.ID in (Select ����ID From ������Ա Where ��ԱID=[1])"
'                        Else
'                            gstrSQL = "" & _
'                            "   SELECT DISTINCT a.id, a.���� " & _
'                            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
'                            "   Where c.�������� = b.���� " & _
'                            "       AND a.id = c.����id " & _
'                            "       AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
'                            '"               AND b.���� in " & strStock
'                        End If
                    Case 1718
                        gstrSQL = "" & _
                            "   SELECT b.Id,b.���� " & _
                            "   FROM ҩƷ�������� A, ҩƷ������ B " & _
                            "   Where A.���id = B.ID AND A.���� = 36 "
                    Case 1715, 1719
                        If Chk����ⷿ.Visible = True Then
                            Chk����ⷿ.Visible = False
                            Cbo����ⷿ.Visible = False
                        End If
                        LblEnterStock.Visible = False
                        Exit Sub
                    Case Else
                End Select
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, strStock)
                With Cbo����ⷿ
                    Do While Not rsTemp.EOF
                        .AddItem rsTemp.Fields(1)
                        .ItemData(.NewIndex) = rsTemp.Fields(0)
                        rsTemp.MoveNext
                    Loop
                    If .ListCount > 0 Then .ListIndex = 0
                End With
                rsTemp.Close
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDept_Change()
    txtDept.Tag = ""
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDept.Tag = "" Then
            If getDept(Trim(txtDept.Text)) = False Then
                Exit Sub
            End If
        End If
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim IntBill As Integer
    Dim lng�ⷿID As Long
    Dim strNo As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    Select Case mlngMode
        Case 1715         '���Ŀ���۵���'
            IntBill = 71        '���Ŀ���۵���
        Case 1716, 1722         '�����������'
            If UCase(mfrmMain.Name) = UCase("frmRequestStuffList") Then
                '��ʾ���쵥
                '��Ϊ����ȷ���ⷿ,���ֳܷ�����
                gstrSQL = "Select ��Ŀ��� From ������Ʊ� where ��Ŀ���=72 and ��Ź���<>2"
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If rsTemp.RecordCount <> 0 Then
                    GoTo NO:
                End If
                rsTemp.Close
                OS.PressKey (vbKeyTab)
                Exit Sub
            End If
            IntBill = 72        '���Ŀⷿת��
        Case 1717         '�������ù���'
            IntBill = 73        '������������
        Case 1718         '���������������'
            IntBill = 74        '������������
        Case 1719         '�����̵����'
            If mfrmMain.TabShow.Tab = 0 Then
                IntBill = 76        '��������̵�
            Else
                IntBill = 75        '��������̵�
            End If
        Case Else
            IntBill = 0
    End Select
    
    If IntBill = 0 Then
NO:
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt����NO = UCase(LTrim(Me.txt����NO))
            intYear = Format(sys.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
        End If
        OS.PressKey (vbKeyTab)
    Else
        lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If KeyCode = vbKeyReturn Then
            If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
                txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, IntBill, lng�ⷿID)
            End If
            OS.PressKey (vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
     
    Dim IntBill As Integer
    Dim lng�ⷿID As Long
    Dim strNo As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo ErrHandle
    
    Select Case mlngMode
        Case 1715         '���Ŀ���۵���'
            IntBill = 71        '���Ŀ���۵���
        Case 1716               '�ƿ�'
            If UCase(mfrmMain.Name) = UCase("frmRequestStuffList") Then
                '��ʾ���쵥
                '��Ϊ����ȷ���ⷿ,���ֳܷ�����
                gstrSQL = "Select ��Ŀ��� From ������Ʊ� where ��Ŀ���=72 and ��Ź���<>2"
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If rsTemp.RecordCount <> 0 Then
                    GoTo NO:
                End If
                rsTemp.Close
                OS.PressKey (vbKeyTab)
                Exit Sub
            End If
            IntBill = 72        '���Ŀⷿת��
        Case 1722               '�����������
        Case 1717         '�������ù���'
            IntBill = 73        '������������
        Case 1718         '���������������'
            IntBill = 74        '������������
        Case 1719         '�����̵����'
            If mfrmMain.TabShow.Tab = 0 Then
                IntBill = 76        '��������̵�
            Else
                IntBill = 75        '��������̵�
            End If
        Case Else
            IntBill = 0
    End Select
    
    If IntBill = 0 Then
NO:
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            Dim intYear  As Integer, strYear As String
            
            Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
            intYear = Format(sys.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
        End If
        OS.PressKey (vbKeyTab)
    Else
        lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If KeyCode = vbKeyReturn Then
            If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
                txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, IntBill, lng�ⷿID)
            End If
            OS.PressKey (vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            cmdȷ��.SetFocus
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)
        
        gstrSQL = "" & _
            "   Select ���,����,���� " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) And (վ��=[2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetMatchingSting(Me.Txt�����), gstrNodeNo)
        
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
                    If .Top + .Height > Me.ScaleHeight Then .Top = .Top - .Height - Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    .ZOrder 0
                    Exit Sub
                    
                End With
            Else
                Txt����� = IIf(IsNull(!����), "", !����)
                cmdȷ��.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            Txt�����.SetFocus
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)
        
        gstrSQL = "" & _
            "   Select ���,����,���� " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) And (վ��=[2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetMatchingSting(Me.Txt������), gstrNodeNo)
        
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
                    If .Top + .Height > Me.ScaleHeight Then .Top = .Top - .Height - Txt������.Height
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt����.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra��������.Left + Txt����.Left
    sngTop = Me.Top + fra��������.Top + Txt����.Top + Txt����.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt����.Height - 3630
    End If
    
    strKey = Trim(Txt����.Text)
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), _
        strKey, sngLeft, sngTop, Txt����.Width, Txt����.Height, _
        True, True, False, False, True, _
        0, True, "", False, 0, _
        False, mstrPrivs)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
    
    If Chk����ⷿ.Visible = True Then
        Chk����ⷿ.SetFocus
    Else
        Txt������.SetFocus
    End If
    
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
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

Private Function getDept(ByVal strKey As String) As Boolean
    Dim rsTemp As New Recordset
    Dim strSeach As String
    Dim vRect As RECT
    Dim lngH As Long
    Dim strStock As String
    Dim blnCancel As Boolean
    Dim strWhere As String
    
    strSeach = strKey
    
    strWhere = ""
    If strSeach <> "" Then
        strSeach = GetMatchingSting(strSeach)
        strWhere = "           and (a.���� like [1] or a.���� like [1] or a.���� like [1]) "
    End If
    
    Select Case mlngMode
    Case 1716
        strStock = "V,W,K,12"
        gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id,a.����, a.����,a.����,a.λ��" & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a, Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       AND b.����=D.Column_value " & _
        "       AND a.id = c.����id " & _
        "       AND (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)" & _
            strWhere
    Case 1717
        'strStock = "O"
        If Check��ͨ���� = True Then
            gstrSQL = "" & _
                "  SELECT DISTINCT a.id,a.����,a.����,a.����,a.λ�� " & _
            "      FROM ���ű� a " & _
            "      Where ( TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null ) And (a.վ��=[2] or a.վ�� is null) " & _
            "           and a.ID in (Select ����ID From ������Ա Where ��ԱID=[4] ) " & _
                strWhere
        Else
            gstrSQL = "" & _
            "   SELECT DISTINCT a.id,a.����, a.����,a.����,a.λ�� " & _
            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
            "       AND a.id = c.����id " & _
            "       AND (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)" & _
                strWhere
        End If
    End Select
    
    vRect = zlControl.GetControlRect(txtDept.hwnd)
    lngH = txtDept.Height
    
    'If strkey = "" Then Exit Function
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����ѡ��", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSeach, gstrNodeNo, strStock, UserInfo.Id)

   '     frmParent=��ʾ�ĸ�����
   '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
   '     bytStyle=ѡ�������
   '       Ϊ0ʱ:�б���:ID,��
   '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
   '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
   '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
   '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
   '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
   '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
   '             bytStyle=1ʱ,�����Ǳ��������
   '     strNote=ѡ������˵������
   '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
   '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
   '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
   '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
   '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
   '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    If rsTemp Is Nothing Then
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    
    Me.txtDept = zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����)
    Me.txtDept.Tag = zlStr.Nvl(rsTemp!Id)
    getDept = True
    
End Function
