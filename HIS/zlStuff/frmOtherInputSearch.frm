VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOtherInputSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmOtherInputSearch.frx":0000
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
      Height          =   2055
      Left            =   1320
      TabIndex        =   27
      Top             =   4080
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
      Height          =   3975
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmOtherInputSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmOtherInputSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   3300
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   5505
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   1650
            TabIndex        =   37
            Top             =   2760
            Width           =   3525
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "��������"
            Height          =   300
            Left            =   480
            TabIndex        =   17
            Top             =   1875
            Width           =   1095
         End
         Begin VB.CheckBox Chk��� 
            Caption         =   "���"
            Height          =   300
            Left            =   480
            TabIndex        =   15
            Top             =   1350
            Width           =   960
         End
         Begin VB.CommandButton Cmd���� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox Txt���� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
         Begin VB.CheckBox Chk���� 
            Caption         =   "��������"
            Height          =   315
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   1065
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   22
            Top             =   2355
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   24
            Top             =   2355
            Width           =   1365
         End
         Begin VB.ComboBox Cbo��� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   3495
         End
         Begin VB.CheckBox Chk������ 
            Caption         =   "������"
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   855
            Width           =   1155
         End
         Begin VB.TextBox txt������ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   13
            Top             =   855
            Width           =   3255
         End
         Begin VB.CommandButton Cmd������ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   14
            Top             =   855
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   2
            Left            =   1650
            TabIndex        =   18
            Top             =   1875
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   2
            Left            =   3540
            TabIndex        =   20
            Top             =   1875
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��  ��"
            Height          =   180
            Left            =   750
            TabIndex        =   38
            Top             =   2820
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   19
            Top             =   1935
            Width           =   180
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   750
            TabIndex        =   21
            Top             =   2415
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   3120
            TabIndex        =   23
            Top             =   2415
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   29
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
            Format          =   162791427
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
            Format          =   162791427
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
            Format          =   162791427
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
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
      TabIndex        =   26
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   25
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmOtherInputSearch"
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
Public lng����ID As Long
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mstrOthers(0 To 13) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ

Public Function GetSearch(ByVal frmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strOthers() As String) As String
        
    mstrFind = ""
    Set mfrmMain = frmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
End Function

Private Sub Cbo���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
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



Private Sub Chk���_Click()
    Cbo���.Enabled = IIf(Chk���.Value = 1, True, False)
End Sub

Private Sub Chk���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk���.Value = 1 Then
        Cbo���.SetFocus
    Else
        Txt������.SetFocus
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

Private Sub chk��������_Click()
    dtp��ʼʱ��(2).Enabled = chk��������.Value = 1
    dtp����ʱ��(2).Enabled = dtp��ʼʱ��(2).Enabled
End Sub

Private Sub chk��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk������_Click()
    Me.txt������.Enabled = IIf(Chk������.Value = 1, True, False)
    Cmd������.Enabled = IIf(Chk������.Value = 1, True, False)
End Sub

Private Sub Chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk������.Value = 1 Then
            txt������.SetFocus
        ElseIf Chk���.Visible = True Then
            Chk���.SetFocus
        Else
            Txt������.SetFocus
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

Private Sub Chk����_Click()
    Txt����.Enabled = IIf(Chk����.Value = 1, True, False)
    Cmd����.Enabled = IIf(Chk����.Value = 1, True, False)
End Sub

Private Sub Chk����_GotFocus()
    sstFilter.Tab = 1
    Chk����.SetFocus
End Sub

Private Sub Chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk����.Value = 1 Then
        Txt����.SetFocus
    Else
        Chk������.SetFocus
    End If
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
            MsgBox "��ѡ�����ѯ��������Ϣ��", vbInformation, gstrSysName
            Me.Txt����.SetFocus
            Exit Sub
        End If
    End If
    
    If Chk������.Value = 1 Then
        If txt������.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��������������Ϣ��", vbInformation, gstrSysName
            Me.txt������.SetFocus
            Exit Sub
        End If
    End If
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "�Բ��𣬱���ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
        

    mstrFind = ""
    '������ѯ����
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��
    mdatStart = Format("1901-01- 01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
      
    If chk����.Value = 1 And chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.�������� Between [2] And [3] and ������� is null) " _
                    & " or (A.������� Between [4] And [5]))"
        Else
            mstrFind = " And ((A.�������� Between [2] And [3] and ������� is null) " _
                    & " or (A.������� Between [4] And [5] and a.��¼״̬ =[6]))  "
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.������� Between [4] And [5] "
        Else
            mstrFind = " And A.������� Between [4] And [5] and a.��¼״̬ =[6] "
            
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [2] And To_Date('" & Format(dtp����ʱ��(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and ������� is null "
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
    End If
        
    
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
        intYear = Format(Sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        Me.txt����NO = UCase(LTrim(Me.txt����NO))
        intYear = Format(Sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
    End If
    
    mstrOthers(1) = Trim(Me.txt��ʼNo.Text)
    mstrOthers(2) = Trim(Me.txt����NO.Text)

    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [7] And A.No <=[8]"
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [7]"
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= [8]"
    
    '��չ��ѯ����
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id),5-������,
    ' 6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
     '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,
     '[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,
     '[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��
  
  
    If Chk����.Value = 1 Then
        lng����ID = Txt����.Tag
        mstrFind = mstrFind & " And A.ҩƷID=[9]"
        mstrOthers(3) = Txt����.Tag
    End If
    If Chk���.Value = 1 Then
        mstrFind = mstrFind & " And A.������ID=[10]"
        mstrOthers(4) = Cbo���.ItemData(Cbo���.ListIndex)
    End If
    
    If Me.Txt������ <> "" Then
        mstrFind = mstrFind & " And A.������ like '" & Me.Txt������ & "%'"
        mstrOthers(5) = Trim(Me.Txt������) & "%"
    End If
    If Me.Txt����� <> "" Then
        mstrFind = mstrFind & " And A.����� like [12]"
        mstrOthers(6) = Trim(Me.Txt�����) & "%"
    End If
    
    If Chk������.Value = 1 Then
        mstrFind = mstrFind & " And A.����=[14]"
        mstrOthers(8) = txt������.Text
    End If
    
    If chk��������.Value = 1 Then
        mstrFind = " And A.�������� Between [15] And [16] "
        mstrOthers(9) = Format(dtp��ʼʱ��(2), "yyyy-mm-dd")
        mstrOthers(10) = Format(dtp����ʱ��(2), "yyyy-mm-dd")
    End If
    
    If gblnCode = True And Trim(txt����.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt����.Text))
        mstrFind = mstrFind & " And (A.��Ʒ���� Like [19] Or A.�ڲ����� Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd������_Click()
    Dim rsTemp As New Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt������.hwnd)
    
    gstrSQL = "Select rownum as id,null as �ϼ�id,����,����,����,1 as ĩ�� From ���������� " & _
              "Where (վ�� = [1] or վ�� is null) "
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "����������ѡ��", True, "", "ѡ�����������̻���", False, False, True, vRect.Left - 15, vRect.Top, txt������.Height, blnCancel, False, False, gstrNodeNo)
    
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
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    With rsTemp
        txt������.Tag = 1
        txt������.Text = zlStr.Nvl(!����)
    End With
End Sub

Private Sub Cmd����_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
    
    Chk������.SetFocus
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
    
    Me.dtp��ʼʱ��(2) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp����ʱ��(2) = Me.dtp����ʱ��(0)
    
    lbl����.Visible = gblnCode
    txt����.Visible = gblnCode
    
    Me.Txt����.Tag = 0
    Me.txt������.Tag = 0
    lng����ID = 0
    
    '�򿪼�¼��
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    CheckCompete = False
    
    gstrSQL = "Select ����,����,���� From ���������� where (վ�� = [1] or վ�� is null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-����������", gstrNodeNo)
    With rsTemp
        If .EOF Then
            MsgBox "������������Ϣ��ȫ,�����ֵ����������������������Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    With rsTemp
        gstrSQL = "SELECT B.Id, b.���� " & _
                  "FROM ҩƷ�������� A, ҩƷ������ B " & _
                  "Where A.���id = B.ID AND A.���� = 32 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        If .EOF Then
            MsgBox "�����������û��������Ӧ����������������������࣡", vbInformation, gstrSysName
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Cbo���.AddItem .Fields(1)
            Cbo���.ItemData(Cbo���.NewIndex) = .Fields(0)
            .MoveNext
        Loop
        Cbo���.ListIndex = 0
        .Close
    End With
    CheckCompete = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Maker"
                txt������.SetFocus
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
            
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
                Case "Maker"
                    txt������.Text = .TextMatrix(.Row, 1)
                    txt������.Tag = 1
                    Chk���.SetFocus
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
            Chk����.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            Chk����.SetFocus
        End If
    End If
    
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long

    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, 70, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, 70, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If

End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmdȷ��.SetFocus
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
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) And (վ�� = [2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) " & _
            "   order by ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", gstrNodeNo)
            
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
                    .Top = sstFilter.Top + fra��������.Top + Txt�����.Top - .Height ' + Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left + Txt�����.Width - .Width
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt������ = "" Then Exit Sub
        If Trim(txt������) = "" Then Exit Sub
        txt������ = UCase(txt������)
    
        Dim rsTemp As New ADODB.Recordset
        
        gstrSQL = "Select ����,����,���� From ���������� Where upper(����) like [1] or Upper(����) like [1] or Upper(����) like [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", IIf(gstrMatchMethod = "0", "%", "") & Me.txt������ & "%")
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Maker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + txt������.Top + txt������.Height
                    .Left = sstFilter.Left + fra��������.Left + txt������.Left
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
                txt������ = IIf(IsNull(!����), "", !����)
                txt������.Tag = 1
            End If
        End With
        
        If Chk���.Visible = True Then
            If Chk���.Value = 1 Then
                Cbo���.SetFocus
            Else
                Chk���.SetFocus
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Me.Txt�����.SetFocus
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
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) And (վ�� = [2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", gstrNodeNo)
        
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
                    .Top = sstFilter.Top + fra��������.Top + Txt������.Top - .Height ' + Txt������.Height
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

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt����.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra��������.Left + Txt����.Left
    sngTop = Me.Top + sstFilter.Top + fra��������.Top + Txt����.Top + Txt����.Height + Me.Height - Me.ScaleHeight '  50
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
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strKey, sngLeft, sngTop, Txt����.Width, Txt����.Height)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
    
    Chk������.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

