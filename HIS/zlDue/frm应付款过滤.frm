VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmӦ������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ�����������"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   29
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   28
      Top             =   405
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2535
      Left            =   750
      TabIndex        =   30
      Top             =   4485
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
      Height          =   3960
      Left            =   105
      TabIndex        =   31
      Top             =   135
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   6985
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmӦ�������.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkDept(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDept(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkDept(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDept(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkDept(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "��������(&F)"
      TabPicture(1)   =   "frmӦ�������.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&W)"
         Height          =   195
         Index           =   4
         Left            =   4785
         TabIndex        =   34
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&Q)"
         Height          =   195
         Index           =   3
         Left            =   3720
         TabIndex        =   18
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "ҩƷ(&D)"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   15
         Tag             =   "1"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&M)"
         Height          =   195
         Index           =   1
         Left            =   1545
         TabIndex        =   16
         Tag             =   "2"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "�豸(&S)"
         Height          =   195
         Index           =   2
         Left            =   2685
         TabIndex        =   17
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra��Χ 
         Height          =   2685
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������(&K)"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   14
            Top             =   2280
            Width           =   1440
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����˵���(&V)"
            Height          =   270
            Left            =   480
            TabIndex        =   9
            Top             =   1560
            Width           =   2070
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��˵���(&W)"
            Height          =   240
            Left            =   480
            TabIndex        =   4
            Top             =   840
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.TextBox txt����NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   3
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt��ʼNo 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   8
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   11
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   13
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   7
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������(&N)"
            Height          =   180
            Index           =   0
            Left            =   630
            TabIndex        =   5
            Top             =   1140
            Width           =   990
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   12
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������(&E)"
            Height          =   180
            Index           =   1
            Left            =   630
            TabIndex        =   10
            Top             =   1905
            Width           =   990
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   2
            Top             =   420
            Width           =   180
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N&o"
            Height          =   180
            Left            =   480
            TabIndex        =   0
            Top             =   420
            Width           =   180
         End
      End
      Begin VB.Frame fra�������� 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   32
         Top             =   585
         Width           =   5505
         Begin VB.ComboBox cboStore 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   940
            Width           =   3615
         End
         Begin VB.CheckBox chkStore 
            Caption         =   "�ⷿ(&W)"
            Height          =   300
            Left            =   435
            TabIndex        =   22
            Top             =   940
            Width           =   975
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   25
            Top             =   1320
            Width           =   3570
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   27
            Top             =   2070
            Width           =   1365
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   26
            Top             =   1695
            Width           =   1365
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   21
            Top             =   540
            Width           =   255
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   20
            Top             =   540
            Width           =   3375
         End
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��(&P)"
            Height          =   300
            Left            =   435
            TabIndex        =   19
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�������(&S)"
            Height          =   180
            Left            =   480
            TabIndex        =   24
            Top             =   1380
            Width           =   990
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������(&T)"
            Height          =   180
            Left            =   660
            TabIndex        =   36
            Top             =   1770
            Width           =   810
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����(&V)"
            Height          =   180
            Left            =   660
            TabIndex        =   35
            Top             =   2145
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmӦ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '�����ַ���
Private mblnAdvance As Boolean '�Ƿ�չ��
Private mdtStart As Date   '��ʼʱ��
Private mdtEnd As Date     '����ʱ��
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mstr���� As String
Private mstrPrivs As String
Private mcllFilter As Collection

Public Function GetSearch(ByVal FrmMain As Form, ByVal strPrivs As String, _
        ByRef dtStart As Date, ByRef dtEnd As Date, _
        ByRef dtVerifyStart As Date, ByRef dtVerifyEnd As Date, ByRef str���� As String, ByRef cllFilter As Collection) As String
    '--------------------------------------------------------------
    '���ܣ���ȡ���������SQL���
    '������FrmMain----------���ô���
    '      dtStart---------��ʼ����
    '      dtEnd-----------��������
    '      dtVerifyStart---�����ʼ����
    '      dtVerifyEnd-----��˽�������
    '���أ�SQL���
    '˵����
    '--------------------------------------------------------------
    mstrFind = "": mstrPrivs = strPrivs
    If Not CheckCompete Then Exit Function
    Call Ȩ������
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    dtStart = mdtStart
    dtEnd = mdtEnd
    dtVerifyStart = mdtVerifyStart
    dtVerifyEnd = mdtVerifyEnd
    str���� = mstr����
    Set cllFilter = mcllFilter
End Function

Public Sub Ȩ������()
    If Check���Ȩ��(gstrPrivs, "ҩƷ") = False Then
        chkDept(0).Enabled = False
        chkDept(0).Value = 0
    End If
    If Check���Ȩ��(gstrPrivs, "����") = False Then
        chkDept(1).Enabled = False
        chkDept(1).Value = 0
    End If
    
    If Check���Ȩ��(gstrPrivs, "�豸") = False Then
        chkDept(2).Enabled = False
        chkDept(2).Value = 0
    End If
    If Check���Ȩ��(gstrPrivs, "����") = False Then
        chkDept(3).Enabled = False
        chkDept(3).Value = 0
    End If
    If Check���Ȩ��(gstrPrivs, "����") = False Then
        chkDept(4).Enabled = False
        chkDept(4).Value = 0
    End If
End Sub

Private Sub chkDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkStore_Click()
    cboStore.Enabled = chkStore.Value = 1
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
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
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If

End Sub

Private Sub chk���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
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

Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim strTemp As String
    
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt��Ӧ��.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    txt��Ӧ��.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    Unload frm��Ӧ��ѡ��
End Sub

Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    '�������
    If Chk��Ӧ��.Value = 1 Then
        '����29757 by lesfeng 2010-05-10
        If Val(txt��Ӧ��.Tag) = 0 Then
            MsgBox "��ѡ�����ѯ�Ĺ�Ӧ����Ϣ��", vbInformation, gstrSysName
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
    '����SQL��ѯ�������
    Dim intTemp As Integer
    
    'by lesfeng 2010-1-7 �����Ż�
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��������"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�������"
    mcllFilter.Add Array("", ""), "���ݺ�"
    mcllFilter.Add "", "�������"
    mcllFilter.Add "", "��Ӧ��id"
    mcllFilter.Add "", "�ⷿID"
    mcllFilter.Add "", "������"
    mcllFilter.Add "", "�����"
            
    If chk����.Value = 1 And chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.�������� Between [1] And [2]) or (A.������� Between [3] And [4]))"
        Else
            mstrFind = " And ((A.�������� Between [1] And [2]) or (A.������� Between [3] And [4])) and a.��¼״̬ =1 "
        End If
        
        mcllFilter.Remove "��������"
        mcllFilter.Remove "�������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59"), "��������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59"), "�������"
        
        mdtStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdtEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdtVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdtVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.������� Between [3] And [4] "
        Else
            mstrFind = " And A.������� Between [3] And [4] and a.��¼״̬ =1 "
        End If
        mcllFilter.Remove "�������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59"), "�������"
        
        mdtVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdtVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdtStart = Format("1901-01-01", "yyyy-mm-dd")
        mdtEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [1] And [2]) and ������� is null "
        mcllFilter.Remove "��������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59"), "��������"
        
        mdtStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdtEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        
        mdtVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdtVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        Me.txt����NO = UCase(LTrim(Me.txt����NO))
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
    End If
    
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [5] And A.No <= [6]"
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [5]"
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= [6]"
    
    mcllFilter.Remove "���ݺ�"
    mcllFilter.Add Array(Trim(Me.txt��ʼNo), Trim(Me.txt����NO)), "���ݺ�"
 
    Dim strTemp As String
    
    Dim intIndex As Integer
    For intIndex = 0 To 4
        If chkDept(intIndex).Value = 1 Then
            strTemp = strTemp & "1"
        Else
            strTemp = strTemp & "0"
        End If
    Next
    mstr���� = strTemp ' Bin2Dec(strTemp)
    
    '��չ��ѯ����
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    If Trim(txtEdit.Text) <> "" Then mstrFind = mstrFind & " And A.������� like [7]"
    
    mcllFilter.Remove "�������"
    mcllFilter.Add GetMatchingSting(txtEdit), "�������"
    
    If Chk��Ӧ��.Value = 1 Then
        mstrFind = mstrFind & " and a.��λID = [8]"
        mcllFilter.Remove "��Ӧ��id"
        mcllFilter.Add txt��Ӧ��.Tag, "��Ӧ��id"
    End If
    
    If Me.Txt������ <> "" Then mstrFind = mstrFind & " And A.������ like [9]"
    If Me.Txt����� <> "" Then mstrFind = mstrFind & " And A.����� like [10]"
    
    If chkStore.Value = 1 Then
        mstrFind = mstrFind & " and a.�ⷿID = [11] "
        mcllFilter.Remove "�ⷿID"
        If cboStore.ListIndex = -1 Then
            mcllFilter.Add "", "�ⷿID"
        Else
            mcllFilter.Add cboStore.ItemData(cboStore.ListIndex), "�ⷿID"
        End If
    End If
    
    mcllFilter.Remove "������"
    mcllFilter.Add GetMatchingSting(Txt������), "������"
    mcllFilter.Remove "�����"
    mcllFilter.Add GetMatchingSting(Txt�����), "�����"
    Unload Me
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub Form_Load()
    Me.dtp����ʱ��(0) = zlDatabase.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    Me.txt��Ӧ��.Tag = 0
    '�򿪼�¼��
    sstFilter.Tab = 0
    mblnAdvance = False
End Sub

Private Function CheckCompete() As Boolean
    '--------------------------------------------------------------
    '���ܣ�����Ƿ��й�Ӧ������
    '������
    '���أ��Ƿ��й�Ӧ������
    '˵����
    '--------------------------------------------------------------
    Dim rsCompete As New Recordset
    
    CheckCompete = False
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select id From ��Ӧ�� Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') and  ĩ��=1 " & zl_��ȡվ������ & " and rownum<=2 "
    zlDatabase.OpenRecordset rsCompete, gstrSQL, Me.Caption
    With rsCompete
        If .EOF Then
            .Close
            ShowMsgbox "��Ӧ����Ϣ��ȫ�����ڹ�Ӧ�̹��������ù�Ӧ����Ϣ��"
            Exit Function
        End If
    End With
    CheckCompete = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
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
    Dim rsTmp As ADODB.Recordset

    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
        End If
        
        cboStore.Clear
        
        Set rsTmp = GetStoreInfo("'H', 'I', 'J', 'K', 'L', 'M', 'R', 'T', 'V', 'S'")
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> adStateOpen Then Exit Sub
        
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        Do While rsTmp.EOF = False
            cboStore.AddItem "[" & rsTmp!���� & "]" & rsTmp!����
            cboStore.ItemData(cboStore.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
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
 
Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
    zlcommfun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit, KeyAscii, m�ı�ʽ)
End Sub

Private Sub txt��Ӧ��_GotFocus()
    SetTxtGotFocus txt��Ӧ��, True
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    Dim strTemp As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    If LTrim(RTrim(txt��Ӧ��)) <> "" Then
        txt��Ӧ�� = UCase(txt��Ӧ��)
        strTemp = GetMatchingSting(txt��Ӧ��)
        Dim strȨ�� As String
        
        strȨ�� = " and " & Get����Ȩ��(gstrPrivs)
            
        gstrSQL = "" & _
            "   Select id,����,����,���� " & _
            "   From ��Ӧ�� " & _
            "   Where  (����ʱ�� is null or  To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
            "       " & zl_��ȡվ������ & " And ĩ��=1" & _
            "     And (���� like [1] or ���� like [1] or ���� like [1]) " & strȨ��
              
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
        If rsTemp.EOF Then
            MsgBox "�޴������Ĺ�Ӧ�̣�", vbInformation, gstrSysName
            KeyCode = 0
            txt��Ӧ��.Tag = 0
            txt��Ӧ��.SelStart = 0
            txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
            
            Exit Sub
        End If
        If rsTemp.RecordCount > 1 Then
            mstrSelectTag = "Provider"
            Set mshSelect.Recordset = rsTemp
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
            txt��Ӧ�� = rsTemp!����
            txt��Ӧ��.Tag = rsTemp!ID
        End If
    
    End If
    Txt������.SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt����NO_GotFocus()
    SetTxtGotFocus txt����NO, False
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt����NO = UCase(LTrim(Me.txt����NO))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt��ʼNo_GotFocus()
      SetTxtGotFocus txt��ʼNo, False
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
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
    SetTxtGotFocus Txt�����, True
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
        
        gstrSQL = "" & _
             "   Select ���,����,���� " & _
             "   From ��Ա�� " & _
             "   Where (���� like [1] or ��� like [1] or ���� like [1] )  " & zl_��ȡվ������ & " " & _
             "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
             "   order by ���"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%")
        
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
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt�����_LostFocus()
    ImeLanguage False
End Sub

Private Sub Txt������_GotFocus()
    SetTxtGotFocus Txt������, True
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
        gstrSQL = "" & _
             "   Select ���,����,���� " & _
             "   From ��Ա�� " & _
             "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & zl_��ȡվ������ & "" & _
             "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
             "   order by ���"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%")
          
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
                    If .Height > ScaleHeight - .Top Then
                        .Height = ScaleHeight - .Top - 20
                    Else
                        .Height = 2535
                    End If
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
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������_LostFocus()
    ImeLanguage False
End Sub
