VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
      TabIndex        =   32
      Top             =   4245
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
      TabIndex        =   33
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
      Tab(0).Control(1)=   "chkDept(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDept(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkDept(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDept(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkDept(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmDrugPaymentSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&W)"
         Height          =   195
         Index           =   4
         Left            =   4830
         TabIndex        =   13
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "�豸(&S)"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   11
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&M)"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   10
         Tag             =   "2"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "ҩƷ(&D)"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Tag             =   "1"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "����(&Q)"
         Height          =   195
         Index           =   3
         Left            =   3675
         TabIndex        =   12
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra�������� 
         Height          =   3225
         Left            =   -74760
         TabIndex        =   41
         Top             =   510
         Width           =   5505
         Begin VB.ComboBox cboStore 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   750
            Width           =   3970
         End
         Begin VB.CheckBox chkStore 
            Caption         =   "�ⷿ"
            Height          =   300
            Left            =   330
            TabIndex        =   17
            Top             =   750
            Width           =   800
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4935
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   390
            Width           =   255
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1245
            TabIndex        =   20
            Tag             =   "Ʒ��"
            Top             =   1125
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1245
            TabIndex        =   22
            Tag             =   "��ʼ��Ʊ��"
            Top             =   1500
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1245
            TabIndex        =   24
            Tag             =   "������Ʊ��"
            Top             =   1875
            Width           =   3945
         End
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��"
            Height          =   300
            Left            =   330
            TabIndex        =   14
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   15
            Top             =   375
            Width           =   3945
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1245
            MaxLength       =   8
            TabIndex        =   26
            Top             =   2250
            Width           =   3945
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   1245
            MaxLength       =   8
            TabIndex        =   28
            Top             =   2625
            Width           =   3945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ʒ��"
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   19
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ʼ��Ʊ��"
            Height          =   180
            Index           =   7
            Left            =   300
            TabIndex        =   21
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������Ʊ��"
            Height          =   180
            Index           =   5
            Left            =   300
            TabIndex        =   23
            Top             =   1935
            Width           =   900
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   660
            TabIndex        =   25
            Top             =   2310
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   660
            TabIndex        =   27
            Top             =   2685
            Width           =   540
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   2850
         Left            =   240
         TabIndex        =   34
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
            Format          =   269221891
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
            Format          =   269221891
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
            Format          =   269221891
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
            Format          =   80281603
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6330
      TabIndex        =   31
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   30
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   29
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
Private mblnAdvance As Boolean '�Ƿ�չ��
Private mdatStart As Date   '��ʼʱ��
Private mdatEnd As Date     '����ʱ��
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mstrPrivs As String
Private mstr���� As String
Private mstrOthers(0 To 9) As String '0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��,9-�ⷿID

Public Function GetSearch(ByVal FrmMain As Form, ByVal strPrivs As String, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, ByRef str���� As String, ByRef strOthers() As String) As String
'--------------------------------------------------------------
'���ܣ���ȡ���������SQL���
'������FrmMain----------���ô���
'      datStart---------��ʼ����
'      datEnd-----------��������
'      datVerifyStart---�����ʼ����
'      datVerifyEnd-----��˽�������
'      strOthers-������������(0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��)
'���أ�SQL���
'˵����
'--------------------------------------------------------------
    mstrFind = "": mstrPrivs = strPrivs
    
    If Not CheckCompete Then Exit Function
    Call Ȩ������
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    str���� = mstr����
    strOthers = mstrOthers
End Function

Private Sub chkDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub chkStore_Click()
    cboStore.Enabled = chkStore.Value = 1
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


Private Sub cmdHelp_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim strTemp As String
    
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt��Ӧ��.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt��Ӧ��.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    Unload frm��Ӧ��ѡ��
End Sub

Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    Dim strFind As String, strKey As String
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
    'by lesfeng 2009-12-2 �����Ż�
    'mstrOthers(0 To 8) '0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��
    '������Χ: 1-��ʼ��������,2-������������
    '          3-��ʼ�������,4-�����������
    '          5-��ʼ����,6-��������,7-��Ӧ��ID,8-�����,9-������,10-��ʼ��Ʊ��,11-������Ʊ��,12-Ʒ��,13-�ⷿID
    
    If chk����.Value = 1 And chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
           mstrFind = " And ((A.�������� Between [1] And [2]) or (A.������� Between [3] And [4]))"
        Else
           mstrFind = " And ((A.�������� Between [1] And [2]) or (A.������� Between [3] And [4])) And A.��¼״̬ =1"
        End If
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.������� Between [3] And [4]"
        Else
            mstrFind = " And A.������� Between [3] And [4] And A.��¼״̬ =1"
        End If
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [1] And [2]) And A.������� is null "
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
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
    
    mstrOthers(1) = Me.txt��ʼNo
    mstrOthers(2) = Me.txt����NO
    
    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [5] And A.No <= [6] "
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [5] "
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <= [6] "
    
    '��չ��ѯ����
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
    
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    mstrOthers(3) = Val(txt��Ӧ��.Tag)
    If Chk��Ӧ��.Value = 1 Then
        mstrFind = mstrFind & " and a.��λid= [7] "
'        mstrOthers(3) = Val(txt��Ӧ��.Tag)
    End If
    
    If Me.Txt����� <> "" Then
        mstrFind = mstrFind & " And A.����� like [8] "
        mstrOthers(4) = IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%"
    End If
    If Me.Txt������ <> "" Then
        mstrFind = mstrFind & " And A.������ like [9] "
        mstrOthers(5) = IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%"
    End If
    
    strTemp = ""
    
    If Me.txtEdit(1).Text <> "" And Me.txtEdit(2) <> "" Then
        strTemp = strTemp & " And ��Ʊ�� >= [10] And ��Ʊ�� <= [11] "
        mstrOthers(6) = Me.txtEdit(1)
        mstrOthers(7) = Me.txtEdit(2)
    End If
    If Me.txtEdit(1) <> "" And Me.txtEdit(2) = "" Then
        strTemp = strTemp & " And ��Ʊ�� >= [10] "
        mstrOthers(6) = Me.txtEdit(1)
    End If
    If Me.txtEdit(1) = "" And Me.txtEdit(2) <> "" Then
        strTemp = strTemp & " And ��Ʊ�� <= [11] "
        mstrOthers(7) = Me.txtEdit(2)
    End If
    If Trim(Me.txtEdit(0)) <> "" Then
        strKey = GetMatchingSting(Trim(Me.txtEdit(0)), True)
        strFind = " And Ʒ�� like [12]"
        mstrOthers(8) = GetMatchingSting(Me.txtEdit(0).Text, False)
        If zlcommfun.IsCharAlpha(Trim(txtEdit(0).Text)) Then          '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            If gSystemPara.int���뷽ʽ = 1 Then
                '������ѯ
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And zltools.zlWBCode(Ʒ��) Like upper([12]) "
            ElseIf gSystemPara.int���뷽ʽ = 0 Then
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And zltools.zlspellcode(Ʒ��) Like upper([12]) "
            Else
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And (zltools.zlWBCode(Ʒ��) Like upper([12]) or zltools.zlspellcode(Ʒ��) Like upper([12])"
            End If
            
        End If
        strTemp = strTemp & strFind
    End If
    
    If chkStore.Value = 1 Then
        strTemp = strTemp & " And �ⷿID = [13] "
        If cboStore.ListIndex = -1 Then
            mstrOthers(9) = ""
        Else
            mstrOthers(9) = cboStore.ItemData(cboStore.ListIndex)
        End If
    End If
    
    If strTemp <> "" Then
        mstrFind = mstrFind & " And  Exists (Select 1 From Ӧ����¼ Where a.�������=������� " & strTemp & ") "
    End If
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
    Me.dtp����ʱ��(0) = zlDatabase.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    
    Me.dtp��ʼʱ��(0) = DateAdd("d", -15, Me.dtp����ʱ��(0))
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
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����,����,ĩ��,���� " & _
        "   From ��Ӧ�� " & _
        "   Where ���� is Not NULL " & zl_��ȡվ������ & " " & _
        "       And (����ʱ�� is null or ����ʱ��>= to_date('3000-01-01','yyyy-mm-dd')) " & _
        "   Start with �ϼ�ID is NULL Connect by prior id=�ϼ�id"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsCompete, gstrSQL, "����"
    
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "��Ӧ����Ϣ��ȫ�����ڹ�Ӧ�̹��������ù�Ӧ����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
    
errHandle:
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

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit(0), KeyAscii, m�ı�ʽ)
End Sub
Private Function Select��Ӧ��(ByVal strKey As String) As Boolean
    '----------------------------------------------------------------------------------------
    '����:ѡ��Ӧ��
    '����:strKey-ѡ��Ӧ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/11/5
    '----------------------------------------------------------------------------------------
    Dim strȨ�� As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
    End If
      
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs)
    gstrSQL = "" & _
        "   Select id, ����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
        "   From ��Ӧ�� " & _
        "   where ĩ��=1 " & zl_��ȡվ������ & " " & _
        "          and  (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
        "          and (���� like [1] or ���� like [1] or ���� like [1])  " & strȨ��
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
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
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim lngH As Long
    vRect = zlControl.GetControlRect(txt��Ӧ��.hwnd)
 
    lngH = txt��Ӧ��.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��ѡ��", False, "", "ѡ��Ӧ��", False, True, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, True, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "�����ڷ��������Ĺ�Ӧ��,����!"
        Exit Function
    End If
    If rsTemp.State <> 1 Then Exit Function
    txt��Ӧ�� = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
    txt��Ӧ��.Tag = Nvl(rsTemp!ID)
    Select��Ӧ�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt��Ӧ��_Change()
    txt��Ӧ��.Tag = ""
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If txt��Ӧ��.Tag <> "" Then
        zlcommfun.PressKey vbKeyTab
        Exit Sub
    End If
    If Select��Ӧ��(txt��Ӧ��.Text) = False Then
        If txt��Ӧ��.Enabled And txt��Ӧ��.Visible Then txt��Ӧ��.SetFocus
        Exit Sub
    End If
    zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
'����29757 by lesfeng 2010-05-10
Private Sub txt��Ӧ��_Validate(Cancel As Boolean)
    Dim rsTemp As New Recordset
    
    If txt��Ӧ��.Tag <> "" Then Exit Sub
    
    If Select��Ӧ��(txt��Ӧ��.Text) = False Then
        Exit Sub
    End If

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

Private Sub Txt�����_Change()
    Txt�����.Tag = ""
End Sub

Private Sub Txt�����_GotFocus()
    Txt�����.SelStart = 0
    Txt�����.SelLength = Len(Txt�����.Text)
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Txt�����.Text <> "" And Txt�����.Tag = "" Then
        Dim lng��ԱID As Long
        
        If MulitSelectPersion(Me, Txt�����, Trim(Txt�����.Text), 0, lng��ԱID) = False Then
            If Txt�����.Enabled Then Txt�����.SetFocus
            Exit Sub
        End If
        Txt�����.Tag = lng��ԱID
    End If
    zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������_Change()
    Txt������.Tag = ""
End Sub

Private Sub Txt������_GotFocus()
    Txt������.SelStart = 0
    Txt������.SelLength = Len(Txt������.Text)
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Txt������.Text <> "" And Txt������.Tag = "" Then
        Dim lng��ԱID As Long
        
        If MulitSelectPersion(Me, Txt������, Trim(Txt������.Text), 0, lng��ԱID) = False Then
            If Txt������.Enabled Then Txt������.SetFocus
            Exit Sub
        End If
        Txt������.Tag = lng��ԱID
    End If
    zlcommfun.PressKey vbKeyTab

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Public Sub Ȩ������()
    
    If Check���Ȩ��(mstrPrivs, "ҩƷ") = False Then
        chkDept(0).Enabled = False
        chkDept(0).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkDept(1).Enabled = False
        chkDept(1).Value = 0
    End If
    
    If Check���Ȩ��(mstrPrivs, "�豸") = False Then
        chkDept(2).Enabled = False
        chkDept(2).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkDept(3).Enabled = False
        chkDept(3).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkDept(4).Enabled = False
        chkDept(4).Value = 0
    End If
End Sub

