VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFinanceSupervisePersonList 
   BorderStyle     =   0  'None
   Caption         =   "��Ա��Ϣ�б�"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   105
      ScaleHeight     =   570
      ScaleWidth      =   11625
      TabIndex        =   12
      Top             =   6915
      Width           =   11655
      Begin VB.Label lblBalance 
         Caption         =   "��ǰ�ݴ��:"
         Height          =   945
         Left            =   30
         TabIndex        =   13
         Top             =   75
         Width           =   7935
      End
   End
   Begin VB.PictureBox picPersonPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   285
      ScaleHeight     =   2520
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   3810
      Width           =   3435
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   -15
         TabIndex        =   10
         Top             =   -30
         Width           =   2865
         _Version        =   589884
         _ExtentX        =   5054
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picOtherList 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   195
      ScaleHeight     =   1140
      ScaleWidth      =   2745
      TabIndex        =   4
      Top             =   2055
      Width           =   2745
      Begin VB.TextBox txtOtherPerson 
         ForeColor       =   &H80000000&
         Height          =   315
         Left            =   675
         TabIndex        =   8
         Tag             =   "���������ֻ���"
         Text            =   "���������ֻ���"
         Top             =   0
         Width           =   2310
      End
      Begin MSComctlLib.ListView lvwOther_S 
         Height          =   825
         Left            =   165
         TabIndex        =   11
         Top             =   555
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "���"
            Text            =   "���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "��������"
            Text            =   "��������"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblOtherPerson 
         AutoSize        =   -1  'True
         Caption         =   "�շ�Ա"
         Height          =   210
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   630
      End
   End
   Begin VB.PictureBox picGroupList 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   135
      ScaleHeight     =   1140
      ScaleWidth      =   2745
      TabIndex        =   2
      Top             =   1410
      Width           =   2745
      Begin MSComctlLib.ListView lvwGroup_S 
         Height          =   825
         Left            =   -15
         TabIndex        =   3
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "������"
            Object.Tag             =   "������"
            Text            =   "������"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "����������"
            Object.Tag             =   "����������"
            Text            =   "����������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "˵��"
            Object.Tag             =   "˵��"
            Text            =   "˵��"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.PictureBox picPersonList 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   390
      ScaleHeight     =   1755
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   75
      Width           =   2700
      Begin VB.TextBox txtChargePerson 
         ForeColor       =   &H80000000&
         Height          =   315
         Left            =   675
         TabIndex        =   6
         Tag             =   "���������ֻ���"
         Text            =   "���������ֻ���"
         Top             =   45
         Width           =   2310
      End
      Begin MSComctlLib.ListView lvwPerson_S 
         Height          =   825
         Left            =   90
         TabIndex        =   1
         Top             =   495
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "���"
            Text            =   "���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "��������"
            Text            =   "��������"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblChargePerson 
         AutoSize        =   -1  'True
         Caption         =   "�շ�Ա"
         Height          =   210
         Left            =   0
         TabIndex        =   5
         Top             =   75
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ilssmall 
      Left            =   4170
      Top             =   1560
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
            Picture         =   "frmFinanceSupervisePersonList.frx":0000
            Key             =   "Man"
            Object.Tag             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":059A
            Key             =   "Woman"
            Object.Tag             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":0B34
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsbig 
      Left            =   4365
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":10CE
            Key             =   "Man"
            Object.Tag             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":19A8
            Key             =   "Woman"
            Object.Tag             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":2282
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFinanceSupervisePersonList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPgIndex
    EM_PG_�շ�Ա = 250101
    EM_PG_������ = 250102
    EM_PG_������Ա = 250103
End Enum
Private Enum mPaneIndex
    EM_PN_�շ�Ա�б� = 1
    EM_PN_��ϸ�б� = 2
    EM_PN_�ݴ��б� = 3
End Enum
Private WithEvents mfrmList As frmFinaceSuperviseCollectList
Attribute mfrmList.VB_VarHelpID = -1
Private mfrmPersonOther As frmFinanceSupervisePersonOthers

Private mlngModule As Long, mstrPrivs As String
Private mrsChargePerson As ADODB.Recordset '�շ�Ա��¼��
Private mrsOtherPerson As ADODB.Recordset   '������Ա��¼��
Private mrsGroup As ADODB.Recordset
Private mstrSelPerson As String '�ϴ�ѡ����շ���Ա
Private mstrSelGroup As String '�ϴ�ѡ�������Ա
Private mstrSelOther As String '�ϴ�ѡ���������Ա
Private mcbsMain As Object
Private mblnNotBrush As Boolean '��ˢ������
Private mobjDetailPane As Pane
Private Function LoadPersonFromLvw(ByVal objLvw As ListView, ByVal rsTemp As ADODB.Recordset) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա���ؼ�
    '���:objLvw-���ص�����
    '       rsTemp-�շ�Ա��(ID,���,����,����,��������)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-24 11:40:05
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, strIcon As String
    On Error GoTo errHandle
    'ȫ������
    With objLvw.ListItems
        .Clear
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Nvl(rsTemp!�Ա�) Like "*��*" Then
                strIcon = "Man"
            Else
                strIcon = "Woman"
            End If
            Set objItem = .Add(, "K" & Nvl(rsTemp!���), Nvl(rsTemp!����), strIcon, strIcon)
            objItem.SubItems(1) = Nvl(rsTemp!���)
            objItem.SubItems(2) = Nvl(rsTemp!����)
            objItem.SubItems(3) = Nvl(rsTemp!��������)
            objItem.Tag = Nvl(rsTemp!ID)
            rsTemp.MoveNext
        Loop
    End With
    LoadPersonFromLvw = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadPerson(Optional blnFilter As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա��Ϣ
    '���:blnFilter-�Ƿ���й���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�") = False Then LoadPerson = True: Exit Function
    
    '��ȡ�շ�Ա��Ϣ
    If blnFilter = False Or mrsChargePerson Is Nothing Then
        strSQL = "" & _
        "   Select distinct A.ID,A.���,A.����,A.����,M.���� as ��������,a.�Ա�" & _
        "   From ��Ա�� A,��Ա����˵�� B, ������Ա C,���ű� M" & _
        "   Where A.id = B.��ԱID And B.��Ա���� In ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���')  " & _
        "               And A.ID=C.��ԱID and C.����ID=M.ID(+) And C.ȱʡ(+)=1 " & _
        "               And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "               And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        "   Order By ���"
        Set mrsChargePerson = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ�Ա��Ϣ")
    End If
    mrsChargePerson.Filter = 0
    strText = UCase(txtChargePerson.Text)
    
    If txtChargePerson.Text = txtChargePerson.Tag Or strText = "" Then
        'ȫ������
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    End If
    
    strCompents = gstrLike & strText & "%"
    If IsNumeric(strText) Then '1.�������ȫ����
    ElseIf zlCommFun.IsCharAlpha(strText) Then '1-�������ȫ��ĸ
        mrsChargePerson.Filter = "���� like '" & gstrLike & strText & "%'"
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    Else
        intInputType = 2   '2-����
        mrsChargePerson.Filter = "���� like '" & strText & "%'"
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    End If
    
    '�������ȫ����
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsChargePerson)
    With mrsChargePerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '������������,��Ҫ���:
            '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
            '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
            '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
            If Nvl(!���) = strText Then
                Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp): Exit Do
            End If
            
            '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��
            If Val(Nvl(!���)) = Val(strText) Then
                Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
            End If
            '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
             If Val(Nvl(!���)) Like strText & "*" Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
            mrsChargePerson.MoveNext
        Loop
    End With
    LoadPerson = LoadPersonFromLvw(lvwPerson_S, rsTemp)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadOtherPerson(Optional blnFilter As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������շ�Ա��Ϣ
    '���:blnFilter-�Ƿ���й���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    
    On Error GoTo errHandle
    If zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�") = False Then LoadOtherPerson = True: Exit Function
    '��ȡ�շ�Ա��Ϣ
    If blnFilter = False Or mrsOtherPerson Is Nothing Then

        strSQL = " " & _
        "   Select Distinct a.Id, a.���, a.����, a.����, m.���� As ��������, a.�Ա� " & _
        "   From ��Ա�ɿ���� A1, ��Ա�� A, ������Ա C, ���ű� M " & _
        "   Where A1.�տ�Ա = a.���� And A1.���� = 1    " & _
        "         And not exists(select 1 From  ��Ա����˵�� B where a.ID=b.��ԱID And  b.��Ա����  In  ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա', '��Ժ�Ǽ�Ա', '�����Ǽ���')) " & _
        "         And a.Id = c.��Աid And c.����id = m.Id(+) And  c.ȱʡ(+) = 1" & _
        "         And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "         And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        "   Order By ���"
        Set mrsOtherPerson = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����շ�Ա��Ϣ")
    End If
    mrsOtherPerson.Filter = 0
    strText = UCase(txtOtherPerson.Text)
    If txtOtherPerson.Text = txtOtherPerson.Tag Or strText = "" Then
        'ȫ������
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    End If
    strCompents = gstrLike & strText & "%"
    If IsNumeric(strText) Then '1.�������ȫ����
    ElseIf zlCommFun.IsCharAlpha(strText) Then '1-�������ȫ��ĸ
        mrsOtherPerson.Filter = "���� like '" & gstrLike & strText & "%'"
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    Else
        intInputType = 2   '2-����
        mrsOtherPerson.Filter = "���� like '" & strText & "%'"
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    End If
    
    '�������ȫ����
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsOtherPerson)
    With mrsOtherPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '������������,��Ҫ���:
            '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
            '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
            '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
            If Nvl(!���) = strText Then
                Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp): Exit Do
            End If
            '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��
            If Val(Nvl(!���)) = Val(strText) Then
                Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp)
            End If
            '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
             If Val(Nvl(!���)) Like strText & "*" Then Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp)
            mrsOtherPerson.MoveNext
        Loop
    End With
    LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, rsTemp)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadGroup() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���������
    '���:  blnFilter-�Ƿ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-24 12:16:41
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim objItem As ListItem, strSQL As String
   Dim rsGroup As ADODB.Recordset
   Dim str������ As String
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "�������տ�") = False Then LoadGroup = True: Exit Function
    '��ȡ������
    strSQL = " " & _
    "   Select a.Id As ����, a.������, a.����, b.���� As �鸺����,A.������id ,A.˵��" & _
    "   From ����ɿ���� A, ��Ա�� B " & _
    "   Where a.������id = b.Id And Nvl(a.ɾ������, To_Date('3000-01-01', 'yyyy-mm-dd')) >= To_Date('3000-01-01', 'yyyy-mm-dd') And " & _
    "         (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And (b.վ�� = 'A' Or b.վ�� Is Null) " & _
    "   Order By a.������"
    Set mrsGroup = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������Ϣ")
    'ȫ������
    With lvwGroup_S.ListItems
        .Clear
        If mrsGroup.RecordCount <> 0 Then mrsGroup.MoveFirst
        Do While Not mrsGroup.EOF
            Set objItem = .Add(, "K" & Nvl(mrsGroup!����), Nvl(mrsGroup!������), "Group", "Group")
            objItem.SubItems(1) = Nvl(mrsGroup!����)
            objItem.SubItems(2) = Nvl(mrsGroup!����)
            strSQL = "Select B.���� From ����ɿ���� A,��Ա�� B Where (A.ɾ������ Is Null or A.ɾ������ Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And A.������ID = B.ID And A.ID = [1]"
            strSQL = strSQL & " Union Select C.���� From �������鳤���� A,����ɿ���� B,��Ա�� C Where A.��ID=B.ID And A.�鳤ID=C.ID And B.ID = [1] And (B.ɾ������ Is Null or B.ɾ������ Between Sysdate And to_date('3000-01-01','YYYY-MM-DD'))"
            Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsGroup!����)))
            str������ = ""
            Do While Not rsGroup.EOF
                str������ = str������ & "," & rsGroup!����
                rsGroup.MoveNext
            Loop
            If str������ <> "" Then str������ = Mid(str������, 2)
            objItem.SubItems(3) = str������
            objItem.SubItems(4) = Nvl(mrsGroup!˵��)
            objItem.Tag = Val(Nvl(mrsGroup!������id))
            mrsGroup.MoveNext
        Loop
    End With
    LoadGroup = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String, ByRef cbsMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:���˺�
    '����:2013-09-09 14:41:46
    '˵��:���ش����,��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    Set mcbsMain = cbsMain
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    Call InitPage: Call InitPanel
    Call LoadGroup: Call LoadPerson(False): Call LoadOtherPerson(False)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmList Is Nothing Then Unload mfrmList
    Set mfrmList = Nothing
    If Not mfrmPersonOther Is Nothing Then Unload mfrmPersonOther
    Set mfrmPersonOther = Nothing
End Sub

Private Sub lvwGroup_S_GotFocus()
    mstrSelGroup = ""
End Sub
Private Sub lvwGroup_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelGroup Then Exit Sub
    mstrSelGroup = Item.Text
    Call LoadLocalePersonDetailData
End Sub

Private Sub lvwGroup_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub

Private Sub lvwOther_S_GotFocus()
    mstrSelOther = ""
End Sub

Private Sub lvwOther_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelOther Then Exit Sub
    mstrSelOther = Item.Text
    Call LoadBalance(mstrSelOther)
    Call LoadLocalePersonDetailData
End Sub
Private Sub lvwOther_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub
Private Sub lvwPerson_S_GotFocus()
    mstrSelPerson = ""
End Sub
Private Sub lvwPerson_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelPerson Then Exit Sub
    mstrSelPerson = Item.Text
    Call LoadLocalePersonDetailData
End Sub
Private Sub lvwPerson_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub
Private Sub mfrmList_PersonChange(ByVal strPerson As String, Cancel As Boolean)
    '��Ա�ı�ʱ,��Ҫ��λ��ָ������Ա��
    Dim objItem As ListItem
    If Val(tbPage.Selected.Tag) = EM_PG_������Ա Then Cancel = True: Exit Sub
    mblnNotBrush = True
    If Val(tbPage.Selected.Tag) = EM_PG_������ Then
        For Each objItem In lvwGroup_S.ListItems
            If InStr("," & objItem.SubItems(3) & ",", "," & strPerson & ",") > 0 Then
                Call LoadBalance(strPerson): mblnNotBrush = False
                objItem.Selected = True: Exit Sub
            End If
        Next
        mblnNotBrush = False
        Cancel = True: Exit Sub
    End If
    '�շ�Ա����
    For Each objItem In lvwPerson_S.ListItems
        If objItem.Text = strPerson Then
            Call LoadBalance(strPerson): mblnNotBrush = False
            objItem.Selected = True: Exit Sub
        End If
    Next
    If txtChargePerson.Tag = txtChargePerson.Text Then
        'δ���˵�,��ʾδ�ҵ�
        mblnNotBrush = False
        Cancel = True: Exit Sub
    End If
    
    '�϶����ڰ�����/���/������ǵ�,������Ҫѡ���
    txtChargePerson.Text = "": txtChargePerson_LostFocus
   '�շ�Ա����
    For Each objItem In lvwPerson_S.ListItems
        If objItem.Text = strPerson Then
            Call LoadBalance(strPerson): mblnNotBrush = False
            objItem.Selected = True: Exit Sub
        End If
    Next
    'δ���˵�,��ʾδ�ҵ�
    mblnNotBrush = False
    Cancel = True
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        lblBalance.Left = .ScaleLeft + 50
        lblBalance.Width = .ScaleWidth - .Left * 2
        lblBalance.Top = .ScaleTop + 50
        lblBalance.Height = .ScaleHeight - .Top * 2
    End With
End Sub

 Private Sub picPersonList_Resize()
    Err = 0: On Error Resume Next
    With picPersonList
        txtChargePerson.Top = .ScaleTop + 50
        lblChargePerson.Top = txtChargePerson.Top + (txtChargePerson.Height - lblChargePerson.Height) \ 2
        txtChargePerson.Width = .ScaleWidth - txtChargePerson.Left
        lvwPerson_S.Left = .ScaleLeft
        lvwPerson_S.Top = txtChargePerson.Top + txtChargePerson.Height + 50
        lvwPerson_S.Width = .ScaleWidth
        lvwPerson_S.Height = .ScaleHeight - lvwPerson_S.Top - 50
    End With
End Sub
 Private Sub picGroupList_Resize()
    Err = 0: On Error Resume Next
    With picGroupList
        lvwGroup_S.Left = .ScaleLeft
        lvwGroup_S.Top = .ScaleTop
        lvwGroup_S.Width = .ScaleWidth
        lvwGroup_S.Height = .ScaleHeight
    End With
End Sub

 Private Sub picOtherList_Resize()
    Err = 0: On Error Resume Next
    With picOtherList
        txtOtherPerson.Top = .ScaleTop + 50
        lblOtherPerson.Top = txtOtherPerson.Top + (txtOtherPerson.Height - lblOtherPerson.Height) \ 2
        txtOtherPerson.Width = .ScaleWidth - txtOtherPerson.Left
        lvwOther_S.Left = .ScaleLeft
        lvwOther_S.Top = txtOtherPerson.Top + txtOtherPerson.Height + 50
        lvwOther_S.Width = .ScaleWidth
        lvwOther_S.Height = .ScaleHeight - lvwOther_S.Top - 50
    End With
End Sub
Private Sub picPersonPage_Resize()
    Err = 0: On Error Resume Next
    With picPersonPage
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub InitPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2013-09-22 17:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    mblnNotBrush = True
    picPersonList.Visible = zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�")
     If zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�") Then
        Set objItem = tbPage.InsertItem(EM_PG_�շ�Ա, "�շ�Ա", picPersonList.hWnd, 0)
        objItem.Tag = EM_PG_�շ�Ա
    End If
    
    picGroupList.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������տ�")
     If zlStr.IsHavePrivs(mstrPrivs, "�������տ�") Then
        Set objItem = tbPage.InsertItem(EM_PG_������, "������", picGroupList.hWnd, 0)
        objItem.Tag = EM_PG_������
     End If
    picOtherList.Visible = zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�")
    If zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�") Then
        Set objItem = tbPage.InsertItem(EM_PG_������Ա, "������Ա", picOtherList.hWnd, 0)
        objItem.Tag = EM_PG_������Ա
    End If
    
     With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        .PaintManager.Position = xtpTabPositionBottom
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    mblnNotBrush = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Sub

Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-22 17:13:23
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objMain As Pane, lngWidth As Long
    Dim lngBalanceHeight As Long
    
    lngWidth = 3435 \ Screen.TwipsPerPixelX
    lngBalanceHeight = 600 \ Screen.TwipsPerPixelY
    With dkpMan
        Set objMain = .CreatePane(mPaneIndex.EM_PN_�շ�Ա�б�, lngWidth, 400, DockLeftOf, Nothing)
        objMain.Title = ""
        objMain.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objMain.Handle = picPersonPage.hWnd
        objMain.MinTrackSize.Width = Int(lngWidth * 0.5): objMain.MaxTrackSize.Width = lngWidth
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_�ݴ��б�, 100, lngBalanceHeight, DockBottomOf, objMain)
        objPane.Title = "��ǰ�ݴ��":
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        objPane.MinTrackSize.Height = Int(lngBalanceHeight * 0.5): objPane.MaxTrackSize.Height = lngBalanceHeight * 1.5
        
        If zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�") Then
            Set mfrmPersonOther = New frmFinanceSupervisePersonOthers
            Load mfrmPersonOther
            Call mfrmPersonOther.zlInitVar(mlngModule, mstrPrivs)
        End If
        
        Set mfrmList = New frmFinaceSuperviseCollectList
        Load mfrmList
        Call mfrmList.zlInitVar(EM_TY_�շ�Ա, mlngModule, mstrPrivs)
        Set mobjDetailPane = .CreatePane(mPaneIndex.EM_PN_��ϸ�б�, 100, 100, DockRightOf, objMain)
        mobjDetailPane.Title = "":
        mobjDetailPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        mobjDetailPane.Handle = mfrmList.hWnd
        
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Public Function zlRollingCurtainCollect(ByVal frmMain As Object, Optional blnCustomCollect As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-11 11:45:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�շ�Ա As String, lng��ԱID As Long, lng�ɿ���ID As Long
    Dim strIDs As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
   Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_�շ�Ա
        If lvwPerson_S.SelectedItem Is Nothing Then Exit Function
        str�շ�Ա = lvwPerson_S.SelectedItem.Text
        lng��ԱID = Val(lvwPerson_S.SelectedItem.Tag)
        lng�ɿ���ID = 0
        If blnCustomCollect Then
            '�ֹ��տ�
            zlRollingCurtainCollect = frmFinaceSuperviseCustomInput.EditCard(frmMain, str�շ�Ա, lng��ԱID, mlngModule, mstrPrivs)
            Exit Function
        End If
        strIDs = mfrmList.GetSelRollingCurtainIds
        If strIDs = "" Then
            MsgBox "δѡ����Ҫ�տ�����ʼ�¼", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        zlRollingCurtainCollect = frmFinanceSuperviseRollingCurtainEdit.zlShowMe(frmMain, mlngModule, mstrPrivs, str�շ�Ա, lng��ԱID, strIDs, lng�ɿ���ID)
   Case EM_PG_������
        If blnCustomCollect Then
            MsgBox "�����鲻֧���ֹ��ɿ����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If lvwGroup_S.SelectedItem Is Nothing Then Exit Function
        
        str�շ�Ա = lvwGroup_S.SelectedItem.SubItems(3)
        '76120,Ƚ����,2014-8-5,û��Ȩ�ޡ��շ�Ա�տ���������տ�,��������տť����δ���ö�������� With block ������
        lng��ԱID = Val(lvwGroup_S.SelectedItem.Tag)
        lng�ɿ���ID = Val(Mid(lvwGroup_S.SelectedItem.Key, 2))
        strIDs = mfrmList.GetSelRollingCurtainIds
        If strIDs = "" Then
            MsgBox "δѡ����Ҫ�տ�����ʼ�¼", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strSQL = "Select Distinct A.�տ�Ա,B.ID From ��Ա�սɼ�¼ A,��Ա�� B Where A.�տ�Ա=B.���� And A.��¼����=3 And A.С������ID In "
        strSQL = strSQL & " (Select Column_Value From Table(f_str2list([1]))) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
        If rsTmp.RecordCount > 1 Then
            MsgBox "��ǰ���ʵĲ������տ��¼���ڶ���鳤,�޷�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        Else
            If Not rsTmp.EOF Then
                str�շ�Ա = Nvl(rsTmp!�տ�Ա)
                lng��ԱID = Val(Nvl(rsTmp!ID))
            End If
        End If
        zlRollingCurtainCollect = frmFinanceSuperviseRollingCurtainEdit.zlShowMe(frmMain, mlngModule, mstrPrivs, str�շ�Ա, lng��ԱID, strIDs, lng�ɿ���ID)
    Case EM_PG_������Ա
        If lvwOther_S.SelectedItem Is Nothing Then Exit Function
        str�շ�Ա = lvwOther_S.SelectedItem.Text
        lng��ԱID = Val(lvwOther_S.SelectedItem.Tag)
        lng�ɿ���ID = 0
        If blnCustomCollect Then
            '�ֹ��տ�
            zlRollingCurtainCollect = frmFinaceSuperviseCustomInput.EditCard(frmMain, str�շ�Ա, lng��ԱID, mlngModule, mstrPrivs, True)
            Exit Function
        End If
        zlRollingCurtainCollect = mfrmPersonOther.SaveData()
    Case Else
        Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get IsAllowCollect() As Boolean
  '�Ƿ������տ�
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_�շ�Ա
        IsAllowCollect = mfrmList.IsSelRollingCurtainRecord
   Case EM_PG_������
        IsAllowCollect = mfrmList.IsSelRollingCurtainRecord
   Case EM_PG_������Ա
        IsAllowCollect = False
   Case Else
        Exit Property
    End Select
End Property
Public Property Get IsAllowOtherCollect() As Boolean
  '�Ƿ������տ�
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_������Ա
        IsAllowOtherCollect = Not lvwOther_S.SelectedItem Is Nothing
   Case EM_PG_�շ�Ա, EM_PG_������
        IsAllowOtherCollect = False
   Case Else
        Exit Property
    End Select
End Property

Public Property Get IsAllowViewChargeList() As Boolean
  '�Ƿ�����鿴��ϸ
    Select Case Val(tbPage.Selected.Tag)
        Case EM_PG_�շ�Ա, EM_PG_������
            IsAllowViewChargeList = mfrmList.GetRollingCurtainID <> 0
        Case EM_PG_������Ա
            IsAllowViewChargeList = True
        Case Else
            Exit Property
    End Select
End Property
Public Property Get IsAllowCustomCollect() As Boolean
  '�Ƿ������ֹ��տ�
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_�շ�Ա
        IsAllowCustomCollect = Not lvwPerson_S.SelectedItem Is Nothing
   Case EM_PG_������
        IsAllowCustomCollect = False
   Case EM_PG_������Ա
        IsAllowCustomCollect = Not lvwOther_S.SelectedItem Is Nothing
   Case Else
        Exit Property
    End Select
End Property
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա
        Call mfrmList.zlPrint(bytMode)
   Case EM_PG_������
        Call mfrmList.zlPrint(bytMode)
   Case EM_PG_������Ա
        Call mfrmPersonOther.zlPrint(bytMode)
   Case Else: Exit Sub
   End Select
End Sub
Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '����:���˺�
    '����:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա
        Call mfrmList.ShowChargeList(frmMain)
   Case EM_PG_������
       Call mfrmList.ShowChargeList(frmMain)
   Case EM_PG_������Ա
       Call mfrmPersonOther.ShowChargeList(Me)
   Case Else: Exit Sub
   End Select
End Sub
Public Sub zlRefresh()
    Call LoadGroup: Call LoadPerson(False): Call LoadOtherPerson(False)
    Call LoadLocalePersonDetailData
End Sub

Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmList.CallCustomRpt(frmMain, lngSys, strRptCode)
End Sub
Public Property Get GetCashMoney() As Double
    '��ȡ�ֽ���
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա, EM_PG_������
       GetCashMoney = mfrmList.GetCashMoney
    Case EM_PG_������Ա
        GetCashMoney = mfrmPersonOther.GetCashMoney
    Case Else
    End Select
End Property
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    'ѡ��ʱ,����ȡCall LoadPerson(False)
    If mblnNotBrush Then Exit Sub
    mblnNotBrush = True
   Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_�շ�Ա
        'If lvwPerson_S.Enabled And lvwPerson_S.Visible Then lvwPerson_S.SetFocus
        If lvwPerson_S.ListItems.Count <> 0 Then lvwPerson_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmList.hWnd
        dkpMan.RecalcLayout
        mfrmPersonOther.Hide
   Case EM_PG_������
        'If lvwGroup_S.Enabled And lvwGroup_S.Visible Then lvwGroup_S.SetFocus
        If lvwGroup_S.ListItems.Count <> 0 Then lvwGroup_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmList.hWnd
        dkpMan.RecalcLayout
        mfrmPersonOther.Hide
    Case EM_PG_������Ա
        'If lvwOther_S.Enabled And lvwOther_S.Visible Then lvwOther_S.SetFocus
        If lvwOther_S.ListItems.Count <> 0 Then lvwOther_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmPersonOther.hWnd
        dkpMan.RecalcLayout
        mfrmList.Hide
   Case Else
        Exit Sub
   End Select
    mblnNotBrush = False
    lblBalance.Caption = ""
    mstrSelOther = ""
    'Call LoadLocalePersonDetailData
End Sub
Private Function LoadLocalePersonDetailData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ָ����Ա����ϸ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-26 11:29:34
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnBalance As Boolean, blnRollingCurtainMgr As Boolean
    Dim strPerson As String, lngGroupID As Long
    On Error GoTo errHandle
    '����ָ����Ա�����ʼ�¼
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_�շ�Ա
        If Not lvwPerson_S.SelectedItem Is Nothing Then
             strPerson = lvwPerson_S.SelectedItem.Text
        End If
        
        '���ص�ǰ�ݴ��
        Call mfrmList.zlClearData
        blnBalance = LoadBalance(strPerson)
        blnRollingCurtainMgr = mfrmList.zlLoadCollectData(EM_TY_�շ�Ա, strPerson)
   Case EM_PG_������
        If Not lvwGroup_S.SelectedItem Is Nothing Then
             strPerson = lvwGroup_S.SelectedItem.SubItems(3)
             lngGroupID = Val(lvwGroup_S.SelectedItem.SubItems(1))
        End If
        '���ص�ǰ�ݴ��
        Call mfrmList.zlClearData
        blnBalance = LoadBalance(strPerson)
        blnRollingCurtainMgr = mfrmList.zlLoadCollectData(EM_TY_С��, strPerson, lngGroupID)
   Case EM_PG_������Ա
        If Not lvwOther_S.SelectedItem Is Nothing Then
             strPerson = lvwOther_S.SelectedItem.Text
        End If
        '���ص�ǰ�ݴ��
        blnBalance = LoadBalance(strPerson)
        mfrmPersonOther.zlLoadPersonData (strPerson)
   Case Else
        Exit Function
   End Select
    LoadLocalePersonDetailData = blnBalance Or blnRollingCurtainMgr
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txtChargePerson_Change()
    If txtChargePerson.Text = txtChargePerson.Tag Then Exit Sub
    '���й���
    Call LoadPerson(True)
    If Not mblnNotBrush Then Call ClearData
End Sub
Private Sub txtChargePerson_GotFocus()
    If txtChargePerson.Text = txtChargePerson.Tag Then
        txtChargePerson.Text = ""
        txtChargePerson.ForeColor = lvwOther_S.ForeColor
    End If
    zlControl.TxtSelAll txtChargePerson
    zlCommFun.OpenIme False
End Sub

Private Sub txtChargePerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvwPerson_S.ListItems.Count = 1 Then
        lvwPerson_S.ListItems(1).Selected = True
        Call lvwPerson_S_ItemClick(lvwPerson_S.SelectedItem)
        If txtChargePerson.Enabled And txtChargePerson.Visible Then txtChargePerson.SetFocus
    End If
End Sub
Private Sub txtChargePerson_LostFocus()
    zlCommFun.OpenIme False
    If txtChargePerson.Text = "" Then
        txtChargePerson.ForeColor = &H80000000
        txtChargePerson.Text = "���������ֻ���"
    End If
End Sub
Private Function LoadBalance(ByVal strPerson As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ա�ݴ��
    '���:strPerson-��Ա
    '����:���سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-25 11:53:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strPreName As String
    
    On Error GoTo errHandle
    strSQL = "Select �տ�Ա,���㷽ʽ,��� From ��Ա�ɿ���� where Instr(',' || [1] || ',' ,',' || �տ�Ա || ',') > 0 and ����=1 Order By �տ�Ա"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPerson)
    With rsTemp
        strSQL = ""
        lblBalance.Caption = ""
        Do While Not .EOF
            If strPreName <> "" And strPreName <> !�տ�Ա Then
                If lblBalance.Caption <> "" Then lblBalance.Caption = lblBalance.Caption & vbCrLf
                lblBalance.Caption = lblBalance.Caption & strPreName & "���ݴ��:" & Mid(strSQL, 2)
                strSQL = ""
            End If
            strSQL = strSQL & " " & Nvl(!���㷽ʽ) & ":" & Format(Val(Nvl(!���)), "0.00")
            strPreName = !�տ�Ա
            .MoveNext
        Loop
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            If lblBalance.Caption <> "" Then lblBalance.Caption = lblBalance.Caption & vbCrLf
            lblBalance.Caption = lblBalance.Caption & strPreName & "���ݴ��:" & strSQL
        End If
    End With
    LoadBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Sub txtOtherPerson_Change()
    If txtOtherPerson.Text = txtChargePerson.Tag Then Exit Sub
    '���й���
    Call LoadOtherPerson(True)
    If Not mblnNotBrush Then Call ClearData
End Sub
Private Sub txtOtherPerson_GotFocus()
    If txtOtherPerson.Text = txtOtherPerson.Tag Then
        txtOtherPerson.Text = ""
        txtOtherPerson.ForeColor = lvwOther_S.ForeColor
    End If
    zlControl.TxtSelAll txtOtherPerson
    zlCommFun.OpenIme False
End Sub

Private Sub txtOtherPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvwOther_S.ListItems.Count <> 1 Then Exit Sub
    lvwOther_S.ListItems(1).Selected = True
    Call lvwOther_S_ItemClick(lvwOther_S.SelectedItem)
    If txtOtherPerson.Enabled And txtOtherPerson.Visible Then txtOtherPerson.SetFocus
End Sub

Private Sub txtOtherPerson_LostFocus()
    zlCommFun.OpenIme False
    If txtOtherPerson.Text = "" Then
        txtOtherPerson.ForeColor = &H80000000
        txtOtherPerson.Text = "���������ֻ���"
    End If
End Sub
Private Sub ShowPopup()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����˵�
    '����:���˺�
    '����:2013-09-27 15:21:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Set objCommandBar = mcbsMain.Add("PopupPati", xtpBarPopup)
    With objCommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����(&D)")
  End With
  If Not objCommandBar Is Nothing Then objCommandBar.ShowPopup
End Sub
 
 '��Ա�б����ʾ��ʽ
Public Property Get zlPersonListShowMode() As Integer
  Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա
        zlPersonListShowMode = lvwPerson_S.View
    Case EM_PG_������
        zlPersonListShowMode = lvwGroup_S.View
    Case EM_PG_������Ա
        zlPersonListShowMode = lvwOther_S.View
    End Select
End Property

Public Property Let zlPersonListShowMode(ByVal vNewValue As Integer)
   Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա
        lvwPerson_S.View = vNewValue
    Case EM_PG_������
        lvwGroup_S.View = vNewValue
    Case EM_PG_������Ա
        lvwOther_S.View = vNewValue
    End Select
End Property
Private Sub ClearData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���˺�
    '����:2013-09-29 11:20:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�շ�Ա, EM_PG_������
          Call mfrmList.zlClearData
    Case EM_PG_������Ա
    End Select
End Sub
