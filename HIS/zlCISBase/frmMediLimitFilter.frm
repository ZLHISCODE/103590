VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediLimitFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmMediLimitFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6630
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ҩƷ����(&0)"
      TabPicture(0)   =   "frmMediLimitFilter.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvw����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ҩƷ����(&1)"
      TabPicture(1)   =   "frmMediLimitFilter.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Chk����"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lvw����"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox Chk���� 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74880
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   675
      End
      Begin MSComctlLib.TreeView tvw���� 
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   11033
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   5925
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   10451
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5160
      TabIndex        =   1
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   0
      Top             =   7200
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLimitFilter.frx":0342
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMediLimitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mfrmMain As Form
Dim mstr���� As String
Dim mstr���� As String
Dim mstr����ID As String
Dim mlng�ⷿID As Long
Dim mstr��� As String
Dim mblnSelect As Boolean

Private Sub GetҩƷ����(ByVal lng�ⷿID As Long)
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln��ҩ�ⷿ As Boolean
    
    '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
    On Error GoTo errHandle
    bln��ҩ�ⷿ = False
    gstrSql = "Select 1 From ��������˵�� " & _
             " Where �������� Like '��ҩ%' And ����ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption & "[��鲿������]", lng�ⷿID)

    If Not rsTemp.EOF Then bln��ҩ�ⷿ = True
    
    gstrSql = "Select Distinct J.����,J.���� " & _
             " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
             " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
             " And A.ִ�п���ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿID)
    Lvw����.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!���� = "����")
            End If
            Lvw����.ListItems.Add , "K" & !����, !����, , 1
            .MoveNext
        Loop
        If bln��ҩ�ⷿ And blnEXIST = False Then
            Lvw����.ListItems.Add , "KK1", "����", , 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk����_Click()
    If Chk����.Value = 2 Then Exit Sub
    Call SetSelect(Lvw����, Chk����.Value)
End Sub


Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim blnAllSelect As Boolean
    
    'ȡ�ü��ͣ���ѡ��������ȡҩƷ����Ҫ�ֹ�¼�룩
    mstr���� = ""
    
    If Chk����.Value = 1 Then
        mstr���� = ""
    Else
        intItems = Me.Lvw����.ListItems.count
        blnAllSelect = True
        For intItem = 1 To intItems
            If Lvw����.ListItems(intItem).Checked Then
                mstr���� = mstr���� & "," & Lvw����.ListItems(intItem).Text
            Else
                blnAllSelect = False
            End If
        Next
    
        If mstr���� <> "" Then mstr���� = Mid(mstr����, 2)
        If blnAllSelect = True Then mstr���� = ""
    End If

    'ȡ��ҩƷ���ࣨ��ѡ�����ʾ���з��ࣩ
    mstr����ID = ""
    mstr���� = ""
    For intItem = 1 To tvw����.Nodes.count
        If tvw����.Nodes(intItem).Key = "Root" And tvw����.Nodes(intItem).Checked = True Then
            mstr���� = "����"
            mstr����ID = ""
            Exit For
        ElseIf tvw����.Nodes(intItem).Key <> "Root" And _
            tvw����.Nodes(intItem).Key <> "_�г�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_�в�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_����ҩ" And _
            tvw����.Nodes(intItem).Checked Then
            mstr���� = mstr���� & "," & tvw����.Nodes(intItem).Text
            mstr����ID = mstr����ID & "," & Mid(tvw����.Nodes(intItem).Key, 2)
        End If
    Next

    If mstr����ID <> "" Then mstr����ID = Mid(mstr����ID, 2)
     
    If mstr���� <> "" Then mstr���� = Mid(mstr����, 2)
    
    mblnSelect = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    GetҩƷ���� mlng�ⷿID
    GetҩƷ���� mlng�ⷿID
End Sub
Private Sub tvw����_NodeCheck(ByVal node As MSComctlLib.node)
    CheckNode node, node.Checked
    SetParentNode node, node.Checked
End Sub

Private Sub SetParentNode(ByVal node As MSComctlLib.node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = node.FirstSibling.Index
            Do While intIdx <> node.LastSibling.Index
                If tvw����.Nodes(intIdx).Checked = False Then
                    node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw����.Nodes(intIdx).Next.Index
            Loop
            If intIdx = node.LastSibling.Index Then
                If tvw����.Nodes(intIdx).Checked = True Then
                    node.Parent.Checked = True
                End If
            End If
        Else
            node.Parent.Checked = False
        End If
        
        Set node = node.Parent
        If Not node Is Nothing Then
            SetParentNode node, blnCheck
        End If
    End If
End Sub

Private Function CheckNode(ByVal node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If node.Children > 0 Then
        Set node = node.Child
        Do While Not node Is Nothing
            node.Checked = blnCheck
            If node.Children > 0 Then
                CheckNode node, blnCheck
            End If
            Set node = node.Next
        Loop
    Else
        node.Checked = blnCheck
    End If
End Function

Public Function GetCondition(FrmMain As Form, ByVal lng�ⷿID As Long, ByVal str��� As String, ByRef str���� As String, ByRef str����ID As String, ByRef str���� As String) As Boolean
    mstr���� = ""
    mstr���� = ""
    mstr����ID = ""
    mblnSelect = False
    
    mstr��� = str���
    mlng�ⷿID = lng�ⷿID
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    str���� = mstr����
    str���� = mstr����
    str����ID = mstr����ID
End Function

Private Sub GetҩƷ����(ByVal lng�ⷿID As Long)
    Dim rsData As ADODB.Recordset
    Dim strCon As String
    Dim objNode As node
    Dim str�������� As String
    Dim int���� As Integer
    
'    str�������� = Get��������(mlng�ⷿID)
'
'    If InStr(1, str��������, "��ҩ") > 0 And InStr(1, str��������, "��ҩ") > 0 And InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� In (1, 2, 3)"
'    ElseIf InStr(1, str��������, "��ҩ") > 0 And InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� In (1, 2)"
'    ElseIf InStr(1, str��������, "��ҩ") > 0 And InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� In (2, 3)"
'    ElseIf InStr(1, str��������, "��ҩ") > 0 And InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� In (1, 3)"
'    ElseIf InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� =1 "
'    ElseIf InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� =2 "
'    ElseIf InStr(1, str��������, "��ҩ") > 0 Then
'        strCon = "���� =3 "
'    ElseIf InStr(1, str��������, "�Ƽ���") > 0 Then
'        strCon = "���� In (1, 2, 3) "
'    End If
    
    If mstr��� = "5" Then
        int���� = 1
    ElseIf mstr��� = "6" Then
        int���� = 2
    ElseIf mstr��� = "7" Then
        int���� = 3
    End If
    
    On Error GoTo errHandle
    gstrSql = "Select Level as ��,ID,�ϼ�ID,����,DECODE(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') As ���� " & _
        " From ���Ʒ���Ŀ¼ " & _
        " Where ����=[1] " & _
        " Start With �ϼ�id Is Null " & _
        " Connect By Prior ID = �ϼ�id"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "ȡҩƷ����", int����)
    
    tvw����.Nodes.Clear
    Set objNode = tvw����.Nodes.Add(, , "Root", "���з���", 1)
'    If InStr(1, str��������, "��ҩ") > 0 Or InStr(1, str��������, "�Ƽ���") > 0 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", 1)
'    If InStr(1, str��������, "��ҩ") > 0 Or InStr(1, str��������, "�Ƽ���") > 0 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)
'    If InStr(1, str��������, "��ҩ") > 0 Or InStr(1, str��������, "�Ƽ���") > 0 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
    
    If int���� = 1 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", 1)
    If int���� = 2 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)
    If int���� = 3 Then Set objNode = tvw����.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
    
    Do While Not rsData.EOF
        If rsData!�� = 1 Then
            Set objNode = tvw����.Nodes.Add("_" & rsData!����, 4, "_" & rsData!ID, rsData!����, 1)
        Else
            Set objNode = tvw����.Nodes.Add("_" & rsData!�ϼ�ID, 4, "_" & rsData!ID, rsData!����, 1)
        End If
        rsData.MoveNext
    Loop

    tvw����.Nodes("Root").Selected = True
    tvw����.Nodes("Root").Expanded = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get��������(ByVal lng����ID As Long) As String
    Dim rsData As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    gstrSql = "Select �������� From ��������˵�� Where ����id = [1]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "ȡ���Ź�������", lng����ID)
    
    With rsData
        Do While Not .EOF
            strTmp = IIf(strTmp = "", "", strTmp & ";") & !��������
            .MoveNext
        Loop
    End With
    
    Get�������� = strTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


