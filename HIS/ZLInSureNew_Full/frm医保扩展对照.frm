VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmҽ����չ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ø�������"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmҽ����չ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4230
      TabIndex        =   12
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdɾ�� 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   4230
      TabIndex        =   14
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmd�޸� 
      Caption         =   "�޸�(&M)"
      Height          =   350
      Left            =   2970
      TabIndex        =   13
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2970
      TabIndex        =   11
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&N)"
      Height          =   350
      Left            =   1710
      TabIndex        =   10
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   30
      TabIndex        =   17
      Top             =   2400
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Height          =   5205
      Left            =   5670
      TabIndex        =   16
      Top             =   -150
      Width           =   30
   End
   Begin VB.ComboBox cbo��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   5940
      TabIndex        =   15
      Top             =   540
      Width           =   1100
   End
   Begin VB.TextBox txt˵�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   8
      Top             =   1500
      Width           =   3495
   End
   Begin VB.TextBox txtҽ����Ŀ��Ϣ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   3
      Top             =   690
      Width           =   3495
   End
   Begin VB.TextBox txt��Ŀ��Ϣ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   1
      Top             =   300
      Width           =   3495
   End
   Begin MSComctlLib.ListView lvwAdvance 
      Height          =   1755
      Left            =   270
      TabIndex        =   9
      Top             =   2580
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3096
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��Ŀ����"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ŀ����"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   3810
      EndProperty
   End
   Begin VB.CommandButton cmdҽ����Ŀ��Ϣ 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5010
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lbl��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lbl˵�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblҽ����Ŀ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����Ŀ��Ϣ(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label lbl��Ŀ��Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HIS��Ŀ��Ϣ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frmҽ����չ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mlng�շ�ϸĿID As Long
Private mrsTemp As New ADODB.Recordset

Private Sub cmd����_Click()
    If Not IsValid Then Exit Sub
    If Not SaveData Then Exit Sub
    
    Call SetConsEnable(False)
    Call RefreshData
End Sub

Private Sub cmdȡ��_Click()
    Call SetConsEnable(False)
    If lvwAdvance.ListItems.Count <> 0 Then Call lvwAdvance_ItemClick(lvwAdvance.ListItems(1))
End Sub

Private Sub cmdɾ��_Click()
    On Error GoTo errHand
    If lvwAdvance.ListItems.Count = 0 Then Exit Sub
    If lvwAdvance.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("��ȷ��Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "ZL_ҽ��������ϸ_Delete(" & mint���� & "," & mlng�շ�ϸĿID & ",'" & lvwAdvance.SelectedItem.SubItems(1) & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd�˳�_Click()
    Unload Me
End Sub

Private Sub cmd����_Click()
    Call SetConsEnable(True)
    
    txtҽ����Ŀ��Ϣ.Text = ""
    cbo���.ListIndex = 0
    txt˵��.Text = ""
    txtҽ����Ŀ��Ϣ.SetFocus
End Sub

Private Sub cmd�޸�_Click()
    Call SetConsEnable(True)
    
    txtҽ����Ŀ��Ϣ.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    '��ȡHIS��Ŀ�ı���������
    gstrSQL = "Select '['||����||']'||���� AS ��Ŀ��Ϣ From �շ�ϸĿ Where ID=[1]"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡHIS��Ŀ�ı���������", mlng�շ�ϸĿID)
    Me.txt��Ŀ��Ϣ.Text = mrsTemp!��Ŀ��Ϣ
    
    '��ȡҽ���������
    gstrSQL = "Select ����,���� From ҽ��������� Where ����=[1] And Nvl(����,0)<>0 Order by ����"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ���������", mint����)
    With mrsTemp
        Me.cbo���.Clear
        Do While Not .EOF
            cbo���.AddItem !����
            cbo���.ItemData(cbo���.NewIndex) = !����
            .MoveNext
        Loop
        cbo���.ListIndex = 0
    End With
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ShowEditor(ByVal int���� As Integer, ByVal lng�շ�ϸĿID As Long)
    mint���� = int����
    mlng�շ�ϸĿID = lng�շ�ϸĿID
    Me.Show 1
End Sub

Private Sub SetConsEnable(ByVal blnEnable As Boolean)
    cmd����.Enabled = blnEnable
    cmdȡ��.Enabled = blnEnable
    txtҽ����Ŀ��Ϣ.Enabled = blnEnable
    cmdҽ����Ŀ��Ϣ.Enabled = blnEnable
    cbo���.Enabled = blnEnable
    txt˵��.Enabled = blnEnable
    
    cmd����.Enabled = Not blnEnable
    cmd�޸�.Enabled = Not blnEnable
    cmdɾ��.Enabled = Not blnEnable
    lvwAdvance.Enabled = Not blnEnable
End Sub

Private Sub lvwAdvance_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intDO As Integer, intCOUNT As Integer
    With lvwAdvance
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        txtҽ����Ŀ��Ϣ.Text = "[" & .SelectedItem.SubItems(1) & "]" & .SelectedItem.SubItems(2)
        txtҽ����Ŀ��Ϣ.Tag = .SelectedItem.SubItems(1)
        txt˵��.Text = .SelectedItem.SubItems(3)
        
        intCOUNT = cbo���.ListCount
        For intDO = 1 To intCOUNT
            If Val(.SelectedItem.Tag) = cbo���.ItemData(intDO - 1) Then
                cbo���.ListIndex = intDO - 1
                Exit For
            End If
        Next
    End With
End Sub

Private Function IsValid() As Boolean
    If txtҽ����Ŀ��Ϣ.Tag = "" Then
        MsgBox "��ѡ��ҽ����Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    IsValid = True
End Function

Private Function SaveData() As Boolean
    On Error GoTo errHand
    gstrSQL = "ZL_ҽ��������ϸ_Modify(" & mint���� & "," & mlng�շ�ϸĿID & "," & Me.cbo���.ItemData(Me.cbo���.ListIndex) & ",'" & txtҽ����Ŀ��Ϣ.Tag & "','" & txt˵��.Text & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    SaveData = True
    
    Call RefreshData
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub RefreshData()
    Dim lvwItem As ListItem
    '��ȡ����ɵĶ�����Ϣ
    gstrSQL = "Select A.��� AS ������,B.���� AS �������,A.�շ�ϸĿID,A.��Ŀ����,C.���� AS ��Ŀ����,A.˵�� " & _
        " From ҽ��������ϸ A,ҽ��������� B,������Ŀ C" & _
        " Where A.����=B.���� And A.����=[1] And A.�շ�ϸĿID=[2]" & _
        " And C.����=A.���� And C.����=A.��Ŀ���� And A.���=B.���� And B.����<>0" & _
        " Order by A.���,A.��Ŀ����"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ɵĶ�����Ϣ", mint����, mlng�շ�ϸĿID)
    With mrsTemp
        lvwAdvance.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvwAdvance.ListItems.Add(, "K_" & lvwAdvance.ListItems.Count, !�������)
            lvwItem.SubItems(1) = !��Ŀ����
            lvwItem.SubItems(2) = !��Ŀ����
            lvwItem.SubItems(3) = Nvl(!˵��)
            lvwItem.Tag = !������
            .MoveNext
        Loop
    End With
    
    If lvwAdvance.ListItems.Count <> 0 Then Call lvwAdvance_ItemClick(lvwAdvance.ListItems(1))
End Sub

Private Sub txtҽ����Ŀ��Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnReturn As Boolean
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    StrInput = UCase(Trim(txtҽ����Ŀ��Ϣ.Text))
    If StrInput = "" Then
        MsgBox "������ҽ����Ŀ��Ϣ!", vbInformation, gstrSysName
        Exit Sub
    End If
    If Mid(StrInput, 1, 1) = "[" Then
        If InStr(2, StrInput, "]") <> 0 Then
            StrInput = Mid(StrInput, 2, InStr(2, StrInput, "]") - 2)
        Else
            StrInput = Mid(StrInput, 2)
        End If
    End If
    
    gstrSQL = "Select ����,����,����,��ע From ������Ŀ " & _
        " Where ����=[1]" & _
        " And (���� Like [2] || '%' Or ���� Like [2] || '%' Or Upper(����) Like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ", mint����, StrInput)
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ�ƥ���ҽ����Ŀ,����������!", vbInformation, gstrSysName
        Exit Sub
    End If
    If rsTemp.RecordCount > 1 Then
        blnReturn = frmListSel.ShowSelect(mint����, rsTemp, "����", "ҽ����Ŀѡ��", "��ѡ���Ӧ��ҽ����Ŀ��")
    Else
        blnReturn = True
    End If
    If blnReturn Then
        txtҽ����Ŀ��Ϣ.Text = "[" & rsTemp!���� & "]" & rsTemp!����
        txtҽ����Ŀ��Ϣ.Tag = rsTemp!����
    End If
End Sub
