VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��ٱ༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ٱ༭"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frm��ٱ༭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComCtl2.DTPicker dtp��ʼ���� 
      Height          =   300
      Left            =   2820
      TabIndex        =   5
      Top             =   3210
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   90046467
      CurrentDate     =   38433
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5730
      Top             =   90
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
            Picture         =   "frm��ٱ༭.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   5070
      TabIndex        =   4
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��ٽ���(&E)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   150
      TabIndex        =   3
      Top             =   3150
      Width           =   1245
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "�������(&A)"
      Height          =   405
      Left            =   1500
      TabIndex        =   2
      Top             =   3150
      Width           =   1245
   End
   Begin MSComctlLib.ListView lvw��ټ�¼ 
      Height          =   2325
      Left            =   150
      TabIndex        =   1
      Top             =   720
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��ٽ�����ˮ��"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��ʼ����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��������"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frm��ٱ༭.frx":1D16
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblPatient 
      Caption         =   "����:�Ա�:ҽ����:��Ժ����"
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   270
      Width           =   5955
   End
End
Attribute VB_Name = "frm��ٱ༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnStart As Boolean
Private mstrInput As String
Private mlng����ID As Long
Private mstr��ˮ�� As String
Private rsTemp As New ADODB.Recordset
Private cn���� As New ADODB.Connection

Public Sub ShowEditor(ByVal lng����ID As Long)
    mlng����ID = lng����ID
    Me.Show 1
End Sub

Private Sub cmdADD_Click()
    Dim blnEnabled As Boolean
    cmdEdit.Enabled = False
    
    With lvw��ټ�¼
        If .ListItems.Count <> 0 Then
            If Not .SelectedItem Is Nothing Then
                blnEnabled = (.SelectedItem.SubItems(2) = "")
            End If
        End If
    End With
    
    If dtp��ʼ����.Visible Then
        dtp��ʼ����.Visible = False
        cmdEdit.Enabled = blnEnabled
    Else
        dtp��ʼ����.Visible = True
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim blnTrans As Boolean
    Dim str��ʼ���� As String
    Dim datCurr As Date
    Dim str������ˮ�� As String
    On Error GoTo errHand
    
    str������ˮ�� = lvw��ټ�¼.SelectedItem.Text
    str��ʼ���� = lvw��ټ�¼.SelectedItem.SubItems(1)
    
    cn����.BeginTrans
    blnTrans = True
    
    '�����µ���ٵǼǼ�¼�������ýӿ�
    gstrSQL = "zl_��ٵǼǼ�¼_END('" & mstr��ˮ�� & "','" & str������ˮ�� & "')"
    cn����.Execute gstrSQL, , adCmdStoredProc
    
    mstrInput = mstr��ˮ�� & "|" & str������ˮ�� & "|" & str��ʼ���� & "|" & Format(datCurr, "yyyyMMdd")
    Call ���ýӿ�_׼��_����������("33", mstrInput)
    If Not ���ýӿ�_���������� Then
        cn����.RollbackTrans
        Exit Sub
    End If
    
    cn����.CommitTrans
    blnTrans = False
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then cn����.RollbackTrans
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub dtp��ʼ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim datCurr As Date
    Dim blnTrans As Boolean
    Dim str��ʼ���� As String
    Dim str������ˮ�� As String
    On Error GoTo errHand
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If MsgBox("��ȷ��Ҫ������ٿ�ʼ�Ǽ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '�������
    If Format(dtp��ʼ����.Value, "yyyy-MM-dd") > Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
        MsgBox "��ٿ�ʼ���ڲ��ܴ��ڵ�ǰ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    str��ʼ���� = Format(dtp��ʼ����.Value, "yyyy-MM-dd")
    
    gstrSQL = "Select 1 From ��ٵǼǼ�¼ Where �������� Is NULL"
    Call OpenRecordset(rsTemp, "����Ƿ�������δ�����ļ�¼", gstrSQL, cn����)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "��ǰ���˻��������״̬��,���ܼ���������ٵǼǣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ٿ�ʼ,����Ƿ�������δ�����ļ�¼
    cn����.BeginTrans
    blnTrans = True
    
    '�����µ���ٵǼǼ�¼�������ýӿ�
    datCurr = zlDatabase.Currentdate()
    gstrSQL = "zl_��ٵǼǼ�¼_START('" & mstr��ˮ�� & "','" & Format(datCurr, "yyyyMMddHHmmss") & "',to_Date('" & str��ʼ���� & "','yyyy-MM-dd'))"
    cn����.Execute gstrSQL, , adCmdStoredProc
    
    mstrInput = mstr��ˮ�� & "|" & Format(datCurr, "yyyyMMddHHmmss") & "|" & Format(str��ʼ����, "yyyyMMdd") & "|"
    Call ���ýӿ�_׼��_����������("33", mstrInput)
    If Not ���ýӿ�_���������� Then
        cn����.RollbackTrans
        Exit Sub
    End If
    
    cn����.CommitTrans
    blnTrans = False
    
    Call cmdADD_Click
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then cn����.RollbackTrans
End Sub

Private Sub Form_Activate()
    If blnStart = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strPatient As String
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTmp As ADODB.Recordset
    blnStart = False
    
    If Not Initҽ�� Then Exit Sub
    
    'ȡ���˵���ˮ��
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
        gstrSQL = " Select A.��ˮ��,B.����,B.�Ա�,A.ҽ����,C.��Ժ���� " & _
              " From �����ʻ� A,������Ϣ B,������ҳ C" & _
              " Where A.����=" & TYPE_���������� & " And A.����ID=" & mlng����ID & _
              " And A.����ID=B.����ID And B.����ID=C.����ID And B.��ҳID=C.��ҳID"
    Else
        gstrSQL = " Select A.��ˮ��,B.����,B.�Ա�,A.ҽ����,C.��Ժ���� " & _
              " From �����ʻ� A,������Ϣ B,������ҳ C" & _
              " Where A.����=" & TYPE_���������� & " And A.����ID=" & mlng����ID & _
              " And A.����ID=B.����ID And B.����ID=C.����ID And B.סԺ����=C.��ҳID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵ľ�����ˮ��")
    mstr��ˮ�� = Nvl(rsTemp!��ˮ��)
    strPatient = "����:" & rsTemp!���� & "|" & "�Ա�:" & Nvl(rsTemp!�Ա�) & "|" & "ҽ����:" & rsTemp!ҽ���� & "|" & "��Ժ����:" & Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    lblPatient.Caption = strPatient
    Me.dtp��ʼ����.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Call RefreshData
    
    blnStart = True
End Sub

Private Sub RefreshData()
    Dim lvwItem As ListItem
    '��ȡ���ò��˱���סԺ��������ټ�¼
    gstrSQL = "Select ��ٽ�����ˮ��,��ʼ����,�������� From ��ٵǼǼ�¼ Where ������ˮ��='" & mstr��ˮ�� & "' Order By ��ٽ�����ˮ��"
    Call OpenRecordset(rsTemp, "��ȡ���ò��˱���סԺ��������ټ�¼", gstrSQL, cn����)
    With rsTemp
        lvw��ټ�¼.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvw��ټ�¼.ListItems.Add(, "K_" & .AbsolutePosition, !��ٽ�����ˮ��, , 1)
            lvwItem.SubItems(1) = Format(!��ʼ����, "yyyyMMdd")
            If Not IsNull(!��������) Then
                lvwItem.SubItems(2) = Format(!��������, "yyyyMMdd")
            End If
            .MoveNext
        Loop
        
        If .RecordCount <> 0 Then
            lvw��ټ�¼.ListItems(1).Selected = True
            lvw��ټ�¼.SelectedItem.Selected = True
            Call lvw��ټ�¼_ItemClick(lvw��ټ�¼.SelectedItem)
        End If
    End With
End Sub

Private Sub lvw��ټ�¼_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '�κ��������������
    '��ǰ��¼�޽���ʱ��ʱ,�������
    cmdEdit.Enabled = False
    If dtp��ʼ����.Visible Then Exit Sub
    
    With lvw��ټ�¼
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        cmdEdit.Enabled = (Item.SubItems(2) = "")
    End With
End Sub

Private Function Initҽ��() As Boolean
    Dim strUser As String, strServer As String, strPass As String
    
    '��������ҽ��������������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_����������
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(cn����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    Initҽ�� = True
End Function
