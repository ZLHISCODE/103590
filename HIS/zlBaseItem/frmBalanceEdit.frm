VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���㷽ʽ�༭"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmBalanceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5010
      TabIndex        =   14
      Top             =   4560
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "Ӧ�ó���"
      Height          =   3390
      Left            =   150
      TabIndex        =   9
      Top             =   1935
      Width           =   4680
      Begin MSComctlLib.ListView lvw���� 
         Height          =   2250
         Left            =   105
         TabIndex        =   11
         Top             =   1020
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   3969
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "_Ӧ�ó���"
            Object.Tag             =   "Ӧ�ó���"
            Text            =   "Ӧ�ó���"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "_ȱʡ��־"
            Object.Tag             =   "ȱʡ��־"
            Text            =   "ȱʡ��־"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label lbl��ʾ 
         Caption         =   $"frmBalanceEdit.frx":000C
         Height          =   735
         Left            =   195
         TabIndex        =   10
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5010
      TabIndex        =   13
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5010
      TabIndex        =   12
      Top             =   255
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1725
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4635
      Begin VB.CheckBox chkӦ���� 
         Caption         =   "Ӧ����"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1328
         Width           =   855
      End
      Begin VB.CheckBox chkDue 
         Caption         =   "Ӧ�տ�"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   1328
         Width           =   975
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         ItemData        =   "frmBalanceEdit.frx":0095
         Left            =   840
         List            =   "frmBalanceEdit.frx":0097
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1305
         Width           =   1850
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   210
         Width           =   405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         Top             =   570
         Width           =   3675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   4
         TabIndex        =   6
         Top             =   930
         Width           =   1850
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&Q)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���(&U)"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   990
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmBalanceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr���� As String        'ԭʼ�ı���
Dim mstr���� As String        'ԭʼ������
Dim mbln�̶� As Boolean       '��ʽ�Ƿ�̶�
Dim mblnItem As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mintSuccess As Integer
Dim mblnCancel As Boolean     'ȡ���༭

Private Function CheckUsedDue() As Boolean
'����Ƿ�ѡ���˽��ʳ���
    Dim i As Long
    If cmb.ListIndex <> -1 Then
        If cmb.ListIndex = 0 Or cmb.ListIndex = 1 Or cmb.ListIndex = 3 Then
            For i = 1 To lvw����.ListItems.Count
                If lvw����.ListItems(i).Checked = True Then
                    If lvw����.ListItems(i).Text = "����" Then CheckUsedDue = True: Exit Function
                End If
            Next
        End If
    End If
End Function
Private Function IsCheckDueValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ�����շѻ���ʳ���,����δʹ����������
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-04 11:01:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnUserOther As Boolean, blnUser As Boolean
    IsCheckDueValied = False
    If cmb.ListIndex <> -1 Then
        If cmb.ListIndex = 0 Or cmb.ListIndex = 1 Then
            With lvw����
                For i = 1 To .ListItems.Count
                    If .ListItems(i).Checked = True Then
                        If InStr(1, ";�շ�;����;", ";" & .ListItems(i).Text & ";") > 0 Then
                            blnUser = True
                        Else
                            blnUserOther = True: Exit For
                        End If
                    End If
                Next
                IsCheckDueValied = Not blnUserOther And blnUser '����Ƿ�ʹ�����շѻ���ʳ���,����δʹ����������
            End With
        End If
    End If
End Function


Private Sub chkӦ����_Click()
    Dim rsTmp As ADODB.Recordset, ObjItem As ListItem
    'Ӧ���ʽֻ����һ��:33722
    If chkӦ����.value = 1 Then
        mblnItem = True
        '��Ҫ����Ƿ�ֻ���շѺͽ���
        For Each ObjItem In Me.lvw����.ListItems
            If InStr(1, ";�շ�;����;", ";" & ObjItem.Text & ";") > 0 Then
                ObjItem.Checked = True
                ObjItem.Selected = True
            Else
                ObjItem.Checked = False
            End If
            ObjItem.SubItems(1) = ""
        Next
        mblnItem = False
    End If
End Sub

Private Sub cmb_Click()
    Dim ObjItem As ListItem
    Dim rsTmp As New ADODB.Recordset
    
    If mblnCancel Then Exit Sub
    mblnChange = True
        
    On Error GoTo ErrHandle
    chkDue.Enabled = CheckUsedDue
    chkӦ����.Enabled = IsCheckDueValied    '�Ƿ�Ҫʹ��Ӧ�������ѡ��:33722
    
    If Not chkӦ����.Enabled Then chkӦ����.value = 0
    If Not chkDue.Enabled Then chkDue.value = 0
    
    '�ֽ�ʽֻ����һ��
    If Trim(mstr����) <> "" Then
        gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where ����<>[1] and nvl(����,1)=1 "
    Else
        gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where  nvl(����,1)=1 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
        
    If rsTmp.RecordCount > 0 Then
        If cmb.ListIndex + 1 = 1 Then
            mblnCancel = True
            cmb.ListIndex = 1
            mblnCancel = False
            Exit Sub
        End If
    End If
    
    '�����ʻ�ֻ����һ��
    If Trim(mstr����) <> "" Then
        gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where ����<>[1] and  nvl(����,1)=3 "
    Else
        gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where  nvl(����,1)=3 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
    
    If rsTmp.RecordCount > 0 Then
        If cmb.ListIndex + 1 = 3 Then
            mblnCancel = True
            cmb.ListIndex = 3
            For Each ObjItem In Me.lvw����.ListItems
                ObjItem.SubItems(1) = ""
                If ObjItem.Text = "������" Or ObjItem.Text = "���ѿ�" Then
                    ObjItem.Checked = False
                End If
            Next
            mblnCancel = False
            Exit Sub
        End If
    End If
    
    If cmb.ListIndex = 4 Then
        '���տ���ֻ��Ӧ����Ԥ����,�Ҳ���Ϊȱʡ
        For Each ObjItem In Me.lvw����.ListItems
            If ObjItem.Text = "Ԥ����" Then
                ObjItem.Checked = True
                ObjItem.Selected = True
            Else
                ObjItem.Checked = False
            End If
            ObjItem.SubItems(1) = ""
        Next
    ElseIf cmb.ListIndex = 2 Or cmb.ListIndex = 3 Then
        'ҽ���Ľ��㷽ʽ����Ϊȱʡ���㷽ʽ
        For Each ObjItem In Me.lvw����.ListItems
            ObjItem.SubItems(1) = ""
            If ObjItem.Text = "������" Or ObjItem.Text = "���ѿ�" Then
                ObjItem.Checked = False
            End If
        Next
    ElseIf cmb.ListIndex = 6 Then 'һ��ͨ������Ԥ����;��￨���Լ�����������ѿ�
        For Each ObjItem In Me.lvw����.ListItems
            If ObjItem.Text = "���￨" Or ObjItem.Text = "Ԥ����" Or ObjItem.Text = "������" Or ObjItem.Text = "���ѿ�" Then
                ObjItem.Checked = False
            End If
        Next
    ElseIf cmb.ListIndex = 5 Then
        For Each ObjItem In Me.lvw����.ListItems
            If ObjItem.Text = "������" Or ObjItem.Text = "���ѿ�" Then
                ObjItem.Checked = False
            End If
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save����() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstr���� <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstr���� = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = zlDatabase.GetMax("���㷽ʽ", "����", 2)
    Me.cmb.ListIndex = 0
    For i = 1 To lvw����.ListItems.Count
        lvw����.ListItems(i).Checked = False
        lvw����.ListItems(i).SubItems(1) = ""
    Next
    mblnChange = False
    txtEdit(1).SetFocus
    frmBalanceManage.Fill���㷽ʽ
End Sub

Private Function IsValid() As Boolean
    '����:���������йؽ��㷽ʽ�������Ƿ���Ч
    '����:
    '����ֵ:��Ч����True,����ΪFalse
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    For i = 1 To 3
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "���������ݲ��ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If Len(txtEdit(1).Text) = 0 Then
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If chkӦ����.value = 1 Then
        If Trim(mstr����) <> "" Then
            gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where ����<>[1] and nvl(Ӧ����,0)=1 "
        Else
            gstrSQL = "select ����,����,����,nvl(����,1) ����,ȱʡ��־ from ���㷽ʽ  where  nvl(Ӧ����,0)=1 "
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
        If Not rsTmp.EOF Then
             If MsgBox("ע��:" & vbCrLf & _
                              "     ����Ӧ�������ʵĽ��㷽ʽֻ��һ��,����ϵͳ�д���" & vbCrLf & _
                              "     ���㷽ʽΪ��" & Nvl(rsTmp!����) & "����Ӧ����,�����������,�������" & vbCrLf & _
                              "     ���㷽ʽ��Ӧ��������,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                 Exit Function
            End If
        End If
        If Trim(mstr����) <> "" Then
            gstrSQL = "select A.���� from ���㷽ʽ A,���㷽ʽӦ�� B   where (A.����=[1]) and a.����=b.���㷽ʽ and nvl(b.ȱʡ��־,0)=1 "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
            If rsTmp.EOF = False Then
                 If MsgBox("ע��:" & vbCrLf & _
                                  "     ����Ӧ�������ʵĽ��㷽ʽ�������ó�ȱʡ,�����޸ĵ�" & vbCrLf & _
                                  "     ���㷽ʽΪ��" & Nvl(rsTmp!����) & "��Ŀǰ������ȱʡ״̬,�����������,�������" & vbCrLf & _
                                  "     �˽��㷽ʽ������ȱʡ��־,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                     Exit Function
                End If
            End If
        End If
    End If
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Save����() As Boolean
'����:����༭�����ݽ��㷽ʽ����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim i As Integer
    Dim str���� As String
    On Error GoTo ErrHandle
    '������ѡ�еĹ�����������һ����
    '������ѡ�еĹ�����������һ����
    For i = 1 To lvw����.ListItems.Count
        If lvw����.ListItems(i).Checked = True Then
            str���� = str���� & lvw����.ListItems(i) & ":"
            If chkӦ����.value = 1 Then
                str���� = str���� & "0;"
            Else
                str���� = str���� & IIF(lvw����.ListItems(i).SubItems(1) = "", "0;", "1;")
            End If
            
        End If
    Next
    
    If mstr���� = "" Then       '����һ����¼
        gstrSQL = "zl_���㷽ʽ_insert( '" & _
            txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & cmb.ListIndex + 1 & ",'" & str���� & "'," & IIF(chkDue.Enabled, chkDue.value, 0) & "," & IIF(chkӦ����.Enabled, chkӦ����.value, 0) & ")"
    Else    '�޸�
        gstrSQL = "zl_���㷽ʽ_update( '" & mstr���� & "','" & _
            txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & cmb.ListIndex + 1 & ",'" & str���� & "'," & IIF(chkDue.Enabled, chkDue.value, 0) & "," & IIF(chkӦ����.Enabled, chkӦ����.value, 0) & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭���㷽ʽ(ByVal str���� As String) As Boolean
'����:��������õĽ��㷽ʽ�����ڽ���ͨѶ�ĳ���
'����:str����     ��ǰ�༭�Ľ��㷽ʽ�ı���
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rs���㷽ʽ As New ADODB.Recordset
    
    '�õ���������ݳ���
    GetDefineSize
    
    mintSuccess = 0
    rs���㷽ʽ.CursorLocation = adUseClient
    rs���㷽ʽ.CursorType = adOpenKeyset
    rs���㷽ʽ.LockType = adLockReadOnly
    
    On Error GoTo ErrHandle
    mstr���� = str����
    mstr���� = ""
    If str���� <> "" Then
        gstrSQL = "select ����,����,����,nvl(����,1) ����,�Ƿ�̶�,ȱʡ��־,Nvl(Ӧ�տ�,0) Ӧ�տ�,Nvl(Ӧ����,0) Ӧ���� from ���㷽ʽ  where ����=[1]"
        Set rs���㷽ʽ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
                
        '75134:���ϴ�,2014/7/14,�������Ϊ9��ֱ���˳�
        If rs���㷽ʽ!���� = 9 Then Exit Function
        txtEdit(1).Text = mstr����
        txtEdit(2).Text = rs���㷽ʽ("����")
        mstr���� = rs���㷽ʽ("����")
        txtEdit(3).Text = IIF(IsNull(rs���㷽ʽ("����")), "", rs���㷽ʽ("����"))
        mblnCancel = True
        cmb.ListIndex = rs���㷽ʽ!���� - 1
        mblnCancel = False
        chkDue.value = rs���㷽ʽ!Ӧ�տ�
        chkӦ����.value = IIF(Val(Nvl(rs���㷽ʽ!Ӧ����)) = 1, 1, 0)
        '75134:���ϴ�,2014/7/14,��ʽ�Ƿ�̶�
        mbln�̶� = IIF(rs���㷽ʽ("�Ƿ�̶�") = 1, True, False)
        txtEdit(2).Enabled = Not mbln�̶�
        txtEdit(3).Enabled = Not mbln�̶�
        cmb.Enabled = Not mbln�̶�
    Else
        txtEdit(1).Text = zlDatabase.GetMax("���㷽ʽ", "����", 2)
    End If
    '�������㳡��
    If rs���㷽ʽ.State = 1 Then rs���㷽ʽ.Close
    gstrSQL = "Select a.����,b.���㷽ʽ,b.ȱʡ��־ from ���㳡�� A,���㷽ʽӦ�� B" & vbNewLine & _
            " Where b.Ӧ�ó���(+)=a.���� and b.���㷽ʽ(+)= [1] And b.���ʽ(+) Is Null" & vbNewLine & _
            " Order by A.����"
    Set rs���㷽ʽ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
         
    lvw����.ListItems.Clear
    Do Until rs���㷽ʽ.EOF
        lvw����.ListItems.Add , "C" & rs���㷽ʽ("����"), rs���㷽ʽ("����")
        If rs���㷽ʽ("���㷽ʽ") = mstr���� Then
            lvw����.ListItems("C" & rs���㷽ʽ("����")).Checked = True
            lvw����.ListItems("C" & rs���㷽ʽ("����")).SubItems(1) = IIF(rs���㷽ʽ("ȱʡ��־") = 1, "ȱʡ", "")
        End If
        rs���㷽ʽ.MoveNext
    Loop
    chkDue.Enabled = CheckUsedDue
    If Not chkDue.Enabled Then chkDue.value = 0
    chkӦ����.Enabled = IsCheckDueValied
    If Not chkӦ����.Enabled Then chkӦ����.value = 0
    mblnChange = False
    frmBalanceEdit.Show vbModal
    �༭���㷽ʽ = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If Not mbln�̶� Then txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    Me.cmb.AddItem "1-�ֽ���㷽ʽ"
    Me.cmb.AddItem "2-������ҽ������"
    Me.cmb.AddItem "3-ҽ�������ʻ�"
    Me.cmb.AddItem "4-ҽ������ͳ��"
    Me.cmb.AddItem "5-���տ���"
    Me.cmb.AddItem "6-�����ۿ�"
    Me.cmb.AddItem "7-һ��ͨ����"
    Me.cmb.AddItem "8-���㿨����"
    mblnCancel = True
    If cmb.ListCount > 1 Then Me.cmb.ListIndex = 1
    chkDue.value = 0
    mblnCancel = False
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub lvw����_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    mblnItem = False
End Sub

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub ChangeServer()
    Dim ObjItem As ListItem
    
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    
    With lvw����.SelectedItem
        If .Checked = False Then Exit Sub
        If InStr("0,1,6,7", cmb.ListIndex) = 0 Then cmb_Click: Exit Sub 'ҽ�����㼰���տ����Ϊȱʡ��
        'Ӧ���������Ϊȱʡ�Ľ��㷽ʽ
        If chkӦ����.Enabled And chkӦ����.value = 1 Then cmb_Click: Exit Sub
        If .SubItems(1) = "" Then
            .SubItems(1) = "ȱʡ"
            mblnChange = True
        Else
            .SubItems(1) = ""
            mblnChange = True
        End If
    End With
End Sub

Private Sub lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    chkDue.Enabled = CheckUsedDue
    If Not chkDue.Enabled Then chkDue.value = 0
    
    chkӦ����.Enabled = IsCheckDueValied
    If Not chkӦ����.Enabled Then chkӦ����.value = 0
    
    '82990:���ϴ�,2015/3/9,ҽ���������ڲ�����
    If Item.Text = "������" Or Item.Text = "���ѿ�" Then
        If cmb.ListIndex <> 0 And cmb.ListIndex <> 1 And cmb.ListIndex <> 7 Then Item.Checked = False
    '���տ���ֻ��Ӧ����Ԥ����,�Ҳ���Ϊȱʡ
    ElseIf cmb.ListIndex = 4 Then
        If Item.Text <> "Ԥ����" Then
            Item.Checked = False
        End If
    ElseIf cmb.ListIndex = 6 Then   'һ��ͨ������Ԥ���;��￨
        If Item.Text = "Ԥ����" Or Item.Text = "���￨" Then Item.Checked = False
    ElseIf cmb.ListIndex = 7 Then
        '���㿨
        If InStr(",�շ�,����,Ԥ����,������,�Һ�,���￨,���ѿ�,", "," & Item.Text & ",") = 0 Then
            Item.Checked = False
        End If
    Else
        mblnChange = True
    End If
    If Item.Checked = False And Item.SubItems(1) = "ȱʡ" Then Item.SubItems(1) = ""
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub GetDefineSize()
'���ܣ��õ����ݿ�ı��ֶεĳ���
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "SELECT ����,����,���� FROM ���㷽ʽ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "���㷽ʽ�༭")
    
    txtEdit(1).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(2).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(3).MaxLength = rsTemp.Fields("����").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
