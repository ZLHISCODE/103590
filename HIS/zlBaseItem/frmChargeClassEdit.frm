VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeClassEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�շ���Ŀ�������"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmChargeClassEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3765
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Height          =   1410
      Left            =   165
      TabIndex        =   8
      Top             =   3885
      Width           =   3945
      Begin VB.CommandButton cmdDef 
         Cancel          =   -1  'True
         Caption         =   "�ָ�(&F)"
         Height          =   350
         Left            =   2715
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   2715
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   585
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   2715
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   210
         Width           =   1100
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "����"
         Top             =   990
         Width           =   1770
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "����"
         Top             =   225
         Width           =   1770
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "����"
         Top             =   615
         Width           =   1770
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   675
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4245
      TabIndex        =   7
      Tag             =   "����"
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4245
      TabIndex        =   12
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4245
      TabIndex        =   13
      Top             =   525
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
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
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":000C
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":0326
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":0778
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":0BCA
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":1022
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":1476
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":18CA
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":1D1E
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeClassEdit.frx":2B70
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
      End
   End
End
Attribute VB_Name = "frmChargeClassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnItemClick As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDef_Click()
    '�̶����Ŀ����޸ļ���,����ֻ�ָ�����
    If lvwMain.SelectedItem.ListSubItems(1).Tag <> 2 And lvwMain.SelectedItem.SubItems(2) <> lvwMain.SelectedItem.SubItems(6) Then
        lvwMain.SelectedItem.SubItems(2) = lvwMain.SelectedItem.SubItems(6)
        txtEdit(3).Text = lvwMain.SelectedItem.SubItems(2)
        cmdDef.Enabled = False
    End If
    
    If lvwMain.SelectedItem.ListSubItems(1).Tag <> 0 Then Exit Sub   '��Ϊ���޸������˳�
    
    lvwMain.SelectedItem.Text = lvwMain.SelectedItem.SubItems(4)
    lvwMain.SelectedItem.SubItems(1) = lvwMain.SelectedItem.SubItems(5)
    lvwMain.SelectedItem.SubItems(2) = lvwMain.SelectedItem.SubItems(6)
    
    lvwMain.SelectedItem.ListSubItems(2).Tag = 0
    txtEdit(1).Text = lvwMain.SelectedItem.Text
    txtEdit(2).Text = lvwMain.SelectedItem.SubItems(1)
    txtEdit(3).Text = lvwMain.SelectedItem.SubItems(2)
    cmdDef.Enabled = False
End Sub

Private Sub cmdDel_Click()
On Error GoTo ErrHandle
    Dim i As Long
    Dim objList As ListItem
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    lblEdit(1).Enabled = True
    lblEdit(2).Enabled = True
    lblEdit(3).Enabled = True
    txtEdit(1).Enabled = True
    txtEdit(2).Enabled = True
    txtEdit(3).Enabled = True
    
    If lvwMain.SelectedItem.ListSubItems(1).Tag = 1 Then
    'ϵͳ�̶����
        MsgBox "ϵͳ�̶������ɾ����", vbInformation, gstrSysName
        Exit Sub
    ElseIf lvwMain.SelectedItem.ListSubItems(1).Tag = 0 Then
    '�����޸ĵ����
        If lvwMain.SelectedItem.ListSubItems(2).Tag = 1 Or lvwMain.SelectedItem.ListSubItems(2).Tag = 0 Then
        'Ҫɾ���ı�Ϊ��ɫ
            lvwMain.SelectedItem.ForeColor = RGB(255, 0, 0)
            For i = 1 To lvwMain.ColumnHeaders.Count - 1
                lvwMain.SelectedItem.ListSubItems(i).ForeColor = RGB(255, 0, 0)
            Next
            lvwMain.SelectedItem.ListSubItems(2).Tag = 2
            cmdDel.Caption = "ȡ��ɾ��"
            lblEdit(1).Enabled = False
            lblEdit(2).Enabled = False
            lblEdit(3).Enabled = False
            txtEdit(1).Enabled = False
            txtEdit(2).Enabled = False
            txtEdit(3).Enabled = False
        Else
        'ȡ��ɾ������ǰ�ĺ�ɫ��Ϊ���ڵĺ�ɫ
            lvwMain.SelectedItem.ForeColor = 0
            For i = 1 To lvwMain.ColumnHeaders.Count - 1
                lvwMain.SelectedItem.ListSubItems(i).ForeColor = 0
            Next
            lvwMain.SelectedItem.ListSubItems(2).Tag = 1
            cmdDel.Caption = "ɾ��(&D)"
        End If
    ElseIf lvwMain.SelectedItem.ListSubItems(1).Tag = 2 Then
    '�����ӵ����
        i = lvwMain.SelectedItem.Index
        lvwMain.ListItems.Remove i
        On Error Resume Next
        Err.Clear
        Set objList = lvwMain.ListItems.Item(i - 1)
        If Err <> 0 Then
            If lvwMain.ListItems.Count > 0 Then
                lvwMain.ListItems(0).Selected = True
                lvwMain.ListItems(0).EnsureVisible
            End If
        Else
            objList.Selected = True
            objList.EnsureVisible
            lvwMain_ItemClick objList
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, "frmChargeClassEdit", Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save���() = False Then Exit Sub
    Unload Me
End Sub

Private Function IsValid() As Boolean
On Error GoTo ErrHandle
    Dim i As Integer, j As Long
    Dim strTemp As String
    '����Ƿ��зǷ��ַ�
    For j = 1 To lvwMain.ListItems.Count
        For i = 0 To 2
            If i = 0 Then
                strTemp = lvwMain.ListItems(j).Text
            Else
                strTemp = lvwMain.ListItems(j).SubItems(i)
            End If
            If zlCommFun.StrIsValid(strTemp, txtEdit(i + 1).MaxLength) = False Then
                lvwMain.ListItems(j).Selected = True
                lvwMain.ListItems(j).EnsureVisible
                lvwMain_ItemClick lvwMain.SelectedItem
                lvwMain.SetFocus
                Exit Function
            End If
        Next
    Next
    '������������
    For i = 1 To lvwMain.ListItems.Count
        If Trim(lvwMain.ListItems(i).Text) = "" Or Trim(lvwMain.ListItems(i).SubItems(1)) = "" Then
            lvwMain.ListItems(i).Selected = True
            lvwMain.ListItems(i).EnsureVisible
            lvwMain_ItemClick lvwMain.SelectedItem
            MsgBox "��������Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
            lvwMain.SetFocus
            Exit Function
        End If
        For j = 1 To lvwMain.ListItems.Count
            If Trim(lvwMain.ListItems(i).Text) = Trim(lvwMain.ListItems(j).Text) And i <> j Then
                lvwMain.ListItems(j).Selected = True
                lvwMain.ListItems(j).EnsureVisible
                lvwMain_ItemClick lvwMain.SelectedItem
                MsgBox "�����ظ���", vbExclamation, gstrSysName
                lvwMain.SetFocus
                Exit Function
            End If
            If Trim(lvwMain.ListItems(i).SubItems(1)) = Trim(lvwMain.ListItems(j).SubItems(1)) And i <> j Then
                lvwMain.ListItems(j).Selected = True
                lvwMain.ListItems(j).EnsureVisible
                lvwMain_ItemClick lvwMain.SelectedItem
                MsgBox "�����ظ���", vbExclamation, gstrSysName
                lvwMain.SetFocus
                Exit Function
            End If
        Next
    Next
    '��������û�иı��
    strTemp = ""
    For i = 1 To lvwMain.ListItems.Count
        If lvwMain.ListItems(i).ListSubItems(1).Tag = 0 Then
            If lvwMain.ListItems(i).Text <> lvwMain.ListItems(i).SubItems(4) Then
                strTemp = strTemp & "ԭ��Ŀ����" & lvwMain.ListItems(i).SubItems(4) & "��" & lvwMain.ListItems(i).SubItems(5) & vbTab & "�޸�Ϊ����" & lvwMain.ListItems(i).Text & "��" & lvwMain.ListItems(i).SubItems(2) & vbCrLf
            End If
        End If
    Next
    If strTemp <> "" Then
        If MsgBox(strTemp & vbCrLf & "�����Щ���ı����Ѿ�ʹ�ý����ܱ��棬�Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            lvwMain.SetFocus
            Exit Function
        End If
    End If
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save���() As Boolean
    Dim nod As Node
    Dim varTag As Variant
    Dim strCaption As String
    Dim i As Long
    
    On Error GoTo ErrHandle
    '��һ��ListSubItem��־ 0 = ��ǰ�еĿ��޸����     1 = ��ǰ�еĹ̶����    2 = �������
    '�ڶ���ListSubItem��־ 0 = ������κα仯         1 = �޸������          2 = ɾ��������
    gcnOracle.BeginTrans
    
    For i = 1 To lvwMain.ListItems.Count    '�̶�����Ҳ�����޸ļ���
        If lvwMain.ListItems(i).ListSubItems(1).Tag = 0 Or (lvwMain.ListItems(i).ListSubItems(1).Tag <> 2 And lvwMain.SelectedItem.SubItems(2) <> lvwMain.SelectedItem.SubItems(6)) Then
            '�����޸ĵ����
            If lvwMain.ListItems(i).ListSubItems(2).Tag = 1 Then
                'Update�Ǹ����
                gstrSQL = "zl_�շ����_update('" & Trim(lvwMain.ListItems(i).SubItems(4)) & "','" & Trim(lvwMain.ListItems(i).Text) & "','" & Trim(lvwMain.ListItems(i).SubItems(1)) & "','" & Trim(lvwMain.ListItems(i).SubItems(2)) & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            ElseIf lvwMain.ListItems(i).ListSubItems(2).Tag = 2 Then
                'Delete�Ǹ����
                gstrSQL = "ZL_�շ����_DELETE('" & Trim(lvwMain.ListItems(i).SubItems(4)) & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        ElseIf lvwMain.ListItems(i).ListSubItems(1).Tag = 2 Then
            '���������
                gstrSQL = "zl_�շ����_insert('" & Trim(lvwMain.ListItems(i).Text) & "','" & Trim(lvwMain.ListItems(i).SubItems(1)) & "','" & Trim(lvwMain.ListItems(i).SubItems(2)) & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    frmChargeManage.FillTree
    gcnOracle.CommitTrans
    
    Save��� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdAdd_Click()
On Error GoTo ErrHandle
Dim objList As ListItem
Dim i As Long
Dim blnOk As Boolean


    Set objList = lvwMain.ListItems.Add(, , "", "RootR", "RootR")
    objList.SubItems(1) = ""
    objList.SubItems(2) = ""
    objList.SubItems(3) = ""
    objList.SubItems(4) = ""
    objList.SubItems(5) = ""
    objList.SubItems(6) = ""
    '�ı�Ϊ2��Ϊ�½�
    objList.ListSubItems(1).Tag = 2
    '�ı�Ϊ0��Ϊ�ޱ仯
    objList.ListSubItems(2).Tag = 0
    objList.Selected = True
    objList.EnsureVisible
    lvwMain_ItemClick objList
    For i = 1 To lvwMain.ListItems.Count
        If UCase(sys.MaxCode("�շ���Ŀ���", "����", 1)) = UCase(lvwMain.ListItems(i).Text) Then
            
            blnOk = True
        End If
    Next
    If blnOk = True Then
        txtEdit(1).Text = ""
    Else
        txtEdit(1).Text = sys.MaxCode("�շ���Ŀ���", "����", 1)
    End If
    objList.Text = txtEdit(1).Text
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init()
'��ʼ������
Dim rsTmp As New ADODB.Recordset
Dim i As Long
Dim objList As ListItem
On Error GoTo ErrHandle

    lvwMain.ColumnHeaders.Clear
    zlControl.LvwSelectColumns lvwMain, "����,550,0,0;�������,1400,0,0;����,1000,0,0;�̶�,550,0,0;ԭ����,0,0,0;ԭ����,0,0,0;ԭ����,0,0,0", True
    
    gstrSQL = "Select nvl(����,'') ����,nvl(����,'') ����,nvl(����,'') ���� , decode(nvl(�̶�,0),0,'',' ��') �̶� From �շ���Ŀ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        '��ʼ���ؼ�
        lvwMain.ListItems.Clear
        rsTmp.MoveFirst
        '��һ��ListSubItem��־ 0 = ��ǰ�еĿ��޸����     1 = ��ǰ�еĹ̶����    2 = �������
        '�ڶ���ListSubItem��־ 0 = ������κα仯         1 = �޸������          2 = ɾ��������
        For i = 0 To rsTmp.RecordCount - 1
            Set objList = lvwMain.ListItems.Add(, "B" & rsTmp!���� & "_" & rsTmp!����, rsTmp!����, "Root", "Root")
            objList.SubItems(1) = rsTmp!����
            objList.SubItems(2) = rsTmp!����
            objList.SubItems(3) = Nvl(rsTmp!�̶�)
            objList.SubItems(4) = rsTmp!����
            objList.SubItems(5) = rsTmp!����
            objList.SubItems(6) = rsTmp!����
            
            If Trim(Nvl(rsTmp!�̶�)) = "" Then
            '�����޸ĵ�
                objList.ListSubItems(1).Tag = 0
            Else
            '�����޸ĵĹ̶����
                objList.ListSubItems(1).Tag = 1
            End If
            '��ʼʱͳһ�ı�־Ϊ0 ���κα仯
            objList.ListSubItems(2).Tag = 0
            rsTmp.MoveNext
        Next
        lvwMain.ListItems(1).Selected = True
        lvwMain.ListItems(1).EnsureVisible
        lvwMain_ItemClick lvwMain.SelectedItem
    Else
        lblEdit(1).Enabled = False
        lblEdit(2).Enabled = False
        lblEdit(3).Enabled = False
        txtEdit(1).Enabled = False
        txtEdit(2).Enabled = False
        txtEdit(3).Enabled = False
        cmdDel.Enabled = False
        cmdDef.Enabled = False
    End If
    If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = Nothing
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Init
    '�����б�Ϊƽ�����ͷ
    zlControl.LvwFlatColumnHeader lvwMain
    
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
blnItemClick = True
If lvwMain.SelectedItem Is Nothing Then Exit Sub
lblEdit(1).Enabled = True
lblEdit(2).Enabled = True
lblEdit(3).Enabled = True
txtEdit(1).Enabled = True
txtEdit(2).Enabled = True
txtEdit(3).Enabled = True

Me.cmdDel.Caption = "ɾ��(&D)"
Me.cmdDel.Enabled = True
Me.cmdDef.Enabled = False
If lvwMain.SelectedItem.ListSubItems(1).Tag = 1 Then
'ϵͳ�̶�
    lblEdit(1).Enabled = False
    lblEdit(2).Enabled = False
    lblEdit(3).Enabled = True
    txtEdit(1).Enabled = False
    txtEdit(2).Enabled = False
    txtEdit(3).Enabled = True
    cmdDel.Enabled = False
ElseIf lvwMain.SelectedItem.ListSubItems(1).Tag = 0 Then
'���޸ĵ�
    If lvwMain.SelectedItem.ListSubItems(2).Tag = 2 Then
        lblEdit(1).Enabled = False
        lblEdit(2).Enabled = False
        lblEdit(3).Enabled = False
        txtEdit(1).Enabled = False
        txtEdit(2).Enabled = False
        txtEdit(3).Enabled = False
        Me.cmdDel.Caption = "ȡ��ɾ��"
    End If
    If lvwMain.SelectedItem.Text <> lvwMain.SelectedItem.SubItems(4) Or lvwMain.SelectedItem.SubItems(1) <> lvwMain.SelectedItem.SubItems(5) Or lvwMain.SelectedItem.SubItems(2) <> lvwMain.SelectedItem.SubItems(6) Then
        Me.cmdDef.Enabled = True
    End If
End If
Me.txtEdit(1).Text = lvwMain.SelectedItem.Text
Me.txtEdit(2).Text = lvwMain.SelectedItem.SubItems(1)
Me.txtEdit(3).Text = lvwMain.SelectedItem.SubItems(2)
blnItemClick = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    '�����һ��˵�
    If Trim(lvwMain.SelectedItem.SubItems(3)) = "��" Then
        mnuEditDel.Caption = "ɾ��(&D)"
        mnuEditDel.Enabled = False
    Else
        mnuEditDel.Enabled = True
        '��һ��ListSubItem��־ 0 = ��ǰ�еĿ��޸����     1 = ��ǰ�еĹ̶����    2 = �������
        '�ڶ���ListSubItem��־ 0 = ������κα仯         1 = �޸������          2 = ɾ��������
        If lvwMain.SelectedItem.ListSubItems(1).Tag = 0 And lvwMain.SelectedItem.ListSubItems(2).Tag = 2 Then
            mnuEditDel.Caption = "ȡ��ɾ��"
        Else
            mnuEditDel.Caption = "ɾ��(&D)"
        End If
    End If
    PopupMenu mnuEdit
End If
End Sub

Private Sub mnuEditAdd_Click()
cmdAdd_Click
End Sub

Private Sub mnuEditDel_Click()
    If cmdDel.Enabled = True Then
        cmdDel_Click
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
If blnItemClick = True Then Exit Sub
    If lvwMain.SelectedItem Is Nothing Then Exit Sub    '��ѡ����Ŀʱ�˳�
    If Index = 1 Then
        lvwMain.SelectedItem.Text = Trim(Me.txtEdit(1).Text)
    ElseIf Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(Trim(txtEdit(2).Text))
        lvwMain.SelectedItem.SubItems(1) = Trim(Me.txtEdit(2).Text)
        lvwMain.SelectedItem.SubItems(2) = Trim(Me.txtEdit(3).Text)
    ElseIf Index = 3 Then
        lvwMain.SelectedItem.SubItems(2) = Trim(Me.txtEdit(3).Text)
    End If
        
    '�����ǿ����޸ĵĻ����½���
    '----------------------------------------------
    If lvwMain.SelectedItem.ListSubItems(2).Tag = 2 Then
    '�����ǰΪɾ��������Ϊɾ��
        lvwMain.SelectedItem.ListSubItems(2).Tag = 2
    Else
    '����Ϊ�޸�
        lvwMain.SelectedItem.ListSubItems(2).Tag = 1    '�̶����Ŀ����޸ļ���
        If lvwMain.SelectedItem.ListSubItems(1).Tag = 0 Or (lvwMain.SelectedItem.ListSubItems(1).Tag <> 2 And Trim(lvwMain.SelectedItem.SubItems(2)) <> Trim(lvwMain.SelectedItem.SubItems(6))) Then
        'ֻ�п��޸ĵĲŽ�����Ϊ�ָ�
            cmdDef.Enabled = True
            If Trim(lvwMain.SelectedItem.Text) = Trim(lvwMain.SelectedItem.SubItems(4)) Then
                If Trim(lvwMain.SelectedItem.SubItems(1)) = Trim(lvwMain.SelectedItem.SubItems(5)) Then
                    If Trim(lvwMain.SelectedItem.SubItems(2)) = Trim(lvwMain.SelectedItem.SubItems(6)) Then
                        cmdDef.Enabled = False
                        lvwMain.SelectedItem.ListSubItems(2).Tag = 0
                    End If
                End If
            End If
        End If
    End If
    '----------------------------------------------
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index <> 1 Then
        OS.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 3 Then KeyAscii = 0
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub



