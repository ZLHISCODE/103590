VERSION 5.00
Begin VB.Form frmPresFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ա����"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmPresFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraƥ�� 
      Caption         =   "ƥ�䷽ʽ"
      Height          =   1515
      Left            =   3030
      TabIndex        =   14
      Top             =   120
      Width           =   1500
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   450
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   5160
      TabIndex        =   12
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   5160
      TabIndex        =   11
      Top             =   180
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   2685
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2760
      Begin VB.ComboBox cmb�Ա� 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   330
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   870
         MaxLength       =   255
         TabIndex        =   3
         Top             =   720
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1110
         Width           =   1725
      End
      Begin VB.ComboBox cmbѧ�� 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1890
         Width           =   1755
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��(&D)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "���(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.Label lbl��� 
      BackStyle       =   0  'Transparent
      Caption         =   " �������������"
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   3000
      TabIndex        =   10
      Top             =   1920
      Width           =   3315
   End
End
Attribute VB_Name = "frmPresFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintfun As Integer  '0-��Ա����,1-���Ų���
Private mbln�Ƿ���ʾͣ�� As Boolean
Private mblnViewDel As Boolean
Dim mrsFind As New ADODB.Recordset
Private mintģʽ As Integer '1-�������ʾ��2-��������ʾ


Private Sub cmb�Ա�_Click()
    If mrsFind.State = 1 Then mrsFind.Close
    lbl���.Caption = "  �����Ѹı䣬�����¶�λ"
    lbl���.ForeColor = &H8000&
    If Not cmdFind.Enabled Then cmdFind.Enabled = True
End Sub

Private Sub cmb�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbѧ��_Click()
    If mrsFind.State = 1 Then mrsFind.Close
    lbl���.Caption = "  �����Ѹı䣬�����¶�λ"
    lbl���.ForeColor = &H8000&
    If Not cmdFind.Enabled Then cmdFind.Enabled = True
End Sub

Private Sub cmbѧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    Select Case mintfun
    Case 0   'Ĭ�ϰ���Ա����
        Dim rsTemp As New ADODB.Recordset
        
        cmb�Ա�.AddItem " "
        gstrSQL = "Select '�Ա�' As ���, ����,���� From �Ա� Union All Select 'ѧ��' As ���, ����,���� From ѧ�� Order By ���,����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        rsTemp.Filter = "���='�Ա�'"
        Do Until rsTemp.EOF
            cmb�Ա�.AddItem rsTemp("����")
            rsTemp.MoveNext
        Loop
        
        cmbѧ��.AddItem " "
        rsTemp.Filter = "���='ѧ��'"
        Do Until rsTemp.EOF
            cmbѧ��.AddItem rsTemp("����")
            rsTemp.MoveNext
        Loop
        
        rsTemp.Close
        cmdFind.Enabled = False
    Case 1
        frmPresFind.Caption = "���Ų���"
        cmbѧ��.Visible = False
        cmb�Ա�.Visible = False
        lbl(1).Caption = "����"
        lbl(3).Visible = False
        lbl(4).Visible = False
        fra����.Height = fraƥ��.Height
        lbl���.Left = fra����.Left
        lbl���.Width = fraƥ��.Left - fra����.Left + fraƥ��.Width
        'lbl���.Height = lbl���.Height / 2
        'Me.Height = Me.Height - lbl���.Height
        
        cmdFind.Enabled = False
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowOfType(frmParent As Object, intType As Integer, Optional blnShowStop As Boolean = False, Optional blnShowDel As Boolean = False, Optional intģʽ As Integer)
    mintfun = intType
    mbln�Ƿ���ʾͣ�� = blnShowStop
    mblnViewDel = blnShowDel
    mintģʽ = intģʽ
    
    frmPresFind.Show vbModal, frmParent
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrHandle
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    
    If txtEdit(0).Text <> "" Then
        gstrSQL = "and upper(" & Choose(mintfun + 1, "A.���", "a.����") & ") like [1]  "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & " and upper(" & Choose(mintfun + 1, "A.����", "a.����") & ") like [2] "
    End If
    
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "and upper(A.����) like [3]  "
    End If
    
    If mintfun = 0 Then
        If Trim(cmb�Ա�.Text) <> "" Then
            gstrSQL = gstrSQL & "and A.�Ա�=[4] "
        End If
        If Trim(cmbѧ��.Text) <> "" Then
            gstrSQL = gstrSQL & "and A.ѧ��=[5] "
        End If
    End If
    
    If gstrSQL = "" Then
'        gstrSQL = Mid(gstrSQL, 1, Len(gstrSQL) - 4)
'    Else
        MsgBox "���������������", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
        
    Select Case mintfun
    Case 0  '������Ա
        If InStr(frmPresManage.mstrPrivs, "���в���") = 0 Then
            gstrSQL = "Select a.Id, a.����, b.����id" & vbNewLine & _
                      "From ��Ա�� A, ������Ա B " & vbNewLine & _
                      "Where a.Id = b.��Աid And b.����id In (Select Distinct ID" & vbNewLine & _
                      "    From ���ű� A" & vbNewLine & _
                      "    Start With ID In (Select ����id From ������Ա Where ��Աid = [6])" & vbNewLine & _
                      "    Connect By Prior ID = �ϼ�id) And b.ȱʡ = 1 " & gstrSQL
        Else
            gstrSQL = "select A.ID,A.����,B.����ID " & _
                       " from ��Ա�� A,������Ա B " & _
                       " where A.ID =B.��ԱID and B.ȱʡ=1  " & gstrSQL
        End If
        If Not mbln�Ƿ���ʾͣ�� Then
            gstrSQL = gstrSQL & " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
        End If
    Case 1 '���Ҳ���
        gstrSQL = "Select a.id,a.�ϼ�id,a.����,a.���� ,c.���� as ���� From ���ű� A, ��������˵�� B,�������ʷ��� c Where b.��������=c.���� " & _
                " and A.ID=B.����ID " & gstrSQL
        
        If Not mbln�Ƿ���ʾͣ�� Then
            gstrSQL = gstrSQL & " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
        End If
        If Not mblnViewDel Then
            gstrSQL = gstrSQL & " and substr(a.����,1,1)<>'-' "
        End If
    End Select
    
    Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%", _
        IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(1).Text) & "%", _
        IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%", _
        cmb�Ա�.Text, cmbѧ��.Text, glngUserId)

    If mrsFind.State = 1 Then
        Call LocateItem
    End If
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    If mrsFind.RecordCount = 0 Then
        lbl���.Caption = " û���ҵ�������������Ϣ!"
        lbl���.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        lbl���.Caption = " �Ѿ���λ�������ҵ�����Ϣ����������������"
        lbl���.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    lbl���.Caption = "  �ҵ�" & mrsFind.RecordCount & "��������������Ϣ��" & vbCrLf & "��ǰ�ǵ�" & mrsFind.AbsolutePosition & _
                    "����" & Choose(mintfun + 1, "������", "���ƣ�") & mrsFind(Choose(mintfun + 1, "����", "����"))
    lbl���.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "��һ��(&L)"
        Else
            cmdFind.Caption = "��λ(&L)"
            cmdFind.Enabled = False
            lbl���.Caption = lbl���.Caption & vbCrLf & "�Ѿ���λ�����һ����Ϣ����������������"
        End If
    End If
    
    Select Case mintfun
    Case 0  '������Ա
        With frmPresManage.tvwMain_S
            .Nodes("C" & mrsFind("����ID")).Selected = True
            .SelectedItem.EnsureVisible
            frmPresManage.FillList "C" & mrsFind("����ID")
        End With
            
        With frmPresManage.lvwMain
            .ListItems("C" & mrsFind("ID")).Selected = True
            .SelectedItem.EnsureVisible
            frmPresManage.lvwMain_ItemClick .SelectedItem
        End With
    Case 1 '���Ҳ���
        With frmDeptManage.tvwMain_S
            If IsNull(mrsFind("�ϼ�ID")) Then
                .Nodes("C" & mrsFind("ID")).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            Else
                If mintģʽ = 1 Then
                    .Nodes("C" & mrsFind("�ϼ�ID")).Selected = True
                    .Nodes("C" & mrsFind("�ϼ�ID")).Expanded = True
                Else
                    strTemp = mrsFind!���� & "|" & mrsFind!ID
                    .Nodes("C" & strTemp).Selected = True
                    .Nodes("C" & strTemp).Expanded = True
                End If
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
                
                If mintģʽ = 1 Then
                    frmDeptManage.lvwMain.ListItems("C" & mrsFind("ID")).Selected = True
                    frmDeptManage.lvwMain.SelectedItem.EnsureVisible
                    frmDeptManage.lvwMain_ItemClick frmDeptManage.lvwMain.SelectedItem
                End If
            End If
        End With
    End Select
End Sub

Private Function IsValid() As Boolean
'����:��������������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 2
        strTemp = Trim(txtEdit(i).Text)
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    IsValid = True
End Function

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl���.Caption = "  �����Ѹı䣬�����¶�λ"
    cmdFind.Enabled = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdFind.SetFocus
        Call cmdFind_Click
'          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        OS.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub
