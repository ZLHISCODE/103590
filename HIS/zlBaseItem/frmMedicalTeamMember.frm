VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalTeamMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��С���Ա"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmMedicalTeamMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraView 
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   4335
      Begin VB.TextBox txtMember 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cboTeam 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtTeam 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblMember 
         AutoSize        =   -1  'True
         Caption         =   "С���Ա(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         Caption         =   "ת��С��(&T)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1250
         Width           =   990
      End
      Begin VB.Label lblFromTeam 
         AutoSize        =   -1  'True
         Caption         =   "����С��(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   770
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   1
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2280
      TabIndex        =   0
      Top             =   4950
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   4683
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgTvw"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   3120
      Top             =   2160
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
            Picture         =   "frmMedicalTeamMember.frx":000C
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeamMember.frx":0326
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeamMember.frx":0640
            Key             =   "Role"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMedicalTeamMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytStatus As Byte
Private mlngDeptID As Long, mlngTeamID As Long, mlngMemberID As Long
Public mblnOK As Boolean
Public mstrPrivs As String

Property Get Status() As Byte
'״̬ 1-��ӳ�Ա; 2-תС��
    Status = mbytStatus
End Property
Property Let Status(ByVal bytStatus As Byte)
    Caption = "ҽ��С���Ա"
    '�������
    If bytStatus = 1 Then
        Caption = Caption & "-���"
    Else
        Caption = Caption & "-תС��"
    End If
    
    'ˢ��rpcView�ؼ�
    RefreshViewRPC bytStatus
    mbytStatus = bytStatus
End Property

Public Sub ShowMe(ByVal frmVal As Form, ByVal bytStatus As Byte, ByVal lngDeptID As Long, _
ByVal lngTeamID As Long, Optional ByVal lngMemberID As Long)
    Dim rsTmp As ADODB.Recordset
    Dim nodTmp As Node, nodParent As Node
    mlngDeptID = lngDeptID
    mlngTeamID = lngTeamID
    mlngMemberID = lngMemberID
    Status = bytStatus
    
    On Error GoTo ErrHandle
    If Status = 1 Then
        gstrSQL = "select a.ID,a.���� from ���ű� a, �ٴ�ҽ��С�� b where a.id=b.����id and b.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTeamID)
        If rsTmp.RecordCount = 1 Then
            '���Ҷ�λ
            For Each nodTmp In tvwList.Nodes
                If Val(Mid(nodTmp.Key, 2)) = rsTmp!ID Then
                    Set nodParent = nodTmp
                    nodTmp.Expanded = True
                    nodTmp.Selected = True
                    Do Until nodParent Is Nothing
                        Set nodParent = nodParent.Parent
                        If Not nodParent Is Nothing Then
                            Set nodTmp = nodParent
                            nodTmp.Expanded = True
                        End If
                    Loop
                    Exit For
                End If
            Next
        End If
    End If
    Show vbModal, frmVal
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Status = 1 Then
        If SaveMember() Then
            Unload Me
        Else
            MsgBox "δ��ѡ��Ա��", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If Me.cboTeam.ListIndex < 0 Then
            MsgBox "δѡ��ҽ��С�飡", vbInformation, gstrSysName
            Exit Sub
        End If
        If TransMember() Then
            Unload Me
'        Else
'            MsgBox "δѡ����Ա��", vbInformation, gstrSysName
'            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    With tvwList
        .ImageList = Me.ImgСͼ��
        .LabelEdit = tvwManual
    End With
    mblnOK = False
End Sub

Private Sub Form_Resize()
    If Status = 1 Then
        With tvwList
            .Visible = True
            .Top = 0
            .Left = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - Me.cmdOK.Height - 200
        End With
        fraView.Visible = False
    Else
        With fraView
            .Visible = True
            .Top = 10
            .Left = 100
            .Width = Me.ScaleWidth - 200
        End With
        cmdOK.Top = fraView.Top + fraView.Height + 100
        cmdCancel.Top = cmdOK.Top
        Top = Top + (Height - (cmdOK.Top + cmdOK.Height + 600)) / 2
        Height = cmdOK.Top + cmdOK.Height + 600
        tvwList.Visible = False
    End If
End Sub

Private Sub RefreshViewRPC(ByVal bytStatus As Byte)
    Dim i As Long, lngChars As Long
    Dim objNode As Node
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If bytStatus = 1 Then
        '��ӳ�Ա
        gstrSQL = "select id,����,����,�ϼ�id From ���ű� where ����<>'-'" & _
                  "  and (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) " & _
                  "  and id=[1] " & _
                  "start with �ϼ�id is null connect by prior id=�ϼ�id "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptID)
        With tvwList
            .Nodes.Clear
            Do While Not rsTmp.EOF
'                If IsNull(rsTmp("�ϼ�id")) Then
                    Set objNode = .Nodes.Add(, , "D" & rsTmp("id"), "��" & rsTmp("����") & "��" & rsTmp("����"), "Dept", "Dept")
'                Else
'                    Set objNode = .Nodes.Add("D" & rsTmp("�ϼ�id"), tvwChild, "D" & rsTmp("id").Value, "��" & rsTmp("����") & "��" & rsTmp("����"), "Dept", "Dept")
'                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
'            If InStr(mstrPrivs, "���п���") = 0 Then
                gstrSQL = "Select a.Id, a.���, a.����, b.����id, (select 1 from ҽ��С����Ա where С��id=[1] and ��Աid=a.id) С���Ա" & vbNewLine & _
                          "From ��Ա�� A, ���ű� C, ������Ա B,��������˵�� D,��Ա����˵�� E " & vbNewLine & _
                          "Where (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And a.Id = b.��Աid And b.����ID=[2] " & vbNewLine & _
                          " And b.����id = c.Id And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                          " And a.ID=e.��ԱID and e.��Ա����='ҽ��' and c.ID=d.����ID and d.��������='�ٴ�' and d.������� in (2,3) " & vbNewLine & _
                          "Order by a.��� "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngDeptID)
'            Else
'                gstrSQL = "Select a.Id, a.���, a.����, b.����id, (select 1 from ҽ��С����Ա where С��id=[1] and ��Աid=a.id) С���Ա" & vbNewLine & _
'                          "From ��Ա�� A, ���ű� C, ������Ա B,��������˵�� D,��Ա����˵�� E " & vbNewLine & _
'                          "Where (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And a.Id = b.��Աid And b.ȱʡ = 1 And b.����id = c.Id And" & vbNewLine & _
'                          "  (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
'                          " and a.ID=e.��ԱID and e.��Ա����='ҽ��' and c.ID=d.����ID and d.��������='�ٴ�' and d.������� in (2,3) " & vbNewLine & _
'                          "Order by a.��� "
'                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID)
'            End If
'            gstrSQL = "Select a.Id, a.���, a.����, b.����id, (Select 1 From ҽ��С����Ա Where С��id = [1] And ��Աid = a.Id) С���Ա " & vbNewLine & _
'                      "From ��Ա�� A, ���ű� C, ������Ա B, ��������˵�� D, ��Ա����˵�� E, ������Ա F " & vbNewLine & _
'                      "Where (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And a.Id = f.��Աid And b.����id = c.Id And " & vbNewLine & _
'                      "      b.����id = f.����id And b.��Աid = [2] And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And " & vbNewLine & _
'                      "      a.Id = e.��Աid And e.��Ա���� = 'ҽ��' And c.Id = d.����id And d.�������� = '�ٴ�' And d.������� In (2, 3) " & vbNewLine & _
'                      "Order By a.��� "
'            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, glngUserId)
            Do Until rsTmp.EOF
                Set objNode = .Nodes.Add("D" & rsTmp("����id"), 4, "P" & rsTmp("id"), "��" & rsTmp("���") & "��" & rsTmp("����"), "User", "User")
                objNode.ForeColor = RGB(0, 0, 255)
                If rsTmp!С���Ա = 1 Then
                    objNode.Checked = True
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
        End With
    Else
        'תС��
        gstrSQL = "select ���,���� from ��Ա�� where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngMemberID)
        If rsTmp.RecordCount > 0 Then
            txtMember.Text = "��" & rsTmp!��� & "��" & rsTmp!����
        End If
        rsTmp.Close
        gstrSQL = "Select ���� From �ٴ�ҽ��С�� Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID)
        If rsTmp.RecordCount > 0 Then
            txtTeam.Text = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            txtTeam.Tag = mlngTeamID
            rsTmp.Close
            cboTeam.Clear
            If InStr(mstrPrivs, "���п���") = 0 Then
                gstrSQL = "Select a.ID, a.����, a.����ID, c.���� ���� From �ٴ�ҽ��С�� a, ������Ա b, ���ű� c " & _
                          "Where a.ID <> [1] and a.����ID=b.����ID and b.��ԱID=[2] and b.����id=c.id and substr(a.����,1,1)<>'-' " & _
                          "  and not a.ID in (select С��id from ҽ��С����Ա where ��Աid=[3]) " & _
                          "  and a.����ʱ��=to_date('3000-1-1', 'YYYY-MM-DD') order by a.����ID,a.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, glngUserId, mlngMemberID)
            Else
                gstrSQL = "Select a.ID, a.����, a.����ID, b.���� ���� From �ٴ�ҽ��С�� a, ���ű� b " & _
                          "Where a.ID <> [1] and substr(a.����,1,1)<>'-' " & _
                          "  and not a.ID in (select С��id from ҽ��С����Ա where ��Աid=[2]) " & _
                          "  and a.����ʱ��=to_date('3000-1-1', 'YYYY-MM-DD') and a.����id=b.id order by a.����ID,a.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngMemberID)
            End If
            For i = 0 To rsTmp.RecordCount - 1
                cboTeam.AddItem IIF(IsNull(rsTmp!����), "", rsTmp!����) & " | " & IIF(IsNull(rsTmp!����), "", rsTmp!����)
                cboTeam.ItemData(i) = rsTmp!ID
                If rsTmp!����ID = mlngDeptID And cboTeam.ListIndex < 0 Then cboTeam.ListIndex = i
                If Len(rsTmp!���� & rsTmp!����) > lngChars Then lngChars = Len(rsTmp!���� & rsTmp!����)
                rsTmp.MoveNext
            Next
            zlControl.CboSetWidth cboTeam.hwnd, (lngChars + 2) * 15 * 6.5
            If cboTeam.ListIndex < 0 And cboTeam.ListCount > 0 Then cboTeam.ListIndex = 0
        End If
        rsTmp.Close
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwList_Click()
    Dim objNode As Node
    Set objNode = tvwList.SelectedItem
    If Left(objNode.Key, 1) = "D" Then objNode.Checked = False
End Sub

Private Sub tvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Left(tvwList.SelectedItem.Key, 1) = "D" Then
            tvwList.SelectedItem.Checked = False
            KeyCode = 0
        End If
    End If
End Sub

Private Sub tvwList_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Left(Node.Key, 1) = "D" Then
        Node.Checked = False
    End If
End Sub

Private Function SaveMember() As Boolean
    Dim objNode As Node
    Dim strMemberIDs As String
    For Each objNode In tvwList.Nodes
        If objNode.Checked Then
            strMemberIDs = strMemberIDs & Mid(objNode.Key, 2, 20) & ";"
        End If
    Next
    If strMemberIDs = "" Then Exit Function
    
    On Error GoTo ErrHandle
    gstrSQL = "ZL_ҽ��С����Ա_INSERT(" & mlngTeamID & ",'" & Left(strMemberIDs, Len(strMemberIDs) - 1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveMember = True
    mblnOK = True
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function TransMember() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strMess As String
    Dim i As Long
    On Error GoTo ErrHandle
    '���סԺҽʦ���в��˾���ʾ
'    gstrSQL = "Select a.����id, a.סԺ��, a.��Ժ����, b.����" & vbNewLine & _
'              "From ������ҳ a, ������Ϣ b " & vbNewLine & _
'              "Where a.סԺҽʦ = (Select ����" & vbNewLine & _
'              "              From ��Ա��" & vbNewLine & _
'              "              Where ID = [2] And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)) And" & vbNewLine & _
'              "      a.ҽ��С��id = [1] and a.����id=b.����id and b.��Ժ=1 "
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTeam.Tag, mlngMemberID)
'    With rsTmp
'        For i = 1 To .RecordCount
'            strMess = strMess & "������" & !���� & "��" & vbTab & _
'                      "סԺ�ţ�" & IIF(IsNull(!סԺ��), "", !סԺ��) & "��" & vbTab & _
'                      "���ţ�" & IIF(IsNull(!��Ժ����), "", !��Ժ����) & vbTab & vbNewLine
'            .MoveNext
'        Next
'    End With
    strMess = MedicalTeamPatients(Val(txtTeam.Tag), mlngMemberID)
    If strMess <> "" Then
        If MsgBox("��ҽ����ǰ��������Ժ���ˣ�" & vbNewLine & vbNewLine & strMess & vbNewLine & "ȷ�����ϲ��˵�ҽ��С��Ҳ��һ��ת����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    '�ж�����С���Ƿ����
    gstrSQL = "select count(*) rec from ҽ��С����Ա where С��ID=[1] and ��ԱID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngMemberID)
    If rsTmp!rec = 1 Then
        'תС��
        gstrSQL = "Zl_ҽ��С����Ա_Update("
        gstrSQL = gstrSQL & txtTeam.Tag & ","                            '����С��ID
        gstrSQL = gstrSQL & mlngMemberID & ","                           '��ԱID
        gstrSQL = gstrSQL & cboTeam.ItemData(cboTeam.ListIndex) & ",'"   'ת��С��ID
        gstrSQL = gstrSQL & gstrUserCode & "','"                         '����Ա���
        gstrSQL = gstrSQL & gstrUserName & "')"                          '����Ա����
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        TransMember = True
        mblnOK = True
    Else
        TransMember = True
        mblnOK = True
        MsgBox "��ҽ���Ѿ��������û��Ƴ���", vbInformation, gstrSysName
    End If
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

