VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRoleBat 
   Caption         =   "������ɫ����"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmPresRoleBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7710
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "����Ȩ��"
      Height          =   350
      Left            =   3600
      TabIndex        =   7
      Top             =   4680
      Width           =   900
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "����Ȩ��"
      Height          =   350
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   900
   End
   Begin VB.ListBox lstRole 
      Height          =   1680
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "���ң�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   4710
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   600
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "B.����"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   7800
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   120
      Top             =   7320
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
            Picture         =   "frmPresRoleBat.frx":6852
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":752C
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":DD8E
            Key             =   "YES"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1560
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":145F0
            Key             =   "YES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":1AE52
            Key             =   "NO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1725
      Left            =   4710
      TabIndex        =   9
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3043
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ϵͳ"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ģ��"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUnGrantedPres 
      Height          =   4005
      Left            =   600
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�û���"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSComctlLib.ListView lvwGrantedPres 
      Height          =   4005
      Left            =   4710
      TabIndex        =   11
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�û���"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label lblRole 
      Caption         =   "��ɫ"
      Height          =   255
      Left            =   98
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblUnGrantedPres 
      Caption         =   "δ����ý�ɫ����Ա"
      Height          =   315
      Left            =   570
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblGrantedPres 
      Caption         =   "������ý�ɫ����Ա"
      Height          =   195
      Left            =   4710
      TabIndex        =   15
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lblModule 
      Caption         =   "ģ���嵥"
      Height          =   255
      Left            =   3825
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����"
      Height          =   180
      Left            =   3825
      TabIndex        =   13
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   98
      TabIndex        =   12
      Top             =   660
      Width           =   360
   End
End
Attribute VB_Name = "frmPresRoleBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDeptId      As Long     '��ǰ����ID
Private mstrDeptName    As String   '��ǰ��������
Private mintOld         As Integer
Public Sub ShowMe(ByVal frmParent As Object, ByVal lngDept As Long, ByVal strDeptName As String)
    mlngDeptId = lngDept
    mstrDeptName = strDeptName
    Me.Show vbModal, frmParent
End Sub

Private Sub cboSystem_Click()
    FillModule
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGrant_Click()
    Dim i As Integer
    Dim strErr As String
    
    On Error Resume Next
    For i = 1 To lvwUnGrantedPres.ListItems.Count
        If lvwUnGrantedPres.ListItems(i).Checked Then
            gstrSQL = "Grant ZL_" & lstRole.Text & "  to " & lvwUnGrantedPres.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL, , adCmdText
            If Err <> 0 Then
                strErr = strErr & vbCrLf & Err.Description
                Err.Clear
            Else
                Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Add('" & lvwUnGrantedPres.ListItems(i).SubItems(1) & "','ZL_" & lstRole.Text & "')", Me.Caption)
            End If
        End If
    Next
    If strErr <> "" Then
        MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    Dim strErr As String
    On Error Resume Next
    
    For i = 1 To lvwGrantedPres.ListItems.Count
        If lvwGrantedPres.ListItems(i).Checked Then
            gstrSQL = "revoke ZL_" & lstRole.Text & " from " & lvwGrantedPres.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL, , adCmdText
            If Err <> 0 Then
                strErr = strErr & vbCrLf & Err.Description
                Err.Clear
            Else
                Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Del('" & lvwGrantedPres.ListItems(i).SubItems(1) & "','ZL_" & lstRole.Text & "')", Me.Caption)
            End If
        End If
    Next
    
    If strErr <> "" Then
        MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub Form_Load()
    lblName.Caption = "���ң�" & mstrDeptName
    Call FillRoleAndSystem
End Sub

Private Sub FillRoleAndSystem()
'���ص�ǰ��¼�û����й���Ȩ�޵Ľ�ɫ��ϵͳ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    Dim lstTmp As ListItem
    
    On Error GoTo errH
    strSQL = "Select Substr(Granted_Role, 4) ��ɫ" & vbNewLine & _
            "From Dba_Role_Privs" & vbNewLine & _
            "Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "And Admin_Option = 'YES'" & vbNewLine & _
            "And Grantee = User" & vbNewLine & _
            "Order By Substr(Granted_Role, 4)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-�û���ɫ")
    With lstRole
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!��ɫ
            rsTmp.MoveNext
        Next
    End With
    
    strSQL = "Select Distinct m.���, m.����, m.�����, m.������, m.��װ����, m.������װ, m.�汾��" & vbNewLine & _
            "From (Select Distinct ��ɫ, ϵͳ From Zlrolegrant Where ��� >= 100) r, Zlsystems m," & vbNewLine & _
            "     (Select Granted_Role" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And Admin_Option = 'YES'" & vbNewLine & _
            "       And Grantee = User) n" & vbNewLine & _
            "Where r.ϵͳ = m.���" & vbNewLine & _
            "And r.��ɫ = n.Granted_Role" & vbNewLine & _
            "Order By m.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-�Ѱ�װϵͳ")
    
    With cboSystem
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & " v" & rsTmp!�汾�� & "��" & rsTmp!��� & "��"
            .ItemData(cboSystem.NewIndex) = rsTmp!���
            If rsTmp!������ = UCase(gstrUserName) And .ListIndex < 0 Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        '������ϵͳ�ǳ���̶���
        If (zlRegTool And 2) = 2 Then .AddItem "�Զ��屨��"
        .AddItem "��������"
        .AddItem "ȡ������"
        .AddItem "��������"
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillModule()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lst As ListItem
    Dim strRole As String, strPre As String
    
    On Error GoTo errH
    lvwModule.ListItems.Clear
    '�����б���
    With lvwModule.ColumnHeaders
        .Clear
        If cboSystem.Text = "��������" Then
            .Add , , "�����", "1200"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "ȡ������" Then
            .Add , , "������", "1200"
            .Add , , "������", "1500"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "��������" Then
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
        Else
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
        End If
    End With
    If lstRole.ListIndex <> -1 Then
        strRole = "ZL_" & lstRole.List(lstRole.ListIndex)
    End If
    If strRole = "" Then Exit Sub
    If cboSystem.Text = "��������" Then '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        strSQL = "Select t.ϵͳ, t.����, t.˵��" & vbNewLine & _
                    "From (Select s.���� || '��' || s.��� || '��' As ϵͳ, s.������, b.����, b.˵��" & vbNewLine & _
                    "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
                    "       Where b.ϵͳ = s.���) t, User_Tab_Privs r" & vbNewLine & _
                    "Where t.������ = r.Owner" & vbNewLine & _
                    "And t.���� = r.Table_Name" & vbNewLine & _
                    "And r.Grantee =[1]" & vbNewLine & _
                    "And r.Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
                    "Group By t.ϵͳ, t.����, t.˵��" & vbNewLine & _
                    "Having Count(r.Privilege) = 4"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!����)
            lst.SubItems(1) = rsTmp!ϵͳ
            lst.SubItems(2) = rsTmp!˵�� & ""
            rsTmp.MoveNext
        Loop
    ElseIf cboSystem.Text = "ȡ������" Then '��ʾ�ý�ɫ�ܷ��ʵ�ȡ������
        strSQL = "Select s.���� || '��' || s.��� || '��' As ϵͳ, s.������, f.������, f.������, f.˵��" & vbNewLine & _
                "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
                "Where f.ϵͳ = s.���" & vbNewLine & _
                "And s.������ = r.Owner" & vbNewLine & _
                "And Upper(f.������) = r.Table_Name" & vbNewLine & _
                "And r.Grantee =[1]" & vbNewLine & _
                "And r.Privilege = 'EXECUTE'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!������)
            lst.SubItems(1) = rsTmp!������
            lst.SubItems(2) = rsTmp!ϵͳ
            lst.SubItems(3) = rsTmp!˵�� & ""
            rsTmp.MoveNext
        Loop
    Else
        If cboSystem.Text = "��������" Then '��ʾ�ý�ɫ�ܷ��ʵĻ�������
            strSQL = "Select p.���, p.����, p.˵��, r.����" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where r.ϵͳ Is Null" & vbNewLine & _
                    "And p.��� = r.���" & vbNewLine & _
                    "And r.��ɫ =[1]" & vbNewLine & _
                    "And p.ϵͳ Is Null" & vbNewLine & _
                    "And p.��� < 100" & vbNewLine & _
                    "Order By p.���"
            
        Else '��ʾ�ý�ɫ�ܷ��ʵ�ģ��
            strSQL = "Select p.���, p.����, p.˵��, r.����" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where Nvl(r.ϵͳ, 0) = Nvl(p.ϵͳ, 0)" & vbNewLine & _
                    "And p.��� = r.���" & vbNewLine & _
                    "And p.��� >= 100" & vbNewLine & _
                    "And r.��ɫ = [1] And " & vbNewLine & _
                    IIF(cboSystem.Text = "�Զ��屨��", " P.ϵͳ is null", " (P.ϵͳ=" & cboSystem.ItemData(cboSystem.ListIndex) & " OR P.��� Between 10000 And 19999)") & vbNewLine & _
                    "Order By p.���"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            If strPre <> rsTmp!��� & "" Then
                Set lst = lvwModule.ListItems.Add(, "K" & rsTmp!���, rsTmp!���)
                lst.SubItems(1) = rsTmp!����
                lst.SubItems(2) = rsTmp!˵�� & ""
                strPre = rsTmp!���
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lstRole_Click()
    FillModule
    LoadPreson
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intOld As Integer
    Dim blnEnd As Boolean
    
    If KeyAscii <> 13 Then Exit Sub
    If txtFind.Text = "" Then Exit Sub
    
    zlControl.TxtSelAll txtFind
    
    If txtFind.Tag <> txtFind.Text Then
        mintOld = 0
        txtFind.Tag = txtFind.Text
    Else
        If mintOld + 1 >= lstRole.ListCount Then
            mintOld = 0
            txtFind.Tag = ""
        End If
    End If
    
RowX:
    For i = mintOld To lstRole.ListCount
        If InStr(1, lstRole.List(i), txtFind.Text) > 0 Or InStr(1, zlStr.GetCodeByVB(lstRole.List(i)), UCase(txtFind.Text)) > 0 Then
            lstRole.Selected(i) = True
            mintOld = i + 1
            blnEnd = True
            Exit Sub
        End If
    Next
    
    If Not blnEnd And mintOld <> 0 Then
        mintOld = 0
        GoTo RowX
    End If
End Sub

Private Sub LoadPreson()
'������Ա
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lstTmp As ListItem
    
    On Error GoTo errH
    strSQL = "Select a.�û���, d.����, Decode(e.Granted_Role, Null, 0, 1) Ȩ��" & vbNewLine & _
            "From �ϻ���Ա�� a, ������Ա b, ���ű� c, ��Ա�� d, (Select Grantee, Granted_Role From Dba_Role_Privs Where Granted_Role = [2]) e" & vbNewLine & _
            "Where a.��Աid = b.��Աid" & vbNewLine & _
            "And b.����id = c.Id" & vbNewLine & _
            "And d.Id = a.��Աid" & vbNewLine & _
            "And a.�û��� = e.Grantee(+)" & vbNewLine & _
            "And b.����id = [1]" & vbNewLine & _
            "And a.�û��� <> [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-�û���ɫ", mlngDeptId, "ZL_" & lstRole.Text, UCase(gstrDbUser))
    
    lvwGrantedPres.ListItems.Clear
    lvwUnGrantedPres.ListItems.Clear
    Do While Not rsTmp.EOF
        If rsTmp!Ȩ�� = 1 Then
            Set lstTmp = lvwGrantedPres.ListItems.Add(, rsTmp!�û���, rsTmp!����, , "YES")
            lstTmp.SubItems(1) = rsTmp!�û���
        Else
            Set lstTmp = lvwUnGrantedPres.ListItems.Add(, rsTmp!�û���, rsTmp!����, , "NO")
            lstTmp.SubItems(1) = rsTmp!�û���
        End If
        lstTmp.Checked = True
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

